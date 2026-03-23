using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.AspNetCore.Mvc;
using Spire.Pdf.Conversion;
using System.IO.Compression;
using System.Xml.Linq;
using ITextRect = iText.Kernel.Geom.Rectangle;

namespace Ofd2PdfService.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ConvertController : ControllerBase
{
    /// <summary>
    /// Maximum number of OFD page batches to convert concurrently per request.
    /// Capped to avoid saturating the CPU when multiple requests run in parallel.
    /// </summary>
    private const int MaxParallelConversions = 4;

    private readonly ILogger<ConvertController> _logger;

    public ConvertController(ILogger<ConvertController> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Convert an OFD file to PDF.
    /// </summary>
    /// <remarks>
    /// POST /api/convert
    /// Content-Type: multipart/form-data
    /// Body: file=&lt;OFD file&gt;
    ///
    /// Returns the converted PDF as a file download.
    /// </remarks>
    [HttpPost]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> Post([FromForm] IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest(new ProblemDetails
            {
                Title = "No file provided",
                Detail = "Please upload an OFD file using the 'file' form field.",
                Status = StatusCodes.Status400BadRequest
            });
        }

        var extension = Path.GetExtension(file.FileName);
        if (!string.Equals(extension, ".ofd", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new ProblemDetails
            {
                Title = "Invalid file type",
                Detail = $"Expected an .ofd file but received '{extension}'.",
                Status = StatusCodes.Status400BadRequest
            });
        }

        var tmpDir = Path.Combine(Path.GetTempPath(), "ofd2pdf", Guid.NewGuid().ToString());
        Directory.CreateDirectory(tmpDir);

        // Use a safe, generated filename to avoid path traversal from user-supplied names
        var safeInputName = Path.GetRandomFileName() + ".ofd";
        var inputPath = Path.Combine(tmpDir, safeInputName);
        var outputPath = Path.ChangeExtension(inputPath, ".pdf");

        try
        {
            using (var stream = new FileStream(inputPath, FileMode.Create, FileAccess.Write))
            {
                await file.CopyToAsync(stream);
            }

            _logger.LogInformation("Converting {FileName} to PDF", file.FileName);
            // Offload CPU-bound conversion work to a thread-pool thread so that the
            // ASP.NET Core request thread is not blocked for the duration of the
            // potentially long-running OFD→PDF transformation.
            await Task.Run(() =>
            {
                ConvertOfdToPdf(inputPath, outputPath);
                RemoveEvaluationWarning(outputPath, _logger);
            });

            var pdfBytes = await System.IO.File.ReadAllBytesAsync(outputPath);
            var pdfFileName = Path.ChangeExtension(Path.GetFileName(file.FileName), ".pdf");

            _logger.LogInformation("Conversion succeeded: {FileName}", pdfFileName);
            return File(pdfBytes, "application/pdf", pdfFileName);
        }
        catch (AggregateException agex)
        {
            // Parallel.ForEach (used inside ConvertOfdToPdf) throws AggregateException.
            // Flatten and re-throw the first real cause for a cleaner error message.
            var inner = agex.Flatten().InnerException ?? agex;
            _logger.LogError(inner, "Conversion failed for {FileName}", file.FileName);
            return StatusCode(StatusCodes.Status500InternalServerError, new ProblemDetails
            {
                Title = "Conversion failed",
                Detail = inner.Message,
                Status = StatusCodes.Status500InternalServerError
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Conversion failed for {FileName}", file.FileName);
            return StatusCode(StatusCodes.Status500InternalServerError, new ProblemDetails
            {
                Title = "Conversion failed",
                Detail = ex.Message,
                Status = StatusCodes.Status500InternalServerError
            });
        }
        finally
        {
            try { Directory.Delete(tmpDir, recursive: true); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to clean up temp directory {TmpDir}", tmpDir); }
        }
    }

    /// <summary>
    /// Converts an OFD file to PDF, splitting it into batches of up to 3 pages to work
    /// around the FreeSpire.PDF evaluation version page limit, then merging the results.
    /// A sanitized copy of the OFD is always created via <see cref="CreateSubOfd"/> before
    /// conversion so that raw CFF fonts stored with a misleading .otf or .ttf extension are
    /// renamed to .cff — preventing a NullReferenceException inside Spire.PDF's OTF/TrueType
    /// font parser.
    /// </summary>
    private static void ConvertOfdToPdf(string inputPath, string outputPath)
    {
        int totalPages;
        try { totalPages = GetOfdPageCount(inputPath); }
        catch { totalPages = 0; }

        if (totalPages <= 0)
        {
            // Unknown page count – fall back to direct conversion without sanitization.
            new OfdConverter(inputPath).ToPdf(outputPath);
            return;
        }

        var tmpDir = Path.Combine(Path.GetTempPath(), "ofd2pdf_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmpDir);
        try
        {
            if (totalPages <= 3)
            {
                // Always create a sanitized copy (font renaming, etc.) even for small
                // documents so that the same CFF font fix applies regardless of page count.
                var sanitizedOfd = Path.Combine(tmpDir, "sanitized.ofd");
                CreateSubOfd(inputPath, sanitizedOfd, 0, totalPages);
                new OfdConverter(sanitizedOfd).ToPdf(outputPath);
                return;
            }

            // Precompute all batch descriptors so the two parallel phases below
            // can work over the same ordered list and the final merge keeps pages
            // in the correct sequence.
            var batchParams = new List<(string PartOfd, string PartPdf, int Start, int Count)>();
            for (int start = 0; start < totalPages; start += 3)
            {
                batchParams.Add((
                    Path.Combine(tmpDir, $"part_{start}.ofd"),
                    Path.Combine(tmpDir, $"part_{start}.pdf"),
                    start,
                    Math.Min(3, totalPages - start)));
            }

            // Phase 1: Create all sub-OFD files in parallel.
            // Each CreateSubOfd call opens its own ZipArchive (ZipFile.OpenRead),
            // so concurrent reads from the same source file are safe.
            Parallel.ForEach(batchParams, batch =>
                CreateSubOfd(inputPath, batch.PartOfd, batch.Start, batch.Count));

            // Phase 2: Convert sub-OFDs to PDFs in parallel, capping concurrency to
            // avoid overwhelming the system when multiple requests run simultaneously.
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Math.Min(Environment.ProcessorCount, MaxParallelConversions)
            };
            Parallel.ForEach(batchParams, parallelOptions, batch =>
                new OfdConverter(batch.PartOfd).ToPdf(batch.PartPdf));

            MergePdfs(batchParams.Select(b => b.PartPdf).ToList(), outputPath);
        }
        finally
        {
            try { Directory.Delete(tmpDir, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Returns the total number of pages in an OFD file by reading its Document.xml.
    /// </summary>
    private static int GetOfdPageCount(string ofdPath)
    {
        using var archive = ZipFile.OpenRead(ofdPath);

        var ofdXmlEntry = archive.Entries.FirstOrDefault(
            e => e.FullName.Equals("OFD.xml", StringComparison.OrdinalIgnoreCase));
        if (ofdXmlEntry == null) return 0;

        string docRootPath;
        using (var stream = ofdXmlEntry.Open())
            docRootPath = XDocument.Load(stream).Descendants()
                .First(e => e.Name.LocalName == "DocRoot")
                .Value.Trim().Replace('\\', '/');

        var docEntry = archive.Entries.FirstOrDefault(
            e => e.FullName.Replace('\\', '/').Equals(docRootPath, StringComparison.OrdinalIgnoreCase));
        if (docEntry == null) return 0;

        using var docStream = docEntry.Open();
        return XDocument.Load(docStream).Descendants()
            .Count(e => e.Name.LocalName == "Page" && e.Attribute("BaseLoc") != null);
    }

    /// <summary>
    /// Creates a sub-OFD ZIP file containing only the specified page range.
    /// All non-page entries (resources, fonts, etc.) are copied as-is so that
    /// each sub-OFD is a valid, self-contained document.
    /// Font files with a .otf or .ttf extension that actually contain raw CFF data (not a
    /// proper OTF/TrueType container) are renamed to .cff to prevent FreeSpire.PDF from
    /// misidentifying them and throwing a NullReferenceException during conversion.
    /// </summary>
    private static void CreateSubOfd(string sourcePath, string destPath, int startPage, int pageCount)
    {
        using var src = ZipFile.OpenRead(sourcePath);

        var ofdXmlEntry = src.Entries.First(
            e => e.FullName.Equals("OFD.xml", StringComparison.OrdinalIgnoreCase));

        string docRootPath;
        using (var stream = ofdXmlEntry.Open())
            docRootPath = XDocument.Load(stream).Descendants()
                .First(e => e.Name.LocalName == "DocRoot")
                .Value.Trim().Replace('\\', '/');

        var docFolder = GetDirectory(docRootPath);

        XDocument docXml;
        List<XElement> allPages;
        var docEntry = src.Entries.First(
            e => e.FullName.Replace('\\', '/').Equals(docRootPath, StringComparison.OrdinalIgnoreCase));
        using (var stream = docEntry.Open())
        {
            docXml = XDocument.Load(stream);
            allPages = docXml.Descendants()
                .Where(e => e.Name.LocalName == "Page" && e.Attribute("BaseLoc") != null)
                .ToList();
        }

        var selectedPages = allPages.Skip(startPage).Take(pageCount).ToList();

        var pagesFolderPrefix = string.IsNullOrEmpty(docFolder)
            ? "Pages/"
            : $"{docFolder}/Pages/";

        var includedPrefixes = selectedPages
            .Select(p =>
            {
                var baseLoc = p.Attribute("BaseLoc")!.Value.Replace('\\', '/').TrimStart('/');
                // BaseLoc may point to a content file (e.g. Pages/Page_0/Content.xml) or a
                // directory (e.g. Pages/Page_0).  Derive the directory so that StartsWith
                // matching covers all files inside that page folder.
                var pageDir = baseLoc.Contains('/') ? baseLoc[..baseLoc.LastIndexOf('/')] : baseLoc;
                var fullDir = string.IsNullOrEmpty(docFolder) ? pageDir : $"{docFolder}/{pageDir}";
                return fullDir + "/";
            })
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        // Build a rename map for font files that have a .otf extension but contain raw
        // CFF data.  FreeSpire.PDF uses the file extension to determine how to parse a
        // font: passing raw CFF bytes to its OTF parser causes a NullReferenceException.
        // Renaming such files to .cff lets Spire.PDF parse them correctly.
        // In addition, some raw CFF fonts use CFF Expert Encoding whose glyph lookup
        // triggers a NullReferenceException deep inside Spire.PDF's CFF renderer even
        // after renaming.  For those fonts we also strip all TextObject elements that
        // reference them from the page content, suppressing the crash.
        // Both PublicRes and DocumentRes resource files are scanned.
        var resourcePaths = GetResourcePaths(docXml, docFolder);
        var resourcePathSet = new HashSet<string>(resourcePaths, StringComparer.OrdinalIgnoreCase);
        var (fontRenameMap, rawCffFontIds) = BuildRawCffFontInfo(src, resourcePaths);

        using var dest = ZipFile.Open(destPath, ZipArchiveMode.Create);

        foreach (var entry in src.Entries)
        {
            var name = entry.FullName.Replace('\\', '/');
            if (name.EndsWith("/")) continue; // skip directory entries

            // Exclude pages that are not in the selected range
            if (name.StartsWith(pagesFolderPrefix, StringComparison.OrdinalIgnoreCase) &&
                !includedPrefixes.Any(p => name.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
                continue;

            if (name.Equals(docRootPath, StringComparison.OrdinalIgnoreCase))
            {
                // Write a modified Document.xml referencing only the selected pages
                var modifiedDoc = new XDocument(docXml);
                var pagesEl = modifiedDoc.Descendants().First(e => e.Name.LocalName == "Pages");
                pagesEl.RemoveNodes();
                foreach (var page in selectedPages)
                    pagesEl.Add(new XElement(page));

                using var destStream = dest.CreateEntry(name).Open();
                modifiedDoc.Save(destStream);
            }
            else if (fontRenameMap.Count > 0 && resourcePathSet.Contains(name))
            {
                // Write a modified resource XML (PublicRes or DocumentRes) with updated
                // font file names so that references point to the renamed .cff files.
                XDocument resXml;
                using (var srcStream = entry.Open())
                    resXml = XDocument.Load(srcStream);

                var resFolder = GetDirectory(name);
                var baseLoc = resXml.Root?.Attribute("BaseLoc")?.Value?.Trim().Replace('\\', '/') ?? "";
                var fontBasePath = CombinePaths(resFolder, baseLoc);

                foreach (var fontFileEl in resXml.Descendants()
                    .Where(e => e.Name.LocalName == "FontFile"))
                {
                    var fontFileName = fontFileEl.Value.Trim().Replace('\\', '/');
                    var fullFontPath = CombinePaths(fontBasePath, fontFileName);
                    if (fontRenameMap.TryGetValue(fullFontPath, out var newFullPath))
                    {
                        int sep = newFullPath.LastIndexOf('/');
                        fontFileEl.Value = sep >= 0 ? newFullPath[(sep + 1)..] : newFullPath;
                    }
                }

                using var destStream = dest.CreateEntry(name).Open();
                resXml.Save(destStream);
            }
            else if (fontRenameMap.TryGetValue(name, out var renamedPath))
            {
                // Copy a font file under its corrected .cff name
                using var destStream = dest.CreateEntry(renamedPath).Open();
                using var srcStream = entry.Open();
                srcStream.CopyTo(destStream);
            }
            else if (rawCffFontIds.Count > 0 &&
                     includedPrefixes.Any(p => name.StartsWith(p, StringComparison.OrdinalIgnoreCase)) &&
                     name.EndsWith("/Content.xml", StringComparison.OrdinalIgnoreCase))
            {
                // Strip any TextObject elements whose Font attribute refers to a raw CFF
                // font that Spire.PDF cannot render without crashing.  Renaming the font
                // file to .cff is insufficient for some CFF fonts (e.g. those using Expert
                // Encoding): Spire.PDF's raw-CFF renderer still throws NullReferenceException.
                // Removing the offending TextObjects prevents the crash; the rest of the
                // page content (images, paths, other text) is preserved intact.
                XDocument contentXml;
                using (var srcStream = entry.Open())
                    contentXml = XDocument.Load(srcStream);

                var toRemove = contentXml.Descendants()
                    .Where(e => e.Name.LocalName == "TextObject" &&
                                rawCffFontIds.Contains((string?)e.Attribute("Font") ?? ""))
                    .ToList();
                foreach (var el in toRemove)
                    el.Remove();

                using var destStream = dest.CreateEntry(name).Open();
                contentXml.Save(destStream);
            }
            else
            {
                using var destStream = dest.CreateEntry(name).Open();
                using var srcStream = entry.Open();
                srcStream.CopyTo(destStream);
            }
        }
    }

    /// <summary>
    /// Returns the archive paths of resource XML files declared in Document.xml.
    /// Both <c>PublicRes</c> and <c>DocumentRes</c> elements are checked so that fonts
    /// embedded via either resource type are included in the rename map.
    /// </summary>
    private static List<string> GetResourcePaths(XDocument docXml, string docFolder)
    {
        var result = new List<string>();
        foreach (var elementName in new[] { "PublicRes", "DocumentRes" })
        {
            var rel = docXml.Descendants()
                .FirstOrDefault(e => e.Name.LocalName == elementName)
                ?.Value.Trim().Replace('\\', '/');
            if (!string.IsNullOrEmpty(rel))
                result.Add(CombinePaths(docFolder, rel));
        }
        return result;
    }

    /// <summary>
    /// Combines two path segments, omitting a separator when either part is empty.
    /// </summary>
    private static string CombinePaths(string prefix, string suffix)
    {
        if (string.IsNullOrEmpty(prefix)) return suffix;
        if (string.IsNullOrEmpty(suffix)) return prefix;
        return $"{prefix}/{suffix}";
    }

    /// <summary>
    /// Returns the directory portion of an archive path (everything before the last '/'),
    /// or an empty string if the path contains no '/'.
    /// </summary>
    private static string GetDirectory(string path) =>
        path.Contains('/') ? path[..path.LastIndexOf('/')] : "";

    /// <summary>
    /// Reads one or more resource XML files (PublicRes / DocumentRes) from the source
    /// archive and returns:
    /// <list type="bullet">
    ///   <item>a mapping from the original archive path of each font file whose name ends
    ///   in ".otf" or ".ttf" but whose content is raw CFF data (not a proper OTF/TrueType
    ///   container) to its corrected ".cff" archive path; and</item>
    ///   <item>the set of Font element IDs (e.g. "764") for those raw-CFF fonts, used to
    ///   strip problematic TextObject elements from page content.</item>
    /// </list>
    /// Both collections are empty when no such fonts exist.
    /// </summary>
    private static (Dictionary<string, string> RenameMap, HashSet<string> RawCffFontIds)
        BuildRawCffFontInfo(ZipArchive src, IEnumerable<string> resourcePaths)
    {
        var renameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var rawCffFontIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var resourcePath in resourcePaths)
        {
            var resEntry = src.Entries.FirstOrDefault(
                e => e.FullName.Replace('\\', '/').Equals(resourcePath, StringComparison.OrdinalIgnoreCase));
            if (resEntry == null) continue;

            XDocument resXml;
            using (var stream = resEntry.Open())
                resXml = XDocument.Load(stream);

            var resFolder = GetDirectory(resourcePath);
            var baseLoc = resXml.Root?.Attribute("BaseLoc")?.Value?.Trim().Replace('\\', '/') ?? "";
            var fontBasePath = CombinePaths(resFolder, baseLoc);

            foreach (var fontEl in resXml.Descendants()
                .Where(e => e.Name.LocalName == "Font"))
            {
                var fontFileEl = fontEl.Descendants()
                    .FirstOrDefault(e => e.Name.LocalName == "FontFile");
                if (fontFileEl == null) continue;

                var fontFileName = fontFileEl.Value.Trim().Replace('\\', '/');

                // Only inspect font file extensions that Spire.PDF may misparse when they
                // actually contain raw CFF data.  Files already named .cff are passed to
                // the CFF parser directly and do not need to be renamed.  Files with no
                // extension or any other extension are left unchanged.
                var fontExt = Path.GetExtension(fontFileName).ToLowerInvariant();
                if (fontExt != ".otf" && fontExt != ".ttf")
                    continue;

                var fullFontPath = CombinePaths(fontBasePath, fontFileName);
                var fontEntry = src.Entries.FirstOrDefault(
                    e => e.FullName.Replace('\\', '/').Equals(fullFontPath, StringComparison.OrdinalIgnoreCase));
                if (fontEntry == null) continue;

                // Read just the first 4 bytes to detect the file type.
                // Valid font-container signatures:
                //   00 01 00 00  – TrueType / OpenType with TT outlines
                //   4F 54 54 4F  – "OTTO" – CFF-based OpenType
                //   74 72 75 65  – "true" – Apple TrueType
                //   74 74 63 66  – "ttcf" – TrueType Collection
                // Anything else is treated as raw CFF data stored with a misleading
                // extension and must be renamed to .cff so that Spire.PDF uses its CFF
                // parser instead of its OTF/TrueType parser (which throws
                // NullReferenceException on raw CFF input).
                var header = new byte[4];
                using (var fs = fontEntry.Open())
                {
                    if (fs.Read(header, 0, 4) < 4) continue;
                }

                bool isValidFontContainer =
                    (header[0] == 0x00 && header[1] == 0x01 && header[2] == 0x00 && header[3] == 0x00) ||
                    (header[0] == 0x4F && header[1] == 0x54 && header[2] == 0x54 && header[3] == 0x4F) ||
                    (header[0] == 0x74 && header[1] == 0x72 && header[2] == 0x75 && header[3] == 0x65) ||
                    (header[0] == 0x74 && header[1] == 0x74 && header[2] == 0x63 && header[3] == 0x66);

                if (!isValidFontContainer)
                {
                    renameMap[fullFontPath] = fullFontPath[..^fontExt.Length] + ".cff";
                    var fontId = fontEl.Attribute("ID")?.Value;
                    if (fontId != null) rawCffFontIds.Add(fontId);
                }
            }
        }

        return (renameMap, rawCffFontIds);
    }

    /// <summary>
    /// Merges multiple PDF files into a single output PDF using iText.
    /// </summary>
    private static void MergePdfs(IReadOnlyList<string> pdfPaths, string outputPath)
    {
        var readers = pdfPaths.Select(p => new PdfReader(p)).ToList();
        var srcDocs = readers.Select(r => new PdfDocument(r)).ToList();
        try
        {
            using var writer = new PdfWriter(outputPath);
            using var mergedDoc = new PdfDocument(writer);
            foreach (var srcDoc in srcDocs)
                srcDoc.CopyPagesTo(1, srcDoc.GetNumberOfPages(), mergedDoc);
        }
        finally
        {
            foreach (var doc in srcDocs)
                try { doc.Close(); } catch { }
        }
    }

    /// <summary>
    /// Removes the Spire.PDF evaluation warning from a PDF file by locating the warning text
    /// via content-stream parsing and painting a white rectangle over the matched lines.
    /// </summary>
    private static void RemoveEvaluationWarning(string pdfPath, ILogger logger)
    {
        var tempPath = pdfPath + ".clean";
        try
        {
            using (var reader = new PdfReader(pdfPath))
            using (var writer = new PdfWriter(tempPath))
            using (var doc = new PdfDocument(reader, writer))
            {
                for (int i = 1; i <= doc.GetNumberOfPages(); i++)
                {
                    var page = doc.GetPage(i);
                    var warningRect = FindEvaluationWarningRect(page);
                    if (warningRect == null)
                        continue;

                    var canvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), doc);
                    canvas.SaveState()
                          .SetFillColor(ColorConstants.WHITE)
                          .Rectangle(warningRect)
                          .Fill()
                          .RestoreState()
                          .Release();
                }
            }
            System.IO.File.Move(tempPath, pdfPath, overwrite: true);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Failed to remove Spire.PDF evaluation warning from {PdfPath}; the original PDF will be used as-is.", pdfPath);
            if (System.IO.File.Exists(tempPath))
                System.IO.File.Delete(tempPath);
        }
    }

    /// <summary>
    /// Parses the page content stream for text containing "Evaluation Warning" and returns
    /// a page-width rectangle that covers those text lines. Returns null if no match is found.
    /// </summary>
    private static ITextRect? FindEvaluationWarningRect(PdfPage page)
    {
        var locator = new EvaluationWarningLocator();
        new PdfCanvasProcessor(locator).ProcessPageContent(page);
        return locator.GetCoveringRect(page.GetPageSize());
    }

    /// <summary>
    /// Collects TextRenderInfo events from the page content stream, groups them into
    /// text lines by baseline Y, and identifies lines containing "Evaluation Warning"
    /// for targeted removal.
    /// </summary>
    private sealed class EvaluationWarningLocator : IEventListener
    {
        private readonly List<TextChunkInfo> _chunks = new();

        private struct TextChunkInfo
        {
            public string Text;
            public float BaselineY;
            public float Top;
            public float Bottom;
        }

        public void EventOccurred(IEventData data, EventType type)
        {
            if (data is not TextRenderInfo info) return;
            var text = info.GetText();
            if (string.IsNullOrEmpty(text)) return;

            _chunks.Add(new TextChunkInfo
            {
                Text = text,
                BaselineY = info.GetBaseline().GetBoundingRectangle().GetBottom(),
                Top = info.GetAscentLine().GetBoundingRectangle().GetTop(),
                Bottom = info.GetDescentLine().GetBoundingRectangle().GetBottom(),
            });
        }

        public ICollection<EventType> GetSupportedEvents() =>
            new HashSet<EventType> { EventType.RENDER_TEXT };

        /// <summary>
        /// Groups chunks by baseline Y, concatenates each line's text, and returns a
        /// page-width rectangle covering every line that contains "Evaluation Warning".
        /// Returns null if no match is found.
        /// </summary>
        public ITextRect? GetCoveringRect(ITextRect pageSize)
        {
            const float yTolerance = 2f;
            const float margin = 2f;

            // Group chunks into lines by baseline Y bucket (O(n) via dictionary).
            var lines = new Dictionary<int, List<TextChunkInfo>>();
            foreach (var chunk in _chunks)
            {
                int key = (int)Math.Round(chunk.BaselineY / yTolerance);
                if (!lines.TryGetValue(key, out var line))
                    lines[key] = line = new List<TextChunkInfo>();
                line.Add(chunk);
            }

            // Find lines whose concatenated text contains "Evaluation Warning".
            float minBottom = float.MaxValue;
            float maxTop = float.MinValue;
            bool found = false;

            foreach (var line in lines.Values)
            {
                var lineText = string.Concat(line.Select(c => c.Text));
                if (lineText.IndexOf("Evaluation Warning", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                found = true;
                foreach (var chunk in line)
                {
                    if (chunk.Bottom < minBottom) minBottom = chunk.Bottom;
                    if (chunk.Top > maxTop) maxTop = chunk.Top;
                }
            }

            if (!found)
                return null;

            float bottom = minBottom - margin;
            return new ITextRect(pageSize.GetLeft(), bottom, pageSize.GetWidth(), maxTop - bottom + margin);
        }
    }
}
