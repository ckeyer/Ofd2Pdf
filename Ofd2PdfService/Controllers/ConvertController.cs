using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.AspNetCore.Mvc;
using Spire.Pdf.Conversion;
using ITextRect = iText.Kernel.Geom.Rectangle;

namespace Ofd2PdfService.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ConvertController : ControllerBase
{
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
            var converter = new OfdConverter(inputPath);
            converter.ToPdf(outputPath);
            RemoveEvaluationWarningPage(outputPath, _logger);
            RemoveEvaluationWarning(outputPath, _logger);

            var pdfBytes = await System.IO.File.ReadAllBytesAsync(outputPath);
            var pdfFileName = Path.ChangeExtension(Path.GetFileName(file.FileName), ".pdf");

            _logger.LogInformation("Conversion succeeded: {FileName}", pdfFileName);
            return File(pdfBytes, "application/pdf", pdfFileName);
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
    /// Removes the first page of the PDF if it contains the Spire.PDF evaluation warning.
    /// The evaluation warning introduced by the free version of Spire.PDF only appears on the
    /// first page. Adding a blank page (when necessary) and then deleting the first page
    /// effectively erases the warning without affecting the remaining content pages.
    /// </summary>
    private static void RemoveEvaluationWarningPage(string pdfPath, ILogger logger)
    {
        var tempPath = pdfPath + ".nocover";
        try
        {
            int totalPages;
            bool firstPageHasWarning;

            using (var reader = new PdfReader(pdfPath))
            using (var doc = new PdfDocument(reader))
            {
                totalPages = doc.GetNumberOfPages();
                firstPageHasWarning = FindEvaluationWarningRect(doc.GetPage(1)) != null;
            }

            if (!firstPageHasWarning)
                return;

            using (var reader = new PdfReader(pdfPath))
            using (var writer = new PdfWriter(tempPath))
            using (var srcDoc = new PdfDocument(reader))
            using (var destDoc = new PdfDocument(writer))
            {
                if (totalPages > 1)
                    srcDoc.CopyPagesTo(2, totalPages, destDoc);
                else
                    destDoc.AddNewPage();
            }

            System.IO.File.Move(tempPath, pdfPath, overwrite: true);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Failed to remove Spire.PDF evaluation warning page from {PdfPath}; the original PDF will be used as-is.", pdfPath);
            if (System.IO.File.Exists(tempPath))
                System.IO.File.Delete(tempPath);
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
