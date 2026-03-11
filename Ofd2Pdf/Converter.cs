using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Spire.Pdf.Conversion;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ofd2Pdf
{
    public enum ConvertResult
    {
        Successful,
        Failed
    }
    public class Converter
    {
        public ConvertResult ConvertToPdf(string Input, string OutPut)
        {
            Console.WriteLine(Input + " " + OutPut);
            if (Input == null || OutPut == null)
            {
                return ConvertResult.Failed;
            }

            if (!File.Exists(Input))
            {
                return ConvertResult.Failed;
            }

            try
            {
                OfdConverter converter = new OfdConverter(Input);
                converter.ToPdf(OutPut);
                RemoveEvaluationWarningPage(OutPut);
                RemoveEvaluationWarning(OutPut);
                return ConvertResult.Successful;
            }
            catch (Exception)
            {
                return ConvertResult.Failed;
            }
        }

        /// <summary>
        /// Removes the first page of the PDF if it contains the Spire.PDF evaluation warning.
        /// The evaluation warning introduced by the free version of Spire.PDF only appears on the
        /// first page. Adding a blank page (when necessary) and then deleting the first page
        /// effectively erases the warning without affecting the remaining content pages.
        /// </summary>
        private static void RemoveEvaluationWarningPage(string pdfPath)
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

                System.IO.File.Replace(tempPath, pdfPath, destinationBackupFileName: null);
            }
            catch (Exception)
            {
                // Removal is best-effort; if it fails the original PDF is kept.
                if (System.IO.File.Exists(tempPath))
                    System.IO.File.Delete(tempPath);
            }
        }

        private static void RemoveEvaluationWarning(string pdfPath)
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
                // File.Replace atomically replaces the destination with the source on the same volume.
                System.IO.File.Replace(tempPath, pdfPath, destinationBackupFileName: null);
            }
            catch (Exception)
            {
                // Removal of the evaluation warning is best-effort; if it fails the original
                // converted PDF (with the watermark) is kept so the conversion still succeeds.
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        /// <summary>
        /// Parses the page content stream for text containing "Evaluation Warning" and returns
        /// a page-width rectangle that covers those text lines. Returns null if no match is found.
        /// </summary>
        private static Rectangle FindEvaluationWarningRect(PdfPage page)
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
            private readonly List<TextChunkInfo> _chunks = new List<TextChunkInfo>();

            private struct TextChunkInfo
            {
                public string Text;
                public float BaselineY;
                public float Top;
                public float Bottom;
            }

            public void EventOccurred(IEventData data, EventType type)
            {
                var info = data as TextRenderInfo;
                if (info == null) return;
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

            public ICollection<EventType> GetSupportedEvents()
            {
                return new HashSet<EventType> { EventType.RENDER_TEXT };
            }

            /// <summary>
            /// Groups chunks by baseline Y, concatenates each line's text, and returns a
            /// page-width rectangle covering every line that contains "Evaluation Warning".
            /// Returns null if no match is found.
            /// </summary>
            public Rectangle GetCoveringRect(Rectangle pageSize)
            {
                const float yTolerance = 2f;
                const float margin = 2f;

                // Group chunks into lines by baseline Y bucket (O(n) via dictionary).
                var lines = new Dictionary<int, List<TextChunkInfo>>();
                foreach (var chunk in _chunks)
                {
                    int key = (int)Math.Round(chunk.BaselineY / yTolerance);
                    List<TextChunkInfo> line;
                    if (!lines.TryGetValue(key, out line))
                    {
                        line = new List<TextChunkInfo>();
                        lines[key] = line;
                    }
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
                return new Rectangle(pageSize.GetLeft(), bottom, pageSize.GetWidth(), maxTop - bottom + margin);
            }
        }
    }
}
