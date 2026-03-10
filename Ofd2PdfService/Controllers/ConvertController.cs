using Microsoft.AspNetCore.Mvc;
using Spire.Pdf.Conversion;

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
}
