using Microsoft.AspNetCore.Mvc.Testing;
using System.Net;
using System.Net.Http.Headers;

namespace Ofd2PdfService.Tests;

/// <summary>
/// Integration tests for the OFD-to-PDF conversion endpoint.
/// Tests POST /api/convert with the sample OFD files in the testdata directory.
/// </summary>
public class ConvertControllerTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly WebApplicationFactory<Program> _factory;

    // Resolve the testdata directory by searching upward from the test binary location.
    private static readonly string TestDataDir = FindTestDataDir();

    private static string FindTestDataDir()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, "testdata");
            if (Directory.Exists(candidate))
                return candidate;
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("testdata directory not found relative to " + AppContext.BaseDirectory);
    }

    public ConvertControllerTests(WebApplicationFactory<Program> factory)
    {
        _factory = factory;
    }

    /// <summary>
    /// Posting sample1.ofd must succeed (HTTP 200) and return a non-empty PDF.
    /// This specifically exercises pages 31-34 that contain the problematic raw-CFF
    /// font (font_764_764.otf / DINPro-Black) that previously triggered a
    /// NullReferenceException inside Spire.PDF's CFF renderer.
    /// </summary>
    [Fact]
    public async Task Post_Sample1Ofd_ReturnsSuccessfulPdf()
    {
        var client = _factory.CreateClient();
        var ofdPath = Path.Combine(TestDataDir, "sample1.ofd");
        Assert.True(File.Exists(ofdPath), $"Test data not found: {ofdPath}");

        using var content = new MultipartFormDataContent();
        var fileBytes = await File.ReadAllBytesAsync(ofdPath);
        var fileContent = new ByteArrayContent(fileBytes);
        fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");
        content.Add(fileContent, "file", "sample1.ofd");

        var response = await client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("application/pdf", response.Content.Headers.ContentType?.MediaType);

        var pdfBytes = await response.Content.ReadAsByteArrayAsync();
        Assert.NotEmpty(pdfBytes);
        // PDF files start with the %PDF- header
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(pdfBytes, 0, 5));
    }

    /// <summary>
    /// Posting a non-OFD file must return HTTP 400.
    /// </summary>
    [Fact]
    public async Task Post_NonOfdFile_ReturnsBadRequest()
    {
        var client = _factory.CreateClient();

        using var content = new MultipartFormDataContent();
        var fileContent = new ByteArrayContent(new byte[] { 1, 2, 3 });
        fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");
        content.Add(fileContent, "file", "test.txt");

        var response = await client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    /// <summary>
    /// Posting with no file must return HTTP 400.
    /// </summary>
    [Fact]
    public async Task Post_NoFile_ReturnsBadRequest()
    {
        var client = _factory.CreateClient();
        using var content = new MultipartFormDataContent();

        var response = await client.PostAsync("/api/convert", content);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }
}
