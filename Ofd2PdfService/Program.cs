// Register system font directories with Spire.PDF on Linux so that Chinese (CJK)
// characters are rendered correctly instead of appearing as squares.
if (OperatingSystem.IsLinux())
{
    Spire.Pdf.PdfDocument.LoadCustomFontFolder("/usr/share/fonts");
}

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

// Allow large OFD file uploads (default is 30 MB; adjust as needed)
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 100 * 1024 * 1024; // 100 MB
});
builder.WebHost.ConfigureKestrel(serverOptions =>
{
    serverOptions.Limits.MaxRequestBodySize = 100 * 1024 * 1024; // 100 MB
});

var app = builder.Build();

app.UseHttpsRedirection();
app.MapControllers();

app.Run();

// Expose Program class for WebApplicationFactory in integration tests.
public partial class Program { }
