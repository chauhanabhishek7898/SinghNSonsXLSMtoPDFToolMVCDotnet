using ExcelToPdfConverter.Controllers;
using ExcelToPdfConverter.Services;
using Microsoft.AspNetCore.Http.Features;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// ✅ EPPlus 8.2.1 License
ExcelPackage.License.SetNonCommercialPersonal("ExcelToPdf Converter");

// Add services to the container
builder.Services.AddControllersWithViews(options =>
{
    options.MaxModelBindingCollectionSize = 1024 * 1024; // Increase model binding size
});

builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
});

// Register services with enhanced configurations
builder.Services.AddScoped<LibreOfficeService>(provider =>
{
    var env = provider.GetRequiredService<IWebHostEnvironment>();
    return new LibreOfficeService(env);
});

builder.Services.AddScoped<ExcelPreviewService>();
builder.Services.AddScoped<ExcelProcessingService>();
builder.Services.AddScoped<PdfProcessingService>();
builder.Services.AddScoped<PdfCompressionService>();
builder.Services.AddHostedService<PdfCleanupService>();

// Increase file upload and request limits
builder.Services.Configure<FormOptions>(options =>
{
    options.ValueLengthLimit = int.MaxValue;
    options.MultipartBodyLengthLimit = 104857600; // 100MB
    options.MultipartBoundaryLengthLimit = int.MaxValue;
    options.MultipartHeadersCountLimit = int.MaxValue;
    options.MultipartHeadersLengthLimit = int.MaxValue;
});

builder.Services.Configure<IISServerOptions>(options =>
{
    options.MaxRequestBodySize = 104857600; // 100MB
});

// Add logging
builder.Services.AddLogging(config =>
{
    config.AddConsole();
    config.AddDebug();
    config.AddConfiguration(builder.Configuration.GetSection("Logging"));
});

var app = builder.Build();

// Configure the HTTP request pipeline
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles(new StaticFileOptions
{
    ServeUnknownFileTypes = true,
    DefaultContentType = "application/octet-stream"
});

app.UseRouting();
app.UseAuthorization();
app.UseSession();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

// Cleanup old files on startup
using (var scope = app.Services.CreateScope())
{
    var libreOfficeService = scope.ServiceProvider.GetRequiredService<LibreOfficeService>();
    libreOfficeService.CleanupOldFiles();
}

Console.WriteLine("✅ Excel to PDF Converter started successfully!");
Console.WriteLine("✅ Supports: .xlsx, .xls, .xlsm files");
Console.WriteLine("✅ Color preservation enabled");
Console.WriteLine("✅ Enhanced PDF preview available");


// ✅ Create previews directory if not exists
var previewsDir = Path.Combine(app.Environment.WebRootPath, "previews");
if (!Directory.Exists(previewsDir))
{
    Directory.CreateDirectory(previewsDir);
}

// ✅ Cleanup old preview files on startup
CleanupOldPreviews(previewsDir);


app.Run();


void CleanupOldPreviews(string directory)
{
    try
    {
        var files = Directory.GetFiles(directory, "*.pdf");
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            if (fileInfo.CreationTime < DateTime.Now.AddHours(-2))
            {
                fileInfo.Delete();
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error cleaning up previews: {ex.Message}");
    }
}