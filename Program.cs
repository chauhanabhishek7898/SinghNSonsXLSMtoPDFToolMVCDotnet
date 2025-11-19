using ExcelToPdfConverter.Services;
using Microsoft.AspNetCore.Http.Features;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// ✅ EPPlus 8.2.1 License
ExcelPackage.License.SetNonCommercialPersonal("ExcelToPdf Converter");

// Add services to the container
builder.Services.AddControllersWithViews();
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

// Register services
builder.Services.AddScoped<LibreOfficeService>();
builder.Services.AddScoped<ExcelPreviewService>();
builder.Services.AddScoped<ExcelProcessingService>();

// Increase file upload limits
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 104857600; // 100MB
});

// Add logging
builder.Services.AddLogging();

var app = builder.Build();

// Configure the HTTP request pipeline
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
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

Console.WriteLine("Excel to PDF Converter started successfully!");
Console.WriteLine("Supports: .xlsx, .xls, .xlsm files");

app.Run();



//using ExcelToPdfConverter.Services;
//using Microsoft.AspNetCore.Http.Features;
//using OfficeOpenXml;

//var builder = WebApplication.CreateBuilder(args);

//// ✅ EPPlus 8+ License Set karein (Yeh line sabse pehle)
////ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
//ExcelPackage.License.SetNonCommercialPersonal("ExcelToPdf Converter");

//// Add services to the container
//builder.Services.AddControllersWithViews();
//builder.Services.AddSession(options =>
//{
//    options.IdleTimeout = TimeSpan.FromMinutes(30);
//    options.Cookie.HttpOnly = true;
//    options.Cookie.IsEssential = true;
//});

//// Register services
//builder.Services.AddScoped<LibreOfficeService>();
//builder.Services.AddScoped<ExcelPreviewService>();

//// Increase file upload limits
//builder.Services.Configure<FormOptions>(options =>
//{
//    options.MultipartBodyLengthLimit = 104857600; // 100MB
//});

//// Add logging
//builder.Services.AddLogging();

//var app = builder.Build();

//// Configure the HTTP request pipeline
//if (!app.Environment.IsDevelopment())
//{
//    app.UseExceptionHandler("/Home/Error");
//    app.UseHsts();
//}

//app.UseHttpsRedirection();
//app.UseStaticFiles();
//app.UseRouting();
//app.UseAuthorization();
//app.UseSession();

//app.MapControllerRoute(
//    name: "default",
//    pattern: "{controller=Home}/{action=Index}/{id?}");

//// Cleanup old files on startup
//using (var scope = app.Services.CreateScope())
//{
//    var libreOfficeService = scope.ServiceProvider.GetRequiredService<LibreOfficeService>();
//    libreOfficeService.CleanupOldFiles();
//}

//Console.WriteLine("Excel to PDF Converter started successfully!");
//Console.WriteLine("Make sure LibreOffice is installed at: C:\\Program Files\\LibreOffice\\program\\soffice.exe");

//app.Run();