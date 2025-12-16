using ExcelToPdfConverter.Models;
using ExcelToPdfConverter.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Utils;
using System.Text.Json;
using Path = System.IO.Path;
using Org.BouncyCastle.Bcpg.Sig;

namespace ExcelToPdfConverter.Controllers
{
    public class HomeController : Controller
    {
        private readonly LibreOfficeService _libreOfficeService;
        private readonly ExcelPreviewService _previewService;
        private readonly ExcelProcessingService _excelProcessingService;
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<HomeController> _logger;

        public HomeController(LibreOfficeService libreOfficeService,
                            ExcelPreviewService previewService,
                            ExcelProcessingService excelProcessingService,
                            IWebHostEnvironment environment,
                            ILogger<HomeController> logger)
        {
            _libreOfficeService = libreOfficeService;
            _previewService = previewService;
            _excelProcessingService = excelProcessingService;
            _environment = environment;
            _logger = logger;
        }

        public IActionResult Index()
        {
            _logger.LogInformation("Home page accessed");
            ViewBag.ValidationResult = null;
            ViewBag.Error = null;
            return View();
        }

        [HttpPost]
        [RequestSizeLimit(100_000_000)]
        public IActionResult Upload(ExcelUploadModel model)
        {
            try
            {
                _logger.LogInformation("Upload action started");
                if (model.ExcelFile == null || model.ExcelFile.Length == 0)
                {
                    ViewBag.Error = "Please select an Excel file.";
                    return View("Index");
                }

                var extension = System.IO.Path.GetExtension(model.ExcelFile.FileName)?.ToLower();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                {
                    ViewBag.Error = "Please upload only Excel files (.xlsx, .xls, or .xlsm).";
                    return View("Index");
                }

                var validationResult = _previewService.QuickValidate(model.ExcelFile);
                ViewBag.ValidationResult = validationResult;

                var sessionId = Guid.NewGuid().ToString();
                var previewModel = _previewService.GeneratePreview(model.ExcelFile, sessionId);

                var filePath = _libreOfficeService.SaveUploadedFile(model.ExcelFile);

                HttpContext.Session.SetString(sessionId + "_filePath", filePath);
                HttpContext.Session.SetString(sessionId + "_fileName", model.ExcelFile.FileName ?? "unknown");
                HttpContext.Session.SetString("CurrentSessionId", sessionId);

                ViewBag.SessionId = sessionId;

                return View("Preview", previewModel);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during file upload");
                ViewBag.Error = $"Error processing file: {ex.Message}";
                return View("Index");
            }
        }

        [HttpPost]
        public async Task<IActionResult> UploadPdfFiles([FromForm] string sessionId, [FromForm] List<IFormFile> pdfFiles)
        {
            try
            {
                Console.WriteLine($"=== UploadPdfFiles Called ===");
                Console.WriteLine($"Session ID: {sessionId}");
                Console.WriteLine($"Number of files: {pdfFiles?.Count ?? 0}");

                if (pdfFiles == null || pdfFiles.Count == 0)
                {
                    return Json(new { success = false, message = "No PDF files uploaded." });
                }

                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", sessionId);
                Directory.CreateDirectory(uploadDirectory);

                var uploadedFiles = new List<object>();

                foreach (var file in pdfFiles)
                {
                    if (file.Length > 50 * 1024 * 1024) // 50MB limit
                    {
                        return Json(new { success = false, message = $"{file.FileName} exceeds 50MB limit." });
                    }

                    var fileName = Guid.NewGuid() + Path.GetExtension(file.FileName);
                    var filePath = Path.Combine(uploadDirectory, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    uploadedFiles.Add(new
                    {
                        name = file.FileName,
                        path = filePath,
                        size = file.Length,
                        uploadTime = DateTime.Now
                    });

                    Console.WriteLine($"✅ PDF uploaded: {file.FileName} -> {filePath}");
                }

                // Store uploaded files info in session
                HttpContext.Session.SetString(sessionId + "_uploadedPdfs", JsonSerializer.Serialize(uploadedFiles));

                return Json(new
                {
                    success = true,
                    message = $"{pdfFiles.Count} PDF file(s) uploaded successfully.",
                    uploadedFiles = uploadedFiles.Select(f => new
                    {
                        name = ((dynamic)f).name,
                        size = ((dynamic)f).size
                    }).ToList()
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UploadPdfFiles: {ex}");
                return Json(new { success = false, message = $"Error uploading files: {ex.Message}" });
            }
        }

        [HttpPost]
        public IActionResult RemovePdfFile([FromBody] RemovePdfRequest request)
        {
            try
            {
                Console.WriteLine($"=== RemovePdfFile Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"File name: {request.FileName}");

                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);
                
                if (!Directory.Exists(uploadDirectory))
                {
                    return Json(new { success = false, message = "Upload directory not found." });
                }

                // Get uploaded files from session
                var uploadedFilesJson = HttpContext.Session.GetString(request.SessionId + "_uploadedPdfs");
                if (string.IsNullOrEmpty(uploadedFilesJson))
                {
                    return Json(new { success = false, message = "No uploaded files found." });
                }

                var uploadedFiles = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(uploadedFilesJson);
                var fileToRemove = uploadedFiles.FirstOrDefault(f => f["name"].ToString() == request.FileName);

                if (fileToRemove != null)
                {
                    var filePath = fileToRemove["path"].ToString();
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                        Console.WriteLine($"🗑️ Deleted PDF file: {filePath}");
                    }

                    uploadedFiles.Remove(fileToRemove);
                    HttpContext.Session.SetString(request.SessionId + "_uploadedPdfs", JsonSerializer.Serialize(uploadedFiles));

                    return Json(new { success = true, message = "PDF file removed successfully." });
                }

                return Json(new { success = false, message = "File not found." });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in RemovePdfFile: {ex}");
                return Json(new { success = false, message = $"Error removing file: {ex.Message}" });
            }
        }

        [HttpPost]
        public async Task<IActionResult> MergeUploadedPdfs([FromBody] MergePdfRequest request)
        {
            try
            {
                Console.WriteLine($"=== MergeUploadedPdfs Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");

                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);
                
                if (!Directory.Exists(uploadDirectory))
                {
                    return Json(new { success = false, message = "No uploaded PDF files found." });
                }

                var pdfFiles = Directory.GetFiles(uploadDirectory, "*.pdf").OrderBy(f => f).ToList();
                
                if (pdfFiles.Count == 0)
                {
                    return Json(new { success = false, message = "No PDF files to merge." });
                }

                Console.WriteLine($"Found {pdfFiles.Count} PDF files to merge");

                // Create merged PDF
                var mergedFileName = $"merged_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                var mergedFilePath = Path.Combine(uploadDirectory, mergedFileName);

                using (var writer = new PdfWriter(mergedFilePath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    foreach (var pdfFile in pdfFiles)
                    {
                        try
                        {
                            using (var reader = new PdfReader(pdfFile))
                            using (var sourcePdf = new PdfDocument(reader))
                            {
                                sourcePdf.CopyPagesTo(1, sourcePdf.GetNumberOfPages(), mergedPdf);
                                Console.WriteLine($"✅ Added: {Path.GetFileName(pdfFile)} ({sourcePdf.GetNumberOfPages()} pages)");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Error adding {pdfFile}: {ex.Message}");
                            continue;
                        }
                    }

                    mergedPdf.Close();
                }

                var fileInfo = new FileInfo(mergedFilePath);
                var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();

                // Get page count
                int totalPages = 0;
                using (var reader = new PdfReader(mergedFilePath))
                using (var pdfDoc = new PdfDocument(reader))
                {
                    totalPages = pdfDoc.GetNumberOfPages();
                }

                mergedPdfs.Add(new MergedPdfInfo
                {
                    FileName = mergedFileName,
                    FilePath = mergedFilePath,
                    FileSize = fileInfo.Length,
                    TotalPages = totalPages,
                    CreatedAt = DateTime.Now
                });

                HttpContext.Session.SetString(request.SessionId + "_mergedPdfs", JsonSerializer.Serialize(mergedPdfs));

                return Json(new
                {
                    success = true,
                    message = $"{pdfFiles.Count} PDF files merged successfully.",
                    fileName = mergedFileName,
                    fileSize = fileInfo.Length,
                    totalPages = totalPages
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in MergeUploadedPdfs: {ex}");
                return Json(new { success = false, message = $"Error merging PDFs: {ex.Message}" });
            }
        }

        [HttpPost]
        public IActionResult DownloadMergedPdf([FromBody] DownloadPdfRequest request)
        {
            try
            {
                Console.WriteLine($"=== DownloadMergedPdf Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"File name: {request.FileName}");

                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);
                var filePath = Path.Combine(uploadDirectory, request.FileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return Json(new { success = false, message = "File not found." });
                }

                var pdfBytes = System.IO.File.ReadAllBytes(filePath);
                var pdfBase64 = Convert.ToBase64String(pdfBytes);

                return Json(new
                {
                    success = true,
                    pdfData = pdfBase64,
                    fileName = request.FileName
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in DownloadMergedPdf: {ex}");
                return Json(new { success = false, message = $"Error downloading PDF: {ex.Message}" });
            }
        }

        [HttpPost]
        public IActionResult RemoveMergedPdf([FromBody] RemovePdfRequest request)
        {
            try
            {
                Console.WriteLine($"=== RemoveMergedPdf Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"File name: {request.FileName}");

                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);
                var filePath = Path.Combine(uploadDirectory, request.FileName);

                if (System.IO.File.Exists(filePath))
                {
                        System.IO.File.Delete(filePath);
                    Console.WriteLine($"🗑️ Deleted merged PDF: {filePath}");
                }

                // Remove from session
                var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();
                mergedPdfs.RemoveAll(m => m.FileName == request.FileName);
                HttpContext.Session.SetString(request.SessionId + "_mergedPdfs", JsonSerializer.Serialize(mergedPdfs));

                return Json(new { success = true, message = "Merged PDF removed successfully." });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in RemoveMergedPdf: {ex}");
                return Json(new { success = false, message = $"Error removing PDF: {ex.Message}" });
            }
        }















        [HttpPost]
        public async Task<IActionResult> GeneratePdfPreviewWithFitToPage([FromBody] PdfPreviewWithFitToPageRequest request)
        {
            try
            {
                Console.WriteLine($"=== GeneratePdfPreviewWithFitToPage Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"Selected Sheets: {request.SelectedSheets?.Count ?? 0}");

                // ✅ DEBUG: Check what visibility data we received
                if (request.PageOrderData != null)
                {
                    int totalPages = request.PageOrderData.Count;
                    int visiblePages = request.PageOrderData.Count(p => p.Visible);
                    Console.WriteLine($"📊 Received {totalPages} pages, {visiblePages} visible, {totalPages - visiblePages} hidden");
                }

                var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");
                var originalFileName = HttpContext.Session.GetString(request.SessionId + "_fileName");

                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    return Json(new { success = false, message = "File not found." });
                }

                // Step 1: Convert Excel to PDF (NO processing, keep original)
                var outputFileName = $"preview_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                var result = await ConvertToPdfWithColorPreservation(
                    filePath,  // Use original file
                    outputFileName,
                    request.SelectedSheets);

                if (!result.Success || !System.IO.File.Exists(result.PdfFilePath))
                {
                    return Json(new { success = false, message = result.Message });
                }

                Console.WriteLine($"✅ PDF created: {result.PdfFilePath}");



                // Convert from Models.PageOrderInfoWithRotation to Controllers.PageOrderInfoWithRotation
                var controllerPageOrderData = request.PageOrderData?.Select(p => new PageOrderInfoWithRotation
                {
                    OriginalPage = p.OriginalPage,
                    CurrentOrder = p.CurrentOrder,
                    Visible = p.Visible,
                    Orientation = p.Orientation,
                    Rotation = p.Rotation
                }).ToList();

                // Step 2: ✅ Apply ONLY FitToPage with current visibility/orientation
                string finalPdfPath = await ApplyOnlyFitToPage(
                    result.PdfFilePath,
                    controllerPageOrderData);

                Console.WriteLine($"✅ FitToPage applied: {finalPdfPath}");

                // Step 3: Check for merged PDFs
                string finalPathWithMerged = finalPdfPath;
                var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();

                if (mergedPdfs.Any())
                {
                    var latestMergedPdf = mergedPdfs.OrderByDescending(m => m.CreatedAt).FirstOrDefault();
                    if (latestMergedPdf != null && System.IO.File.Exists(latestMergedPdf.FilePath))
                    {
                        finalPathWithMerged = await MergeExcelWithUploadedPdfs(finalPdfPath, latestMergedPdf.FilePath);
                        Console.WriteLine($"✅ Merged with uploaded PDFs");
                    }
                }

                // Step 4: Read and return PDF
                var pdfBytes = await System.IO.File.ReadAllBytesAsync(finalPathWithMerged);
                var pdfBase64 = Convert.ToBase64String(pdfBytes);

                // Get final page count
                int totalPagesFinal = 0;
                try
                {
                    using (var reader = new PdfReader(finalPathWithMerged))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        totalPagesFinal = pdfDoc.GetNumberOfPages();
                        Console.WriteLine($"📊 Final Preview PDF: {totalPagesFinal} pages");
                    }
                }
                catch
                {
                    totalPagesFinal = 1;
                }

                // Step 5: Cleanup
                System.IO.File.Delete(result.PdfFilePath);
                if (System.IO.File.Exists(finalPdfPath))
                    System.IO.File.Delete(finalPdfPath);
                if (finalPathWithMerged != finalPdfPath && System.IO.File.Exists(finalPathWithMerged))
                    System.IO.File.Delete(finalPathWithMerged);

                return Json(new
                {
                    success = true,
                    pdfData = pdfBase64,
                    fileName = outputFileName,
                    totalPages = totalPagesFinal
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GeneratePdfPreviewWithFitToPage: {ex}");
                return Json(new { success = false, message = $"Preview generation failed: {ex.Message}" });
            }
        }

        // ✅ NEW: Apply ONLY FitToPage (keep existing orientation/rotation)
        private async Task<string> ApplyOnlyFitToPage(
            string pdfPath,
            List<PageOrderInfoWithRotation> pageOrderData)
        {
            var outputPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"preview_fittopage_only_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Applying ONLY FitToPage...");

                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(outputPath))
                using (var newPdf = new PdfDocument(writer))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    int sourcePageCount = sourcePdf.GetNumberOfPages();
                    Console.WriteLine($"📄 Source PDF: {sourcePageCount} pages");

                    // Use provided pageOrderData or create default
                    List<PageOrderInfoWithRotation> visiblePages;

                    if (pageOrderData != null && pageOrderData.Any())
                    {
                        visiblePages = pageOrderData
                            .Where(p => p.Visible)
                            .OrderBy(p => p.CurrentOrder)
                            .ToList();
                        Console.WriteLine($"📄 Using provided visibility: {visiblePages.Count} visible pages");
                    }
                    else
                    {
                        // Default: all pages visible
                        visiblePages = new List<PageOrderInfoWithRotation>();
                        for (int i = 1; i <= sourcePageCount; i++)
                        {
                            visiblePages.Add(new PageOrderInfoWithRotation
                            {
                                OriginalPage = i,
                                CurrentOrder = i,
                                Visible = true,
                                Orientation = "portrait",
                                Rotation = 0
                            });
                        }
                        Console.WriteLine($"📄 Using default: all {sourcePageCount} pages visible");
                    }

                    // Process each page
                    foreach (var pageInfo in visiblePages)
                    {
                        int sourcePageNum = pageInfo.OriginalPage;

                        if (sourcePageNum > 0 && sourcePageNum <= sourcePageCount)
                        {
                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
                            var sourcePageSize = sourcePage.GetPageSize();

                            // ✅ USE EXISTING ORIENTATION from pageOrderData
                            string orientation = pageInfo.Orientation ?? "portrait";
                            PageSize targetPageSize = orientation == "landscape" ? PageSize.A4.Rotate() : PageSize.A4;

                            // ✅ USE EXISTING ROTATION from pageOrderData
                            int rotation = pageInfo.Rotation;

                            // Create new page
                            var newPage = newPdf.AddNewPage(targetPageSize);
                            var copiedPage = sourcePage.CopyAsFormXObject(newPdf);
                            var canvas = new PdfCanvas(newPage);

                            // ✅ FIT TO PAGE CALCULATION
                            float margin = 20;
                            float targetWidth = targetPageSize.GetWidth();
                            float targetHeight = targetPageSize.GetHeight();
                            float availableWidth = targetWidth - (2 * margin);
                            float availableHeight = targetHeight - (2 * margin);

                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            // Calculate scaling
                            float scaleX = availableWidth / sourceWidth;
                            float scaleY = availableHeight / sourceHeight;
                            float scale = Math.Min(scaleX, scaleY);

                            // Adjust for existing rotation
                            if (rotation != 0)
                            {
                                float rotatedWidth = Math.Abs(sourceWidth * (float)Math.Cos(rotation * Math.PI / 180)) +
                                                   Math.Abs(sourceHeight * (float)Math.Sin(rotation * Math.PI / 180));
                                float rotatedHeight = Math.Abs(sourceWidth * (float)Math.Sin(rotation * Math.PI / 180)) +
                                                    Math.Abs(sourceHeight * (float)Math.Cos(rotation * Math.PI / 180));

                                float rotatedScaleX = availableWidth / rotatedWidth;
                                float rotatedScaleY = availableHeight / rotatedHeight;
                                scale = Math.Min(rotatedScaleX, rotatedScaleY);
                            }

                            // Calculate centered position
                            float scaledWidth = sourceWidth * scale;
                            float scaledHeight = sourceHeight * scale;
                            float xOffset = margin + (availableWidth - scaledWidth) / 2;
                            float yOffset = margin + (availableHeight - scaledHeight) / 2;

                            // Apply transformations (keeping existing rotation)
                            canvas.SaveState();

                            // Move to center
                            canvas.ConcatMatrix(1, 0, 0, 1, xOffset + scaledWidth / 2, yOffset + scaledHeight / 2);

                            // Apply existing rotation
                            if (rotation != 0)
                            {
                                canvas.ConcatMatrix((float)Math.Cos(rotation * Math.PI / 180),
                                                    (float)Math.Sin(rotation * Math.PI / 180),
                                                    (float)-Math.Sin(rotation * Math.PI / 180),
                                                    (float)Math.Cos(rotation * Math.PI / 180),
                                                    0, 0);
                            }

                            // Move back and apply scaling
                            canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
                            canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);

                            // Draw content
                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();
                            canvas.Release();

                            Console.WriteLine($"✅ Page {sourcePageNum}: FitToPage applied (Orientation: {orientation}, Rotation: {rotation}°)");
                        }
                    }

                    newPdf.Close();
                    sourcePdf.Close();

                    Console.WriteLine($"✅ Preview with FitToPage created: {outputPath} ({visiblePages.Count} pages)");
                }

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplyOnlyFitToPage: {ex}");
                return pdfPath; // Return original if fails
            }
        }








        [HttpPost]
        public async Task<IActionResult> GeneratePdfPreview([FromBody] PdfPreviewRequest request)
        {
            try
            {
                Console.WriteLine($"=== GeneratePdfPreview Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"Selected Sheets: {string.Join(", ", request.SelectedSheets ?? new List<string>())}");

                var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");
                var originalFileName = HttpContext.Session.GetString(request.SessionId + "_fileName");

                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    return Json(new { success = false, message = "File not found. Please upload again." });
                }

                // Step 1: Process Excel file with color preservation
                string processedFilePath;
                try
                {
                    processedFilePath = await ProcessExcelWithColorPreservation(filePath, request.SelectedSheets);
                    Console.WriteLine($"✅ Excel file processed with color preservation: {processedFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Excel processing failed: {ex.Message}");
                    return Json(new { success = false, message = $"Excel processing failed: {ex.Message}" });
                }

                // Step 2: Convert to PDF
                var outputFileName = $"preview_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                var result = await ConvertToPdfWithColorPreservation(processedFilePath, outputFileName, request.SelectedSheets);

                // Cleanup processed file
                if (processedFilePath != filePath && System.IO.File.Exists(processedFilePath))
                {
                    _excelProcessingService.CleanupProcessedFile(processedFilePath);
                }

                if (result.Success && System.IO.File.Exists(result.PdfFilePath))
                {
                    // Check if we need to merge with uploaded PDFs
                    var finalPdfPath = result.PdfFilePath;
                    
                    // Get merged PDFs from session
                    var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                    var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();
                    
                    if (mergedPdfs.Any())
                    {
                        // Get the latest merged PDF
                        var latestMergedPdf = mergedPdfs.OrderByDescending(m => m.CreatedAt).FirstOrDefault();
                        if (latestMergedPdf != null && System.IO.File.Exists(latestMergedPdf.FilePath))
                        {
                            // Merge Excel PDF with uploaded merged PDF
                            finalPdfPath = await MergeExcelWithUploadedPdfs(result.PdfFilePath, latestMergedPdf.FilePath);
                            Console.WriteLine($"✅ Merged with uploaded PDFs: {finalPdfPath}");
                        }
                    }

                    var pdfBytes = await System.IO.File.ReadAllBytesAsync(finalPdfPath);
                    var pdfBase64 = Convert.ToBase64String(pdfBytes);

                    // Get page count
                    int totalPages = 0;
                    try
                    {
                        using (var reader = new PdfReader(finalPdfPath))
                        using (var pdfDoc = new PdfDocument(reader))
                        {
                            totalPages = pdfDoc.GetNumberOfPages();
                        }
                    }
                    catch
                    {
                        totalPages = 1;
                    }

                    // Cleanup temporary files
                    System.IO.File.Delete(result.PdfFilePath);
                    if (finalPdfPath != result.PdfFilePath && System.IO.File.Exists(finalPdfPath))
                    {
                        System.IO.File.Delete(finalPdfPath);
                    }
                   
                    return Json(new
                    {
                        success = true,
                        pdfData = pdfBase64,
                        fileName = outputFileName,
                        generatedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        totalPages = totalPages
                    });
                }
                else
                {
                    return Json(new { success = false, message = result.Message });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GeneratePdfPreview: {ex}");
                return Json(new { success = false, message = $"Preview generation failed: {ex.Message}" });
            }
        }

        private async Task<string> MergeExcelWithUploadedPdfs(string excelPdfPath, string uploadedPdfPath)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"preview_merged_{Guid.NewGuid()}.pdf");

            try
            {
                using (var writer = new PdfWriter(outputPath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    // Add Excel PDF first
                    if (System.IO.File.Exists(excelPdfPath))
                    {
                        using (var reader = new PdfReader(excelPdfPath))
                        using (var sourcePdf = new PdfDocument(reader))
                        {
                            sourcePdf.CopyPagesTo(1, sourcePdf.GetNumberOfPages(), mergedPdf);
                            Console.WriteLine($"✅ Added Excel PDF to preview: {excelPdfPath}");
                        }
                    }

                    // Add uploaded merged PDF
                    if (System.IO.File.Exists(uploadedPdfPath))
                    {
                        using (var reader = new PdfReader(uploadedPdfPath))
                        using (var sourcePdf = new PdfDocument(reader))
                        {
                            sourcePdf.CopyPagesTo(1, sourcePdf.GetNumberOfPages(), mergedPdf);
                            Console.WriteLine($"✅ Added uploaded PDF to preview: {uploadedPdfPath}");
                        }
                    }

                    mergedPdf.Close();
                }

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in MergeExcelWithUploadedPdfs: {ex}");
                return excelPdfPath; // Return original if merge fails
            }
        }

        private async Task<string> ProcessExcelWithColorPreservation(string inputFilePath, List<string> selectedSheets)
        {
            var outputFilePath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"color_preserved_{Guid.NewGuid()}.xlsx");

            try
            {
                using (var sourcePackage = new OfficeOpenXml.ExcelPackage(new FileInfo(inputFilePath)))
                using (var targetPackage = new OfficeOpenXml.ExcelPackage())
                {
                    var sourceWorkbook = sourcePackage.Workbook;
                    var targetWorkbook = targetPackage.Workbook;

                    // Add sheets in order
                    foreach (var sheetName in selectedSheets)
                    {
                        var sourceWorksheet = sourceWorkbook.Worksheets[sheetName];
                        if (sourceWorksheet != null)
                        {
                            var targetWorksheet = targetWorkbook.Worksheets.Add(sheetName, sourceWorksheet);
                            Console.WriteLine($"✅ Copied sheet with full formatting: {sheetName}");
                        }
                    }

                    targetPackage.SaveAs(new FileInfo(outputFilePath));
                }

                return outputFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ProcessExcelWithColorPreservation: {ex}");
                return await CreateSimpleCopy(inputFilePath, selectedSheets);
            }
        }

        private async Task<string> CreateSimpleCopy(string inputFilePath, List<string> selectedSheets)
        {
            var outputFilePath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"simple_copy_{Guid.NewGuid()}.xlsx");

            try
            {
                using (var sourcePackage = new OfficeOpenXml.ExcelPackage(new FileInfo(inputFilePath)))
                using (var targetPackage = new OfficeOpenXml.ExcelPackage())
                {
                    var sourceWorkbook = sourcePackage.Workbook;
                    var targetWorkbook = targetPackage.Workbook;

                    foreach (var sheetName in selectedSheets)
                    {
                        var sourceWorksheet = sourceWorkbook.Worksheets[sheetName];
                        if (sourceWorksheet != null)
                        {
                            var targetWorksheet = targetWorkbook.Worksheets.Add(sheetName);

                            // Copy data only (no formatting)
                            if (sourceWorksheet.Dimension != null)
                            {
                                int maxRows = Math.Min(sourceWorksheet.Dimension.End.Row, 1000);
                                int maxCols = Math.Min(sourceWorksheet.Dimension.End.Column, 100);

                                for (int row = 1; row <= maxRows; row++)
                                {
                                    for (int col = 1; col <= maxCols; col++)
                                    {
                                        targetWorksheet.Cells[row, col].Value = sourceWorksheet.Cells[row, col].Value;
                                    }
                                }
                            }

                            Console.WriteLine($"✅ Created simple copy of sheet: {sheetName}");
                        }
                    }

                    targetPackage.SaveAs(new FileInfo(outputFilePath));
                }

                return outputFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in CreateSimpleCopy: {ex}");
                throw;
            }
        }

        private async Task<ConversionResult> ConvertToPdfWithColorPreservation(
            string inputFilePath, string outputFileName, List<string> selectedSheets)
        {
            var outputDirectory = System.IO.Path.Combine(_environment.WebRootPath, "App_Data", "Previews");
            Directory.CreateDirectory(outputDirectory);
            var outputFilePath = System.IO.Path.Combine(outputDirectory, outputFileName);

            try
            {
                var arguments = BuildLibreOfficeArguments(inputFilePath, outputDirectory, selectedSheets);
                Console.WriteLine($"LibreOffice arguments: {arguments}");

                var processStartInfo = new ProcessStartInfo
                {
                    FileName = GetLibreOfficePath(),
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = outputDirectory
                };

                using (var process = new Process())
                {
                    process.StartInfo = processStartInfo;
                    process.Start();

                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();

                    bool processExited = process.WaitForExit(180000); // 3 minutes

                    if (processExited && process.ExitCode == 0)
                    {
                        var inputFileName = System.IO.Path.GetFileNameWithoutExtension(inputFilePath);
                        var possibleOutputPaths = new[]
                        {
                            System.IO.Path.Combine(outputDirectory, inputFileName + ".pdf"),
                            System.IO.Path.Combine(outputDirectory, outputFileName)
                        };

                        foreach (var path in possibleOutputPaths)
                        {
                            if (System.IO.File.Exists(path))
                            {
                                return new ConversionResult
                                {
                                    Success = true,
                                    Message = "Conversion successful",
                                    PdfFilePath = path,
                                    FileName = outputFileName
                                };
                            }
                        }
                    }

                    return new ConversionResult
                    {
                        Success = false,
                        Message = $"Conversion failed. Exit code: {process.ExitCode}, Error: {error}"
                    };
                }
            }
            catch (Exception ex)
            {
                return new ConversionResult
                {
                    Success = false,
                    Message = $"Error during conversion: {ex.Message}"
                };
            }
        }

        private string BuildLibreOfficeArguments(string inputFilePath, string outputDirectory, List<string> selectedSheets)
        {
            var arguments = new List<string>
            {
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                "--convert-to pdf:calc_pdf_Export",
                $"--outdir \"{outputDirectory}\"",
                $"\"{inputFilePath}\""
            };

            return string.Join(" ", arguments);
        }

        private string GetLibreOfficePath()
        {
            string[] possiblePaths = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                @"C:\Program Files\LibreOffice\program\soffice.com",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.com"
            };

            foreach (var path in possiblePaths)
            {
                if (System.IO.File.Exists(path))
                {
                    Console.WriteLine($"LibreOffice found at: {path}");
                    return path;
                }
            }

            throw new Exception("LibreOffice not found. Please install LibreOffice from https://www.libreoffice.org/download/download-libreoffice/");
        }

        // New Request Model with Rotation
        public class PdfRequestWithRotation
        {
            public string SessionId { get; set; }
            public List<string> SelectedSheets { get; set; }
            public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
            public Dictionary<int, string> OrientationData { get; set; }
            public Dictionary<int, int> RotationData { get; set; }
        }

        public class PageOrderInfoWithRotation
        {
            public int OriginalPage { get; set; }
            public int CurrentOrder { get; set; }
            public bool Visible { get; set; }
            public string Orientation { get; set; } = "portrait";
            public int Rotation { get; set; } = 0;
        }

        [HttpPost]
        public async Task<IActionResult> GenerateReorderedPdf([FromBody] PdfRequestWithRotation request)
        {
            try
            {
                Console.WriteLine($"=== GenerateReorderedPdf Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"Selected Sheets: {request.SelectedSheets?.Count ?? 0}");
                Console.WriteLine($"Page Order Data: {request.PageOrderData?.Count ?? 0} pages");
                Console.WriteLine($"Rotation Data: {request.RotationData?.Count ?? 0} rotations");

                var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");
                var originalFileName = HttpContext.Session.GetString(request.SessionId + "_fileName");

                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    return Json(new { success = false, message = "File not found. Please upload again." });
                }

                // Step 1: Get selected sheets
                var selectedSheets = request.SelectedSheets ?? new List<string>();

                // Step 2: Create initial PDF
                var outputFileName = $"document_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                var conversionResult = await ConvertToPdfWithColorPreservation(
                    filePath, outputFileName, selectedSheets);

                if (!conversionResult.Success || !System.IO.File.Exists(conversionResult.PdfFilePath))
                {
                    return Json(new { success = false, message = conversionResult.Message });
                }

                Console.WriteLine($"✅ Initial PDF created: {conversionResult.PdfFilePath}");

                // Step 3: Apply reordering, orientation and rotation
                string finalPdfPath = await ApplyPageReorderingOrientationAndRotation(
                    conversionResult.PdfFilePath,
                    request.PageOrderData,
                    request.OrientationData,
                    request.RotationData);

                // Step 4: Check for merged PDFs and combine
                var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();
                
                if (mergedPdfs.Any())
                {
                    // Get the latest merged PDF
                    var latestMergedPdf = mergedPdfs.OrderByDescending(m => m.CreatedAt).FirstOrDefault();
                    if (latestMergedPdf != null && System.IO.File.Exists(latestMergedPdf.FilePath))
                    {
                        // Merge the Excel PDF with uploaded merged PDF
                        var combinedPdfPath = await MergeExcelWithUploadedPdfs(finalPdfPath, latestMergedPdf.FilePath);
                        System.IO.File.Delete(finalPdfPath);
                        finalPdfPath = combinedPdfPath;
                        Console.WriteLine($"✅ Combined with uploaded PDFs: {finalPdfPath}");
                    }
                }

                // Step 5: Read final PDF
                var finalPdfBytes = await System.IO.File.ReadAllBytesAsync(finalPdfPath);
                var finalPdfBase64 = Convert.ToBase64String(finalPdfBytes);

                // Step 6: Cleanup
                if (System.IO.File.Exists(conversionResult.PdfFilePath))
                    System.IO.File.Delete(conversionResult.PdfFilePath);
                if (System.IO.File.Exists(finalPdfPath))
                    System.IO.File.Delete(finalPdfPath);

                return Json(new
                {
                    success = true,
                    pdfData = finalPdfBase64,
                    fileName = outputFileName,
                    message = "PDF generated successfully with rotations"
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GenerateReorderedPdf: {ex}");
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }


        private async Task<string> ApplyPageReorderingOrientationAndRotation(
    string pdfPath,
    List<PageOrderInfoWithRotation> pageOrderData,
    Dictionary<int, string> orientationData,
    Dictionary<int, int> rotationData)
        {
            var outputPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"final_with_orientation_rotation_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Applying page reordering, orientation and rotation...");

                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(outputPath))
                using (var newPdf = new PdfDocument(writer))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    // Filter visible pages and sort by current order
                    var visiblePages = pageOrderData
                        .Where(p => p.Visible)
                        .OrderBy(p => p.CurrentOrder)
                        .ToList();

                    Console.WriteLine($"📄 Processing {visiblePages.Count} visible pages");

                    foreach (var pageInfo in visiblePages)
                    {
                        int sourcePageNum = pageInfo.OriginalPage;

                        if (sourcePageNum > 0 && sourcePageNum <= sourcePdf.GetNumberOfPages())
                        {
                            // Get the source page
                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
                            var sourcePageSize = sourcePage.GetPageSize();

                            // Get orientation for this page
                            string orientation = pageInfo.Orientation;
                            if (orientationData != null && orientationData.ContainsKey(sourcePageNum))
                            {
                                orientation = orientationData[sourcePageNum];
                            }

                            // Get rotation for this page
                            int rotation = pageInfo.Rotation;
                            if (rotationData != null && rotationData.ContainsKey(sourcePageNum))
                            {
                                rotation = rotationData[sourcePageNum];
                            }

                            // Create page with appropriate orientation
                            PageSize targetPageSize = orientation == "landscape" ? PageSize.A4.Rotate() : PageSize.A4;

                            // Create new page
                            var newPage = newPdf.AddNewPage(targetPageSize);

                            // Copy content
                            var copiedPage = sourcePage.CopyAsFormXObject(newPdf);
                            var canvas = new PdfCanvas(newPage);

                            // Get page dimensions
                            float targetWidth = targetPageSize.GetWidth();
                            float targetHeight = targetPageSize.GetHeight();
                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            // Calculate scaling to fit page (with margins)
                            float margin = 20; // 20 points margin
                            float availableWidth = targetWidth - (2 * margin);
                            float availableHeight = targetHeight - (2 * margin);

                            // Calculate scale without rotation
                            float scaleX = availableWidth / sourceWidth;
                            float scaleY = availableHeight / sourceHeight;
                            float scale = Math.Min(scaleX, scaleY);

                            // Adjust for rotation
                            if (rotation != 0)
                            {
                                // When rotated, the bounding box changes
                                float rotatedWidth = Math.Abs(sourceWidth * (float)Math.Cos(rotation * Math.PI / 180)) +
                                                   Math.Abs(sourceHeight * (float)Math.Sin(rotation * Math.PI / 180));
                                float rotatedHeight = Math.Abs(sourceWidth * (float)Math.Sin(rotation * Math.PI / 180)) +
                                                    Math.Abs(sourceHeight * (float)Math.Cos(rotation * Math.PI / 180));

                                float rotatedScaleX = availableWidth / rotatedWidth;
                                float rotatedScaleY = availableHeight / rotatedHeight;
                                scale = Math.Min(rotatedScaleX, rotatedScaleY);
                            }

                            // Calculate centered position
                            float scaledWidth = sourceWidth * scale;
                            float scaledHeight = sourceHeight * scale;
                            float xOffset = margin + (availableWidth - scaledWidth) / 2;
                            float yOffset = margin + (availableHeight - scaledHeight) / 2;

                            // Apply transformations
                            canvas.SaveState();

                            // Move to center of scaled content
                            canvas.ConcatMatrix(1, 0, 0, 1, xOffset + scaledWidth / 2, yOffset + scaledHeight / 2);

                            // Apply rotation
                            canvas.ConcatMatrix((float)Math.Cos(rotation * Math.PI / 180),
                                                (float)Math.Sin(rotation * Math.PI / 180),
                                                (float)-Math.Sin(rotation * Math.PI / 180),
                                                (float)Math.Cos(rotation * Math.PI / 180),
                                                0, 0);

                            // Move back and apply scaling
                            canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
                            canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);

                            // Draw the content
                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();

                            canvas.Release();

                            Console.WriteLine($"✅ Page {sourcePageNum} → Position {pageInfo.CurrentOrder} (Orientation: {orientation}, Rotation: {rotation}°)");
                        }
                        else
                        {
                            Console.WriteLine($"⚠️ Page {sourcePageNum} not found in source PDF");
                        }
                    }

                    newPdf.Close();
                    sourcePdf.Close();
                }

                Console.WriteLine($"✅ PDF with orientation and rotation created: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplyPageReorderingOrientationAndRotation: {ex}");
                throw;
            }
        }


        [HttpPost]
        public async Task<IActionResult> ConvertToPdf(string sessionId, List<string> selectedSheets,
            List<int> sheetOrders, List<string> sheetOrientations)
        {
            try
            {
                Console.WriteLine($"=== ConvertToPdf Called ===");
                Console.WriteLine($"Session ID: {sessionId}");
                Console.WriteLine($"Selected Sheets: {string.Join(", ", selectedSheets ?? new List<string>())}");

                var filePath = HttpContext.Session.GetString(sessionId + "_filePath");
                var originalFileName = HttpContext.Session.GetString(sessionId + "_fileName");

                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    return Json(new { success = false, message = "File not found. Please upload again." });
                }

                var orderedSheets = new List<string>();
                if (selectedSheets != null && sheetOrders != null && selectedSheets.Count == sheetOrders.Count)
                {
                    var sheetOrderMap = selectedSheets.Zip(sheetOrders, (s, o) => new { Sheet = s, Order = o })
                                                    .OrderBy(x => x.Order)
                                                    .Select(x => x.Sheet)
                                                    .ToList();
                    orderedSheets = sheetOrderMap;
                    Console.WriteLine($"✅ Ordered sheets (drag & drop order): {string.Join(" → ", orderedSheets)}");
                }
                else
                {
                    orderedSheets = selectedSheets ?? new List<string>();
                    Console.WriteLine($"ℹ️ Using default sheet order: {string.Join(" → ", orderedSheets)}");
                }

                Console.WriteLine($"🔄 Starting Excel file processing...");
                string processedFilePath;
                try
                {
                    processedFilePath = await _excelProcessingService.CreateProcessedExcelFileAsync(
                        filePath, orderedSheets, orderedSheets);
                    Console.WriteLine($"✅ Excel file processed successfully: {processedFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Excel processing failed, using original file: {ex.Message}");
                    processedFilePath = filePath;
                }

                var orientationMap = new Dictionary<string, string>();
                if (selectedSheets != null && sheetOrientations != null && selectedSheets.Count == sheetOrientations.Count)
                {
                    orientationMap = selectedSheets.Zip(sheetOrientations, (s, o) => new { Sheet = s, Orientation = o })
                                                 .ToDictionary(x => x.Sheet, x => x.Orientation);
                    Console.WriteLine($"✅ Orientation mapping: {string.Join(", ", orientationMap)}");
                }

                var outputFileName = System.IO.Path.GetFileNameWithoutExtension(originalFileName ?? "converted") + ".pdf";

                Console.WriteLine($"🔄 Starting Excel to PDF conversion...");
                var result = await _libreOfficeService.ConvertToPdfAsync(
                    processedFilePath, outputFileName, orderedSheets, orientationMap);

                if (processedFilePath != filePath)
                {
                    _excelProcessingService.CleanupProcessedFile(processedFilePath);
                }

                if (result.Success)
                {
                    Console.WriteLine($"✅ Excel to PDF conversion successful: {result.PdfFilePath}");

                    // Check for merged PDFs and combine
                    var mergedPdfsJson = HttpContext.Session.GetString(sessionId + "_mergedPdfs") ?? "[]";
                    var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();
                    
                    if (mergedPdfs.Any())
                    {
                        // Get the latest merged PDF
                        var latestMergedPdf = mergedPdfs.OrderByDescending(m => m.CreatedAt).FirstOrDefault();
                        if (latestMergedPdf != null && System.IO.File.Exists(latestMergedPdf.FilePath))
                        {
                            // Merge the Excel PDF with uploaded merged PDF
                            var combinedPdfPath = await MergeExcelWithUploadedPdfs(result.PdfFilePath, latestMergedPdf.FilePath);
                            System.IO.File.Delete(result.PdfFilePath);
                            result.PdfFilePath = combinedPdfPath;
                            result.FileName = $"merged_{result.FileName}";
                            Console.WriteLine($"✅ Combined with uploaded PDFs: {combinedPdfPath}");
                        }
                    }

                    Console.WriteLine($"🔄 Starting PDF merge process...");
                    var finalResult = await MergeAllPdfsWithiText7(result.PdfFilePath, outputFileName, orderedSheets, orientationMap);

                    HttpContext.Session.Remove(sessionId + "_filePath");
                    HttpContext.Session.Remove(sessionId + "_fileName");
                    HttpContext.Session.Remove(sessionId + "_uploadedPdfs");
                    HttpContext.Session.Remove(sessionId + "_mergedPdfs");

                    if (finalResult.Success)
                    {
                        Console.WriteLine($"✅ Final PDF created: {finalResult.PdfFilePath} ({finalResult.TotalPages} pages)");

                        var fileBytes = await System.IO.File.ReadAllBytesAsync(finalResult.PdfFilePath);

                        if (System.IO.File.Exists(result.PdfFilePath))
                            System.IO.File.Delete(result.PdfFilePath);
                        if (System.IO.File.Exists(finalResult.PdfFilePath))
                            System.IO.File.Delete(finalResult.PdfFilePath);

                        Console.WriteLine($"✅ Returning final PDF: {finalResult.FileName}");
                        return File(fileBytes, "application/pdf", finalResult.FileName);
                    }
                    else
                    {
                        Console.WriteLine($"❌ PDF merge failed, returning original PDF");
                        var fileBytes = await System.IO.File.ReadAllBytesAsync(result.PdfFilePath);
                        System.IO.File.Delete(result.PdfFilePath);
                        return File(fileBytes, "application/pdf", result.FileName);
                    }
                }
                else
                {
                    Console.WriteLine($"❌ Excel to PDF conversion failed: {result.Message}");
                    return Json(new { success = false, message = result.Message });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Exception in ConvertToPdf: {ex}");
                return Json(new { success = false, message = $"Conversion failed: {ex.Message}" });
            }
        }

        private async Task<ConversionResult> MergeAllPdfsWithiText7(string newPdfPath, string outputFileName,
            List<string> orderedSheets, Dictionary<string, string> orientationMap)
        {
            var result = new ConversionResult();
            var mergedPdfPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"final_merged_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Starting PDF merge process with iText7...");

                using (var writer = new PdfWriter(mergedPdfPath))
                using (var mergedPdfDoc = new PdfDocument(writer))
                {
                    if (System.IO.File.Exists(newPdfPath))
                    {
                        using (var reader = new PdfReader(newPdfPath))
                        using (var sourceDoc = new PdfDocument(reader))
                        {
                            sourceDoc.CopyPagesTo(1, sourceDoc.GetNumberOfPages(), mergedPdfDoc);
                            Console.WriteLine($"✅ Added converted Excel PDF: {newPdfPath} ({sourceDoc.GetNumberOfPages()} pages)");
                        }
                    }

                    var pdfDirectory = @"D:\CIPL\SinghAndSons\pdf";
                    if (System.IO.Directory.Exists(pdfDirectory))
                    {
                        var existingPdfFiles = System.IO.Directory.GetFiles(pdfDirectory, "*.pdf")
                            .OrderBy(f => f)
                            .ToList();

                        Console.WriteLine($"📁 Found {existingPdfFiles.Count} existing PDF files to merge");

                        foreach (var existingPdf in existingPdfFiles)
                        {
                            try
                            {
                                using (var reader = new PdfReader(existingPdf))
                                using (var sourceDoc = new PdfDocument(reader))
                                {
                                    sourceDoc.CopyPagesTo(1, sourceDoc.GetNumberOfPages(), mergedPdfDoc);
                                    Console.WriteLine($"✅ Added existing PDF: {System.IO.Path.GetFileName(existingPdf)} ({sourceDoc.GetNumberOfPages()} pages)");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"❌ Error adding existing PDF {existingPdf}: {ex.Message}");
                                continue;
                            }
                        }
                    }

                    mergedPdfDoc.Close();
                }

                var fileInfo = new System.IO.FileInfo(mergedPdfPath);
                if (fileInfo.Exists && fileInfo.Length > 0)
                {
                    result.Success = true;
                    result.PdfFilePath = mergedPdfPath;
                    result.FileName = $"merged_{System.IO.Path.GetFileNameWithoutExtension(outputFileName)}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                    result.TotalPages = await GetPageCount(mergedPdfPath);
                    result.Message = $"Successfully merged PDF with {result.TotalPages} total pages";
                    Console.WriteLine($"✅ Final merged PDF created: {result.PdfFilePath} ({result.TotalPages} pages)");
                }
                else
                {
                    throw new Exception("Merged PDF file was not created or is empty");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in MergeAllPdfsWithiText7: {ex}");
                result.Success = false;
                result.Message = $"PDF merge failed: {ex.Message}";

                if (System.IO.File.Exists(newPdfPath))
                {
                    result.Success = true;
                    result.PdfFilePath = newPdfPath;
                    result.FileName = outputFileName;
                    Console.WriteLine($"🔄 Fallback to original PDF: {newPdfPath}");
                }
            }

            return result;
        }

        private async Task<int> GetPageCount(string pdfPath)
        {
            try
            {
                using var reader = new PdfReader(pdfPath);
                using var pdfDoc = new PdfDocument(reader);
                return pdfDoc.GetNumberOfPages();
            }
            catch
            {
                return 0;
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

    // Request Models
    public class RemovePdfRequest
    {
        public string SessionId { get; set; }
        public string FileName { get; set; }
    }

    public class MergePdfRequest
    {
        public string SessionId { get; set; }
    }

    public class DownloadPdfRequest
    {
        public string SessionId { get; set; }
        public string FileName { get; set; }
    }

    public class MergedPdfInfo
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public long FileSize { get; set; }
        public int TotalPages { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    public class PdfPreviewRequest
    {
        public string SessionId { get; set; }
        public List<string> SelectedSheets { get; set; }
    }

    public class ErrorViewModel
    {
        public string? RequestId { get; set; }
        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
}


//using ExcelToPdfConverter.Models;
//using ExcelToPdfConverter.Services;
//using Microsoft.AspNetCore.Mvc;
//using System.Diagnostics;
//using iText.Kernel.Pdf;
//using iText.Kernel.Geom;
//using iText.Kernel.Pdf.Canvas;
//using iText.Kernel.Pdf.Xobject;
//using iText.Kernel.Utils;
//using System.Text.Json;

//namespace ExcelToPdfConverter.Controllers
//{
//    public class HomeController : Controller
//    {
//        private readonly LibreOfficeService _libreOfficeService;
//        private readonly ExcelPreviewService _previewService;
//        private readonly ExcelProcessingService _excelProcessingService;
//        private readonly IWebHostEnvironment _environment;
//        private readonly ILogger<HomeController> _logger;

//        public HomeController(LibreOfficeService libreOfficeService,
//                            ExcelPreviewService previewService,
//                            ExcelProcessingService excelProcessingService,
//                            IWebHostEnvironment environment,
//                            ILogger<HomeController> logger)
//        {
//            _libreOfficeService = libreOfficeService;
//            _previewService = previewService;
//            _excelProcessingService = excelProcessingService;
//            _environment = environment;
//            _logger = logger;
//        }

//        public IActionResult Index()
//        {
//            _logger.LogInformation("Home page accessed");
//            ViewBag.ValidationResult = null;
//            ViewBag.Error = null;
//            return View();
//        }

//        [HttpPost]
//        [RequestSizeLimit(100_000_000)]
//        public IActionResult Upload(ExcelUploadModel model)
//        {
//            try
//            {
//                _logger.LogInformation("Upload action started");
//                if (model.ExcelFile == null || model.ExcelFile.Length == 0)
//                {
//                    ViewBag.Error = "Please select an Excel file.";
//                    return View("Index");
//                }

//                var extension = System.IO.Path.GetExtension(model.ExcelFile.FileName)?.ToLower();
//                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
//                {
//                    ViewBag.Error = "Please upload only Excel files (.xlsx, .xls, or .xlsm).";
//                    return View("Index");
//                }

//                var validationResult = _previewService.QuickValidate(model.ExcelFile);
//                ViewBag.ValidationResult = validationResult;

//                var sessionId = Guid.NewGuid().ToString();
//                var previewModel = _previewService.GeneratePreview(model.ExcelFile, sessionId);

//                var filePath = _libreOfficeService.SaveUploadedFile(model.ExcelFile);

//                HttpContext.Session.SetString(sessionId + "_filePath", filePath);
//                HttpContext.Session.SetString(sessionId + "_fileName", model.ExcelFile.FileName ?? "unknown");
//                HttpContext.Session.SetString("CurrentSessionId", sessionId);

//                ViewBag.SessionId = sessionId;

//                return View("Preview", previewModel);
//            }
//            catch (Exception ex)
//            {
//                _logger.LogError(ex, "Error during file upload");
//                ViewBag.Error = $"Error processing file: {ex.Message}";
//                return View("Index");
//            }
//        }

//        [HttpPost]
//        public async Task<IActionResult> GeneratePdfPreview([FromBody] PdfPreviewRequest request)
//        {
//            try
//            {
//                Console.WriteLine($"=== GeneratePdfPreview Called ===");
//                Console.WriteLine($"Session ID: {request.SessionId}");
//                Console.WriteLine($"Selected Sheets: {string.Join(", ", request.SelectedSheets ?? new List<string>())}");

//                var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");
//                var originalFileName = HttpContext.Session.GetString(request.SessionId + "_fileName");

//                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
//                {
//                    return Json(new { success = false, message = "File not found. Please upload again." });
//                }

//                // Step 1: Process Excel file with color preservation
//                string processedFilePath;
//                try
//                {
//                    processedFilePath = await ProcessExcelWithColorPreservation(filePath, request.SelectedSheets);
//                    Console.WriteLine($"✅ Excel file processed with color preservation: {processedFilePath}");
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine($"❌ Excel processing failed: {ex.Message}");
//                    return Json(new { success = false, message = $"Excel processing failed: {ex.Message}" });
//                }

//                // Step 2: Convert to PDF
//                var outputFileName = $"preview_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
//                var result = await ConvertToPdfWithColorPreservation(processedFilePath, outputFileName, request.SelectedSheets);

//                // Cleanup processed file
//                if (processedFilePath != filePath && System.IO.File.Exists(processedFilePath))
//                {
//                    _excelProcessingService.CleanupProcessedFile(processedFilePath);
//                }

//                if (result.Success && System.IO.File.Exists(result.PdfFilePath))
//                {
//                    var pdfBytes = await System.IO.File.ReadAllBytesAsync(result.PdfFilePath);
//                    var pdfBase64 = Convert.ToBase64String(pdfBytes);

//                    // Get page count
//                    int totalPages = 0;
//                    try
//                    {
//                        using (var reader = new PdfReader(result.PdfFilePath))
//                        using (var pdfDoc = new PdfDocument(reader))
//                        {
//                            totalPages = pdfDoc.GetNumberOfPages();
//                        }
//                    }
//                    catch
//                    {
//                        totalPages = 1;
//                    }

//                    // Cleanup PDF file
//                    System.IO.File.Delete(result.PdfFilePath);

//                    return Json(new
//                    {
//                        success = true,
//                        pdfData = pdfBase64,
//                        fileName = outputFileName,
//                        generatedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
//                        totalPages = totalPages
//                    });
//                }
//                else
//                {
//                    return Json(new { success = false, message = result.Message });
//                }
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in GeneratePdfPreview: {ex}");
//                return Json(new { success = false, message = $"Preview generation failed: {ex.Message}" });
//            }
//        }

//        private async Task<string> ProcessExcelWithColorPreservation(string inputFilePath, List<string> selectedSheets)
//        {
//            var outputFilePath = System.IO.Path.Combine(
//                System.IO.Path.GetTempPath(),
//                $"color_preserved_{Guid.NewGuid()}.xlsx");

//            try
//            {
//                using (var sourcePackage = new OfficeOpenXml.ExcelPackage(new FileInfo(inputFilePath)))
//                using (var targetPackage = new OfficeOpenXml.ExcelPackage())
//                {
//                    var sourceWorkbook = sourcePackage.Workbook;
//                    var targetWorkbook = targetPackage.Workbook;

//                    // Add sheets in order
//                    foreach (var sheetName in selectedSheets)
//                    {
//                        var sourceWorksheet = sourceWorkbook.Worksheets[sheetName];
//                        if (sourceWorksheet != null)
//                        {
//                            var targetWorksheet = targetWorkbook.Worksheets.Add(sheetName, sourceWorksheet);
//                            Console.WriteLine($"✅ Copied sheet with full formatting: {sheetName}");
//                        }
//                    }

//                    targetPackage.SaveAs(new FileInfo(outputFilePath));
//                }

//                return outputFilePath;
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in ProcessExcelWithColorPreservation: {ex}");
//                return await CreateSimpleCopy(inputFilePath, selectedSheets);
//            }
//        }

//        private async Task<string> CreateSimpleCopy(string inputFilePath, List<string> selectedSheets)
//        {
//            var outputFilePath = System.IO.Path.Combine(
//                System.IO.Path.GetTempPath(),
//                $"simple_copy_{Guid.NewGuid()}.xlsx");

//            try
//            {
//                using (var sourcePackage = new OfficeOpenXml.ExcelPackage(new FileInfo(inputFilePath)))
//                using (var targetPackage = new OfficeOpenXml.ExcelPackage())
//                {
//                    var sourceWorkbook = sourcePackage.Workbook;
//                    var targetWorkbook = targetPackage.Workbook;

//                    foreach (var sheetName in selectedSheets)
//                    {
//                        var sourceWorksheet = sourceWorkbook.Worksheets[sheetName];
//                        if (sourceWorksheet != null)
//                        {
//                            var targetWorksheet = targetWorkbook.Worksheets.Add(sheetName);

//                            // Copy data only (no formatting)
//                            if (sourceWorksheet.Dimension != null)
//                            {
//                                int maxRows = Math.Min(sourceWorksheet.Dimension.End.Row, 1000);
//                                int maxCols = Math.Min(sourceWorksheet.Dimension.End.Column, 100);

//                                for (int row = 1; row <= maxRows; row++)
//                                {
//                                    for (int col = 1; col <= maxCols; col++)
//                                    {
//                                        targetWorksheet.Cells[row, col].Value = sourceWorksheet.Cells[row, col].Value;
//                                    }
//                                }
//                            }

//                            Console.WriteLine($"✅ Created simple copy of sheet: {sheetName}");
//                        }
//                    }

//                    targetPackage.SaveAs(new FileInfo(outputFilePath));
//                }

//                return outputFilePath;
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in CreateSimpleCopy: {ex}");
//                throw;
//            }
//        }

//        private async Task<ConversionResult> ConvertToPdfWithColorPreservation(
//            string inputFilePath, string outputFileName, List<string> selectedSheets)
//        {
//            var outputDirectory = System.IO.Path.Combine(_environment.WebRootPath, "App_Data", "Previews");
//            Directory.CreateDirectory(outputDirectory);
//            var outputFilePath = System.IO.Path.Combine(outputDirectory, outputFileName);

//            try
//            {
//                var arguments = BuildLibreOfficeArguments(inputFilePath, outputDirectory, selectedSheets);
//                Console.WriteLine($"LibreOffice arguments: {arguments}");

//                var processStartInfo = new ProcessStartInfo
//                {
//                    FileName = GetLibreOfficePath(),
//                    Arguments = arguments,
//                    UseShellExecute = false,
//                    CreateNoWindow = true,
//                    RedirectStandardOutput = true,
//                    RedirectStandardError = true,
//                    WindowStyle = ProcessWindowStyle.Hidden,
//                    WorkingDirectory = outputDirectory
//                };

//                using (var process = new Process())
//                {
//                    process.StartInfo = processStartInfo;
//                    process.Start();

//                    string output = await process.StandardOutput.ReadToEndAsync();
//                    string error = await process.StandardError.ReadToEndAsync();

//                    bool processExited = process.WaitForExit(180000); // 3 minutes

//                    if (processExited && process.ExitCode == 0)
//                    {
//                        var inputFileName = System.IO.Path.GetFileNameWithoutExtension(inputFilePath);
//                        var possibleOutputPaths = new[]
//                        {
//                            System.IO.Path.Combine(outputDirectory, inputFileName + ".pdf"),
//                            System.IO.Path.Combine(outputDirectory, outputFileName)
//                        };

//                        foreach (var path in possibleOutputPaths)
//                        {
//                            if (System.IO.File.Exists(path))
//                            {
//                                return new ConversionResult
//                                {
//                                    Success = true,
//                                    Message = "Conversion successful",
//                                    PdfFilePath = path,
//                                    FileName = outputFileName
//                                };
//                            }
//                        }
//                    }

//                    return new ConversionResult
//                    {
//                        Success = false,
//                        Message = $"Conversion failed. Exit code: {process.ExitCode}, Error: {error}"
//                    };
//                }
//            }
//            catch (Exception ex)
//            {
//                return new ConversionResult
//                {
//                    Success = false,
//                    Message = $"Error during conversion: {ex.Message}"
//                };
//            }
//        }

//        private string BuildLibreOfficeArguments(string inputFilePath, string outputDirectory, List<string> selectedSheets)
//        {
//            var arguments = new List<string>
//            {
//                "--headless",
//                "--norestore",
//                "--nofirststartwizard",
//                "--convert-to pdf:calc_pdf_Export",
//                $"--outdir \"{outputDirectory}\"",
//                $"\"{inputFilePath}\""
//            };

//            return string.Join(" ", arguments);
//        }

//        private string GetLibreOfficePath()
//        {
//            string[] possiblePaths = {
//                @"C:\Program Files\LibreOffice\program\soffice.exe",
//                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
//                @"C:\Program Files\LibreOffice\program\soffice.com",
//                @"C:\Program Files (x86)\LibreOffice\program\soffice.com"
//            };

//            foreach (var path in possiblePaths)
//            {
//                if (System.IO.File.Exists(path))
//                {
//                    Console.WriteLine($"LibreOffice found at: {path}");
//                    return path;
//                }
//            }

//            throw new Exception("LibreOffice not found. Please install LibreOffice from https://www.libreoffice.org/download/download-libreoffice/");
//        }

//        // New Request Model with Rotation
//        public class PdfRequestWithRotation
//        {
//            public string SessionId { get; set; }
//            public List<string> SelectedSheets { get; set; }
//            public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
//            public Dictionary<int, string> OrientationData { get; set; }
//            public Dictionary<int, int> RotationData { get; set; }
//        }

//        public class PageOrderInfoWithRotation
//        {
//            public int OriginalPage { get; set; }
//            public int CurrentOrder { get; set; }
//            public bool Visible { get; set; }
//            public string Orientation { get; set; } = "portrait";
//            public int Rotation { get; set; } = 0;
//        }

//        [HttpPost]
//        public async Task<IActionResult> GenerateReorderedPdf([FromBody] PdfRequestWithRotation request)
//        {
//            try
//            {
//                Console.WriteLine($"=== GenerateReorderedPdf Called ===");
//                Console.WriteLine($"Session ID: {request.SessionId}");
//                Console.WriteLine($"Selected Sheets: {request.SelectedSheets?.Count ?? 0}");
//                Console.WriteLine($"Page Order Data: {request.PageOrderData?.Count ?? 0} pages");
//                Console.WriteLine($"Rotation Data: {request.RotationData?.Count ?? 0} rotations");

//                var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");
//                var originalFileName = HttpContext.Session.GetString(request.SessionId + "_fileName");

//                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
//                {
//                    return Json(new { success = false, message = "File not found. Please upload again." });
//                }

//                // Step 1: Get selected sheets
//                var selectedSheets = request.SelectedSheets ?? new List<string>();

//                // Step 2: Create initial PDF
//                var outputFileName = $"document_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
//                var conversionResult = await ConvertToPdfWithColorPreservation(
//                    filePath, outputFileName, selectedSheets);

//                if (!conversionResult.Success || !System.IO.File.Exists(conversionResult.PdfFilePath))
//                {
//                    return Json(new { success = false, message = conversionResult.Message });
//                }

//                Console.WriteLine($"✅ Initial PDF created: {conversionResult.PdfFilePath}");

//                // Step 3: Apply reordering, orientation and rotation
//                string finalPdfPath = await ApplyPageReorderingOrientationAndRotation(
//                    conversionResult.PdfFilePath,
//                    request.PageOrderData,
//                    request.OrientationData,
//                    request.RotationData);

//                // Step 4: Read final PDF
//                var finalPdfBytes = await System.IO.File.ReadAllBytesAsync(finalPdfPath);
//                var finalPdfBase64 = Convert.ToBase64String(finalPdfBytes);

//                // Step 5: Cleanup
//                if (System.IO.File.Exists(conversionResult.PdfFilePath))
//                    System.IO.File.Delete(conversionResult.PdfFilePath);
//                if (System.IO.File.Exists(finalPdfPath))
//                    System.IO.File.Delete(finalPdfPath);

//                return Json(new
//                {
//                    success = true,
//                    pdfData = finalPdfBase64,
//                    fileName = outputFileName,
//                    message = "PDF generated successfully with rotations"
//                });
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in GenerateReorderedPdf: {ex}");
//                return Json(new { success = false, message = $"Error: {ex.Message}" });
//            }
//        }


//        private async Task<string> ApplyPageReorderingOrientationAndRotation(
//    string pdfPath,
//    List<PageOrderInfoWithRotation> pageOrderData,
//    Dictionary<int, string> orientationData,
//    Dictionary<int, int> rotationData)
//        {
//            var outputPath = System.IO.Path.Combine(
//                System.IO.Path.GetTempPath(),
//                $"final_with_orientation_rotation_{Guid.NewGuid()}.pdf");

//            try
//            {
//                Console.WriteLine("🔄 Applying page reordering, orientation and rotation...");

//                using (var reader = new PdfReader(pdfPath))
//                using (var writer = new PdfWriter(outputPath))
//                using (var newPdf = new PdfDocument(writer))
//                using (var sourcePdf = new PdfDocument(reader))
//                {
//                    // Filter visible pages and sort by current order
//                    var visiblePages = pageOrderData
//                        .Where(p => p.Visible)
//                        .OrderBy(p => p.CurrentOrder)
//                        .ToList();

//                    Console.WriteLine($"📄 Processing {visiblePages.Count} visible pages");

//                    foreach (var pageInfo in visiblePages)
//                    {
//                        int sourcePageNum = pageInfo.OriginalPage;

//                        if (sourcePageNum > 0 && sourcePageNum <= sourcePdf.GetNumberOfPages())
//                        {
//                            // Get the source page
//                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
//                            var sourcePageSize = sourcePage.GetPageSize();

//                            // Get orientation for this page
//                            string orientation = pageInfo.Orientation;
//                            if (orientationData != null && orientationData.ContainsKey(sourcePageNum))
//                            {
//                                orientation = orientationData[sourcePageNum];
//                            }

//                            // Get rotation for this page
//                            int rotation = pageInfo.Rotation;
//                            if (rotationData != null && rotationData.ContainsKey(sourcePageNum))
//                            {
//                                rotation = rotationData[sourcePageNum];
//                            }

//                            // Create page with appropriate orientation
//                            PageSize targetPageSize = orientation == "landscape" ? PageSize.A4.Rotate() : PageSize.A4;

//                            // Create new page
//                            var newPage = newPdf.AddNewPage(targetPageSize);

//                            // Copy content
//                            var copiedPage = sourcePage.CopyAsFormXObject(newPdf);
//                            var canvas = new PdfCanvas(newPage);

//                            // Get page dimensions
//                            float targetWidth = targetPageSize.GetWidth();
//                            float targetHeight = targetPageSize.GetHeight();
//                            float sourceWidth = sourcePageSize.GetWidth();
//                            float sourceHeight = sourcePageSize.GetHeight();

//                            // Calculate scaling to fit page (with margins)
//                            float margin = 20; // 20 points margin
//                            float availableWidth = targetWidth - (2 * margin);
//                            float availableHeight = targetHeight - (2 * margin);

//                            // Calculate scale without rotation
//                            float scaleX = availableWidth / sourceWidth;
//                            float scaleY = availableHeight / sourceHeight;
//                            float scale = Math.Min(scaleX, scaleY);

//                            // Adjust for rotation
//                            if (rotation != 0)
//                            {
//                                // When rotated, the bounding box changes
//                                float rotatedWidth = Math.Abs(sourceWidth * (float)Math.Cos(rotation * Math.PI / 180)) +
//                                                   Math.Abs(sourceHeight * (float)Math.Sin(rotation * Math.PI / 180));
//                                float rotatedHeight = Math.Abs(sourceWidth * (float)Math.Sin(rotation * Math.PI / 180)) +
//                                                    Math.Abs(sourceHeight * (float)Math.Cos(rotation * Math.PI / 180));

//                                float rotatedScaleX = availableWidth / rotatedWidth;
//                                float rotatedScaleY = availableHeight / rotatedHeight;
//                                scale = Math.Min(rotatedScaleX, rotatedScaleY);
//                            }

//                            // Calculate centered position
//                            float scaledWidth = sourceWidth * scale;
//                            float scaledHeight = sourceHeight * scale;
//                            float xOffset = margin + (availableWidth - scaledWidth) / 2;
//                            float yOffset = margin + (availableHeight - scaledHeight) / 2;

//                            // Apply transformations
//                            canvas.SaveState();

//                            // Move to center of scaled content
//                            canvas.ConcatMatrix(1, 0, 0, 1, xOffset + scaledWidth / 2, yOffset + scaledHeight / 2);

//                            // Apply rotation
//                            canvas.ConcatMatrix((float)Math.Cos(rotation * Math.PI / 180),
//                                                (float)Math.Sin(rotation * Math.PI / 180),
//                                                (float)-Math.Sin(rotation * Math.PI / 180),
//                                                (float)Math.Cos(rotation * Math.PI / 180),
//                                                0, 0);

//                            // Move back and apply scaling
//                            canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
//                            canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);

//                            // Draw the content
//                            canvas.AddXObjectAt(copiedPage, 0, 0);
//                            canvas.RestoreState();

//                            canvas.Release();

//                            Console.WriteLine($"✅ Page {sourcePageNum} → Position {pageInfo.CurrentOrder} (Orientation: {orientation}, Rotation: {rotation}°)");
//                        }
//                        else
//                        {
//                            Console.WriteLine($"⚠️ Page {sourcePageNum} not found in source PDF");
//                        }
//                    }

//                    newPdf.Close();
//                    sourcePdf.Close();
//                }

//                Console.WriteLine($"✅ PDF with orientation and rotation created: {outputPath}");
//                return outputPath;
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in ApplyPageReorderingOrientationAndRotation: {ex}");
//                throw;
//            }
//        }


//        [HttpPost]
//        public async Task<IActionResult> ConvertToPdf(string sessionId, List<string> selectedSheets,
//            List<int> sheetOrders, List<string> sheetOrientations)
//        {
//            try
//            {
//                Console.WriteLine($"=== ConvertToPdf Called ===");
//                Console.WriteLine($"Session ID: {sessionId}");
//                Console.WriteLine($"Selected Sheets: {string.Join(", ", selectedSheets ?? new List<string>())}");

//                var filePath = HttpContext.Session.GetString(sessionId + "_filePath");
//                var originalFileName = HttpContext.Session.GetString(sessionId + "_fileName");

//                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
//                {
//                    return Json(new { success = false, message = "File not found. Please upload again." });
//                }

//                var orderedSheets = new List<string>();
//                if (selectedSheets != null && sheetOrders != null && selectedSheets.Count == sheetOrders.Count)
//                {
//                    var sheetOrderMap = selectedSheets.Zip(sheetOrders, (s, o) => new { Sheet = s, Order = o })
//                                                    .OrderBy(x => x.Order)
//                                                    .Select(x => x.Sheet)
//                                                    .ToList();
//                    orderedSheets = sheetOrderMap;
//                    Console.WriteLine($"✅ Ordered sheets (drag & drop order): {string.Join(" → ", orderedSheets)}");
//                }
//                else
//                {
//                    orderedSheets = selectedSheets ?? new List<string>();
//                    Console.WriteLine($"ℹ️ Using default sheet order: {string.Join(" → ", orderedSheets)}");
//                }

//                Console.WriteLine($"🔄 Starting Excel file processing...");
//                string processedFilePath;
//                try
//                {
//                    processedFilePath = await _excelProcessingService.CreateProcessedExcelFileAsync(
//                        filePath, orderedSheets, orderedSheets);
//                    Console.WriteLine($"✅ Excel file processed successfully: {processedFilePath}");
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine($"❌ Excel processing failed, using original file: {ex.Message}");
//                    processedFilePath = filePath;
//                }

//                var orientationMap = new Dictionary<string, string>();
//                if (selectedSheets != null && sheetOrientations != null && selectedSheets.Count == sheetOrientations.Count)
//                {
//                    orientationMap = selectedSheets.Zip(sheetOrientations, (s, o) => new { Sheet = s, Orientation = o })
//                                                 .ToDictionary(x => x.Sheet, x => x.Orientation);
//                    Console.WriteLine($"✅ Orientation mapping: {string.Join(", ", orientationMap)}");
//                }

//                var outputFileName = System.IO.Path.GetFileNameWithoutExtension(originalFileName ?? "converted") + ".pdf";

//                Console.WriteLine($"🔄 Starting Excel to PDF conversion...");
//                var result = await _libreOfficeService.ConvertToPdfAsync(
//                    processedFilePath, outputFileName, orderedSheets, orientationMap);

//                if (processedFilePath != filePath)
//                {
//                    _excelProcessingService.CleanupProcessedFile(processedFilePath);
//                }

//                if (result.Success)
//                {
//                    Console.WriteLine($"✅ Excel to PDF conversion successful: {result.PdfFilePath}");

//                    Console.WriteLine($"🔄 Starting PDF merge process...");
//                    var finalResult = await MergeAllPdfsWithiText7(result.PdfFilePath, outputFileName, orderedSheets, orientationMap);

//                    HttpContext.Session.Remove(sessionId + "_filePath");
//                    HttpContext.Session.Remove(sessionId + "_fileName");

//                    if (finalResult.Success)
//                    {
//                        Console.WriteLine($"✅ Final PDF created: {finalResult.PdfFilePath} ({finalResult.TotalPages} pages)");

//                        var fileBytes = await System.IO.File.ReadAllBytesAsync(finalResult.PdfFilePath);

//                        if (System.IO.File.Exists(result.PdfFilePath))
//                            System.IO.File.Delete(result.PdfFilePath);
//                        if (System.IO.File.Exists(finalResult.PdfFilePath))
//                            System.IO.File.Delete(finalResult.PdfFilePath);

//                        Console.WriteLine($"✅ Returning final PDF: {finalResult.FileName}");
//                        return File(fileBytes, "application/pdf", finalResult.FileName);
//                    }
//                    else
//                    {
//                        Console.WriteLine($"❌ PDF merge failed, returning original PDF");
//                        var fileBytes = await System.IO.File.ReadAllBytesAsync(result.PdfFilePath);
//                        System.IO.File.Delete(result.PdfFilePath);
//                        return File(fileBytes, "application/pdf", result.FileName);
//                    }
//                }
//                else
//                {
//                    Console.WriteLine($"❌ Excel to PDF conversion failed: {result.Message}");
//                    return Json(new { success = false, message = result.Message });
//                }
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Exception in ConvertToPdf: {ex}");
//                return Json(new { success = false, message = $"Conversion failed: {ex.Message}" });
//            }
//        }

//        private async Task<ConversionResult> MergeAllPdfsWithiText7(string newPdfPath, string outputFileName,
//            List<string> orderedSheets, Dictionary<string, string> orientationMap)
//        {
//            var result = new ConversionResult();
//            var mergedPdfPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"final_merged_{Guid.NewGuid()}.pdf");

//            try
//            {
//                Console.WriteLine("🔄 Starting PDF merge process with iText7...");

//                using (var writer = new PdfWriter(mergedPdfPath))
//                using (var mergedPdfDoc = new PdfDocument(writer))
//                {
//                    if (System.IO.File.Exists(newPdfPath))
//                    {
//                        using (var reader = new PdfReader(newPdfPath))
//                        using (var sourceDoc = new PdfDocument(reader))
//                        {
//                            sourceDoc.CopyPagesTo(1, sourceDoc.GetNumberOfPages(), mergedPdfDoc);
//                            Console.WriteLine($"✅ Added converted Excel PDF: {newPdfPath} ({sourceDoc.GetNumberOfPages()} pages)");
//                        }
//                    }

//                    var pdfDirectory = @"D:\CIPL\SinghAndSons\pdf";
//                    if (System.IO.Directory.Exists(pdfDirectory))
//                    {
//                        var existingPdfFiles = System.IO.Directory.GetFiles(pdfDirectory, "*.pdf")
//                            .OrderBy(f => f)
//                            .ToList();

//                        Console.WriteLine($"📁 Found {existingPdfFiles.Count} existing PDF files to merge");

//                        foreach (var existingPdf in existingPdfFiles)
//                        {
//                            try
//                            {
//                                using (var reader = new PdfReader(existingPdf))
//                                using (var sourceDoc = new PdfDocument(reader))
//                                {
//                                    sourceDoc.CopyPagesTo(1, sourceDoc.GetNumberOfPages(), mergedPdfDoc);
//                                    Console.WriteLine($"✅ Added existing PDF: {System.IO.Path.GetFileName(existingPdf)} ({sourceDoc.GetNumberOfPages()} pages)");
//                                }
//                            }
//                            catch (Exception ex)
//                            {
//                                Console.WriteLine($"❌ Error adding existing PDF {existingPdf}: {ex.Message}");
//                                continue;
//                            }
//                        }
//                    }

//                    mergedPdfDoc.Close();
//                }

//                var fileInfo = new System.IO.FileInfo(mergedPdfPath);
//                if (fileInfo.Exists && fileInfo.Length > 0)
//                {
//                    result.Success = true;
//                    result.PdfFilePath = mergedPdfPath;
//                    result.FileName = $"merged_{System.IO.Path.GetFileNameWithoutExtension(outputFileName)}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
//                    result.TotalPages = await GetPageCount(mergedPdfPath);
//                    result.Message = $"Successfully merged PDF with {result.TotalPages} total pages";
//                    Console.WriteLine($"✅ Final merged PDF created: {result.PdfFilePath} ({result.TotalPages} pages)");
//                }
//                else
//                {
//                    throw new Exception("Merged PDF file was not created or is empty");
//                }
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in MergeAllPdfsWithiText7: {ex}");
//                result.Success = false;
//                result.Message = $"PDF merge failed: {ex.Message}";

//                if (System.IO.File.Exists(newPdfPath))
//                {
//                    result.Success = true;
//                    result.PdfFilePath = newPdfPath;
//                    result.FileName = outputFileName;
//                    Console.WriteLine($"🔄 Fallback to original PDF: {newPdfPath}");
//                }
//            }

//            return result;
//        }

//        private async Task<int> GetPageCount(string pdfPath)
//        {
//            try
//            {
//                using var reader = new PdfReader(pdfPath);
//                using var pdfDoc = new PdfDocument(reader);
//                return pdfDoc.GetNumberOfPages();
//            }
//            catch
//            {
//                return 0;
//            }
//        }

//        public IActionResult Privacy()
//        {
//            return View();
//        }

//        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
//        public IActionResult Error()
//        {
//            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
//        }
//    }

//    public class PdfPreviewRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//    }

//    public class ErrorViewModel
//    {
//        public string? RequestId { get; set; }
//        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
//    }
//}
