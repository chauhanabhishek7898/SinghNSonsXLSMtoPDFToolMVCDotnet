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
using iText.Layout;
using Microsoft.AspNetCore.Hosting;

namespace ExcelToPdfConverter.Controllers
{
    public class HomeController : Controller
    {
        private readonly LibreOfficeService _libreOfficeService;
        private readonly ExcelPreviewService _previewService;
        private readonly ExcelProcessingService _excelProcessingService;
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<HomeController> _logger;
        private readonly PdfCompressionService _pdfCompressionService;
        private readonly string _previewsDirectory;


        public HomeController(LibreOfficeService libreOfficeService,
                            ExcelPreviewService previewService,
                            ExcelProcessingService excelProcessingService,
                            IWebHostEnvironment environment,
                            ILogger<HomeController> logger,
                            PdfCompressionService pdfCompressionService)
        {
            _libreOfficeService = libreOfficeService;
            _previewService = previewService;
            _excelProcessingService = excelProcessingService;
            _environment = environment;
            _logger = logger;
            _pdfCompressionService = pdfCompressionService;

            // Previews directory setup
            _previewsDirectory = Path.Combine(_environment.WebRootPath, "previews");
            if (!Directory.Exists(_previewsDirectory))
            {
                Directory.CreateDirectory(_previewsDirectory);
            }
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


                // ✅ Check total size before processing
                long totalSize = pdfFiles.Sum(f => f.Length);
                if (totalSize > 100 * 1024 * 1024) // 100MB total limit
                {
                    return Json(new
                    {
                        success = false,
                        message = $"Total file size ({totalSize / 1024 / 1024}MB) exceeds 100MB limit."
                    });
                }


                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", sessionId);
                Directory.CreateDirectory(uploadDirectory);

                // Get existing uploaded files
                var existingFilesJson = HttpContext.Session.GetString(sessionId + "_uploadedPdfs");
                var existingFiles = new List<UploadedPdfInfo>();
                if (!string.IsNullOrEmpty(existingFilesJson))
                {
                    existingFiles = JsonSerializer.Deserialize<List<UploadedPdfInfo>>(existingFilesJson) ?? new List<UploadedPdfInfo>();
                }

                var uploadedFilesList = new List<UploadedPdfInfo>();

                foreach (var file in pdfFiles)
                {
                    if (file.Length > 50 * 1024 * 1024) // 50MB limit
                    {
                        return Json(new { success = false, message = $"{file.FileName} exceeds 50MB limit." });
                    }

                    var uniqueFileName = Guid.NewGuid() + Path.GetExtension(file.FileName);
                    var filePath = Path.Combine(uploadDirectory, uniqueFileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    var uploadedPdfInfo = new UploadedPdfInfo
                    {
                        Name = file.FileName,
                        Path = filePath,
                        Size = file.Length,
                        UploadTime = DateTime.Now,
                        UniqueName = uniqueFileName
                    };

                    existingFiles.Add(uploadedPdfInfo);
                    uploadedFilesList.Add(uploadedPdfInfo);

                    Console.WriteLine($"✅ PDF uploaded: {file.FileName} -> {filePath}");
                }

                // Store uploaded files info in session
                HttpContext.Session.SetString(sessionId + "_uploadedPdfs", JsonSerializer.Serialize(existingFiles));

                return Json(new
                {
                    success = true,
                    message = $"{pdfFiles.Count} PDF file(s) uploaded successfully.",
                    uploadedFiles = uploadedFilesList.Select(f => new
                    {
                        name = f.Name,
                        size = f.Size,
                        uniqueName = f.UniqueName
                    }).ToList()
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in UploadPdfFiles: {ex}");
                return Json(new { success = false, message = $"Error uploading files: {ex.Message}" });
            }
        }

        // Add this class in HomeController
        public class UploadedPdfInfo
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public long Size { get; set; }
            public DateTime UploadTime { get; set; }
            public string UniqueName { get; set; }
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


        //[HttpPost]
        //public async Task<IActionResult> MergeUploadedPdfs([FromBody] MergePdfRequest request)
        //{
        //    try
        //    {
        //        Console.WriteLine($"=== MergeUploadedPdfs Called ===");
        //        Console.WriteLine($"Session ID: {request.SessionId}");

        //        var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);

        //        if (!Directory.Exists(uploadDirectory))
        //        {
        //            return Json(new { success = false, message = "No uploaded PDF files found." });
        //        }

        //        // ✅ STEP 1: Get ONLY NEW uploaded PDF files (excluding merged files)
        //        var uploadedPdfFiles = Directory.GetFiles(uploadDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
        //            .Where(f => !Path.GetFileName(f).StartsWith("merged_"))
        //            .ToList();

        //        if (uploadedPdfFiles.Count == 0)
        //        {
        //            return Json(new { success = false, message = "No NEW PDF files to merge." });
        //        }

        //        Console.WriteLine($"Found {uploadedPdfFiles.Count} NEW PDF files to merge");

        //        // ✅ STEP 2: Get existing merged file (if any)
        //        var existingMergedFiles = Directory.GetFiles(uploadDirectory, "merged_*.pdf", SearchOption.TopDirectoryOnly)
        //            .OrderByDescending(f => f)
        //            .ToList();

        //        string mergedFilePath;
        //        string mergedFileName;
        //        int totalPages = 0;
        //        List<string> allSourceFiles = new List<string>();

        //        // ✅ STEP 3: Create merged file name
        //        mergedFileName = $"merged_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
        //        mergedFilePath = Path.Combine(uploadDirectory, mergedFileName);

        //        using (var writer = new PdfWriter(mergedFilePath))
        //        using (var mergedPdf = new PdfDocument(writer))
        //        {
        //            // ✅ STEP 4: First add pages from existing merged file (if exists)
        //            if (existingMergedFiles.Count > 0)
        //            {
        //                var latestMergedFile = existingMergedFiles.First();
        //                try
        //                {
        //                    using (var reader = new PdfReader(latestMergedFile))
        //                    using (var existingPdf = new PdfDocument(reader))
        //                    {
        //                        int existingPages = existingPdf.GetNumberOfPages();
        //                        existingPdf.CopyPagesTo(1, existingPages, mergedPdf);
        //                        totalPages += existingPages;

        //                        Console.WriteLine($"📄 Added existing merged file: {Path.GetFileName(latestMergedFile)} ({existingPages} pages)");
        //                        allSourceFiles.Add(latestMergedFile);
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    Console.WriteLine($"❌ Error reading existing merged file: {ex.Message}");
        //                }
        //            }

        //            // ✅ STEP 5: Then add NEW uploaded PDF files
        //            foreach (var pdfFile in uploadedPdfFiles)
        //            {
        //                try
        //                {
        //                    using (var reader = new PdfReader(pdfFile))
        //                    using (var sourcePdf = new PdfDocument(reader))
        //                    {
        //                        int sourcePages = sourcePdf.GetNumberOfPages();
        //                        sourcePdf.CopyPagesTo(1, sourcePages, mergedPdf);
        //                        totalPages += sourcePages;

        //                        Console.WriteLine($"✅ Added NEW file: {Path.GetFileName(pdfFile)} ({sourcePages} pages)");
        //                        allSourceFiles.Add(pdfFile);
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    Console.WriteLine($"❌ Error adding {pdfFile}: {ex.Message}");
        //                    continue;
        //                }
        //            }

        //            mergedPdf.Close();
        //        }

        //        Console.WriteLine($"✅ Created merged file: {mergedFileName} with {totalPages} total pages");

        //        // ✅ STEP 6: Clean up OLD files (keep only the NEW merged file)
        //        foreach (var pdfFile in allSourceFiles)
        //        {
        //            try
        //            {
        //                if (pdfFile != mergedFilePath && System.IO.File.Exists(pdfFile))
        //                {
        //                    System.IO.File.Delete(pdfFile);
        //                    Console.WriteLine($"🗑️ Cleaned up: {Path.GetFileName(pdfFile)}");
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine($"⚠️ Error deleting {pdfFile}: {ex.Message}");
        //            }
        //        }

        //        var fileInfo = new FileInfo(mergedFilePath);

        //        // ✅ STEP 7: Update session with ONLY the new merged file
        //        var mergedPdfsList = new List<MergedPdfInfo>
        //{
        //    new MergedPdfInfo
        //    {
        //        FileName = mergedFileName,
        //        FilePath = mergedFilePath,
        //        FileSize = fileInfo.Length,
        //        TotalPages = totalPages,
        //        CreatedAt = DateTime.Now
        //    }
        //};

        //        HttpContext.Session.SetString(request.SessionId + "_mergedPdfs", JsonSerializer.Serialize(mergedPdfsList));

        //        return Json(new
        //        {
        //            success = true,
        //            message = $"Successfully merged {allSourceFiles.Count} files into {totalPages} pages.",
        //            fileName = mergedFileName,
        //            fileSize = fileInfo.Length,
        //            totalPages = totalPages
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"❌ Error in MergeUploadedPdfs: {ex}");
        //        return Json(new { success = false, message = $"Error merging PDFs: {ex.Message}" });
        //    }
        //}


        [HttpPost]
        public async Task<IActionResult> MergeUploadedPdfs([FromBody] MergePdfRequest request)
        {
            //try
            //{
            //    Console.WriteLine($"=== MergeUploadedPdfs Called (Simple Version) ===");

            //    var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);

            //    if (!Directory.Exists(uploadDirectory))
            //    {
            //        return Json(new { success = false, message = "No uploaded PDF files found." });
            //    }

            //    // Get ONLY NEW uploaded PDF files
            //    var uploadedPdfFiles = Directory.GetFiles(uploadDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
            //        .Where(f => !Path.GetFileName(f).StartsWith("merged_"))
            //        .ToList();

            //    if (uploadedPdfFiles.Count == 0)
            //    {
            //        return Json(new { success = false, message = "No NEW PDF files to merge." });
            //    }

            //    Console.WriteLine($"Found {uploadedPdfFiles.Count} NEW PDF files to merge");

            //    string mergedFilePath;
            //    string mergedFileName;
            //    int totalPages = 0;
            //    List<string> allSourceFiles = new List<string>();

            //    // Create merged file name
            //    mergedFileName = $"merged_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            //    mergedFilePath = Path.Combine(uploadDirectory, mergedFileName);

            //    using (var writer = new PdfWriter(mergedFilePath))
            //    using (var mergedPdf = new PdfDocument(writer))
            //    {
            //        // Get existing merged file (if any)
            //        var existingMergedFiles = Directory.GetFiles(uploadDirectory, "merged_*.pdf", SearchOption.TopDirectoryOnly)
            //            .OrderByDescending(f => f)
            //            .ToList();

            //        if (existingMergedFiles.Count > 0)
            //        {
            //            var latestMergedFile = existingMergedFiles.First();
            //            try
            //            {
            //                using (var reader = new PdfReader(latestMergedFile))
            //                using (var existingPdf = new PdfDocument(reader))
            //                {
            //                    int existingPages = existingPdf.GetNumberOfPages();

            //                    Console.WriteLine($"📄 Adding existing merged file: {Path.GetFileName(latestMergedFile)} ({existingPages} pages)");

            //                    // ✅ SIMPLE: Copy pages without rotation logic
            //                    CopyPagesSimple(existingPdf, mergedPdf, 1, existingPages);
            //                    totalPages += existingPages;
            //                    allSourceFiles.Add(latestMergedFile);
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Console.WriteLine($"❌ Error reading existing merged file: {ex.Message}");
            //            }
            //        }

            //        // Add NEW uploaded PDF files
            //        foreach (var pdfFile in uploadedPdfFiles)
            //        {
            //            try
            //            {
            //                using (var reader = new PdfReader(pdfFile))
            //                using (var sourcePdf = new PdfDocument(reader))
            //                {
            //                    int sourcePages = sourcePdf.GetNumberOfPages();

            //                    Console.WriteLine($"📄 Adding uploaded file: {Path.GetFileName(pdfFile)} ({sourcePages} pages)");

            //                    // ✅ SIMPLE: Copy pages without rotation logic
            //                    CopyPagesSimple(sourcePdf, mergedPdf, 1, sourcePages);
            //                    totalPages += sourcePages;
            //                    allSourceFiles.Add(pdfFile);
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //                Console.WriteLine($"❌ Error adding {pdfFile}: {ex.Message}");
            //                continue;
            //            }
            //        }

            //        mergedPdf.Close();
            //    }

            //    Console.WriteLine($"✅ Created merged file: {mergedFileName} with {totalPages} total pages");


            try
            {
                Console.WriteLine($"=== MergeUploadedPdfs Called (Preserve Original) ===");

                var uploadDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "UploadedPdfs", request.SessionId);

                if (!Directory.Exists(uploadDirectory))
                {
                    return Json(new { success = false, message = "No uploaded PDF files found." });
                }

                // Get ONLY NEW uploaded PDF files
                var uploadedPdfFiles = Directory.GetFiles(uploadDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                    .Where(f => !Path.GetFileName(f).StartsWith("merged_"))
                    .ToList();

                if (uploadedPdfFiles.Count == 0)
                {
                    return Json(new { success = false, message = "No NEW PDF files to merge." });
                }

                Console.WriteLine($"Found {uploadedPdfFiles.Count} NEW PDF files to merge");

                string mergedFilePath;
                string mergedFileName;
                int totalPages = 0;
                List<string> allSourceFiles = new List<string>();

                // Create merged file name
                mergedFileName = $"merged_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                mergedFilePath = Path.Combine(uploadDirectory, mergedFileName);

                using (var writer = new PdfWriter(mergedFilePath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    // Get existing merged file (if any)
                    var existingMergedFiles = Directory.GetFiles(uploadDirectory, "merged_*.pdf", SearchOption.TopDirectoryOnly)
                        .OrderByDescending(f => f)
                        .ToList();

                    if (existingMergedFiles.Count > 0)
                    {
                        var latestMergedFile = existingMergedFiles.First();
                        try
                        {
                            using (var reader = new PdfReader(latestMergedFile))
                            using (var existingPdf = new PdfDocument(reader))
                            {
                                int existingPages = existingPdf.GetNumberOfPages();

                                Console.WriteLine($"📄 Adding existing merged file: {Path.GetFileName(latestMergedFile)} ({existingPages} pages)");

                                // ✅ PRESERVE ORIGINAL: Copy with original orientation
                                CopyPagesWithOriginalOrientation(existingPdf, mergedPdf, 1, existingPages);
                                totalPages += existingPages;
                                allSourceFiles.Add(latestMergedFile);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Error reading existing merged file: {ex.Message}");
                        }
                    }

                    // Add NEW uploaded PDF files
                    foreach (var pdfFile in uploadedPdfFiles)
                    {
                        try
                        {
                            using (var reader = new PdfReader(pdfFile))
                            using (var sourcePdf = new PdfDocument(reader))
                            {
                                int sourcePages = sourcePdf.GetNumberOfPages();

                                Console.WriteLine($"📄 Adding uploaded file: {Path.GetFileName(pdfFile)} ({sourcePages} pages)");

                                // ✅ PRESERVE ORIGINAL: Copy with original orientation
                                CopyPagesWithOriginalOrientation(sourcePdf, mergedPdf, 1, sourcePages);
                                totalPages += sourcePages;
                                allSourceFiles.Add(pdfFile);
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

                Console.WriteLine($"✅ Created merged file: {mergedFileName} with {totalPages} total pages");

                // Clean up OLD files
                foreach (var pdfFile in allSourceFiles)
                {
                    try
                    {
                        if (pdfFile != mergedFilePath && System.IO.File.Exists(pdfFile))
                        {
                            System.IO.File.Delete(pdfFile);
                            Console.WriteLine($"🗑️ Cleaned up: {Path.GetFileName(pdfFile)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ Error deleting {pdfFile}: {ex.Message}");
                    }
                }

                var fileInfo = new FileInfo(mergedFilePath);

                // Update session
                var mergedPdfsList = new List<MergedPdfInfo>
        {
            new MergedPdfInfo
            {
                FileName = mergedFileName,
                FilePath = mergedFilePath,
                FileSize = fileInfo.Length,
                TotalPages = totalPages,
                CreatedAt = DateTime.Now
            }
        };

                HttpContext.Session.SetString(request.SessionId + "_mergedPdfs", JsonSerializer.Serialize(mergedPdfsList));

                return Json(new
                {
                    success = true,
                    message = $"Successfully merged {allSourceFiles.Count} files into {totalPages} pages.",
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

        // ✅ SIMPLE COPY METHOD: No rotation, just fit to page

        private void CopyPagesWithOriginalOrientation(PdfDocument sourcePdf, PdfDocument destPdf, int startPage, int endPage)
        {
            for (int i = startPage; i <= endPage; i++)
            {
                var sourcePage = sourcePdf.GetPage(i);
                var sourcePageSize = sourcePage.GetPageSize();

                // ✅ Determine original orientation from page size
                bool isLandscapeOriginal = sourcePageSize.GetWidth() > sourcePageSize.GetHeight();

                Console.WriteLine($"   Copying page {i}: Size = {sourcePageSize.GetWidth()}x{sourcePageSize.GetHeight()}, " +
                                 $"Original Orientation = {(isLandscapeOriginal ? "Landscape" : "Portrait")}");

                // ✅ Use original orientation
                PageSize targetPageSize = isLandscapeOriginal ? PageSize.A4.Rotate() : PageSize.A4;

                var newPage = destPdf.AddNewPage(targetPageSize);
                var copiedPage = sourcePage.CopyAsFormXObject(destPdf);
                var canvas = new PdfCanvas(newPage);

                // Calculate scaling to fit
                float margin = 20;
                float targetWidth = targetPageSize.GetWidth();
                float targetHeight = targetPageSize.GetHeight();
                float sourceWidth = sourcePageSize.GetWidth();
                float sourceHeight = sourcePageSize.GetHeight();

                float scaleX = (targetWidth - (2 * margin)) / sourceWidth;
                float scaleY = (targetHeight - (2 * margin)) / sourceHeight;
                float scale = Math.Min(scaleX, scaleY);

                float scaledWidth = sourceWidth * scale;
                float scaledHeight = sourceHeight * scale;
                float xOffset = margin + (targetWidth - (2 * margin) - scaledWidth) / 2;
                float yOffset = margin + (targetHeight - (2 * margin) - scaledHeight) / 2;

                // Apply simple scaling - NO ROTATION
                canvas.SaveState();
                canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                canvas.AddXObjectAt(copiedPage, 0, 0);
                canvas.RestoreState();
                canvas.Release();

                Console.WriteLine($"   ✅ Copied page {i} with original orientation");
            }
        }


        private void CopyPagesSimple(PdfDocument sourcePdf, PdfDocument destPdf, int startPage, int endPage)
        {
            for (int i = startPage; i <= endPage; i++)
            {
                var sourcePage = sourcePdf.GetPage(i);
                var sourcePageSize = sourcePage.GetPageSize();

                Console.WriteLine($"   Copying page {i}: Size = {sourcePageSize.GetWidth()}x{sourcePageSize.GetHeight()}");

                // ✅ ALWAYS use Portrait A4 - NO auto-rotation
                var newPage = destPdf.AddNewPage(PageSize.A4);
                var copiedPage = sourcePage.CopyAsFormXObject(destPdf);
                var canvas = new PdfCanvas(newPage);

                // Simple scaling to fit page
                float margin = 20;
                float targetWidth = PageSize.A4.GetWidth();
                float targetHeight = PageSize.A4.GetHeight();
                float sourceWidth = sourcePageSize.GetWidth();
                float sourceHeight = sourcePageSize.GetHeight();

                // Calculate scale to fit
                float scaleX = (targetWidth - (2 * margin)) / sourceWidth;
                float scaleY = (targetHeight - (2 * margin)) / sourceHeight;
                float scale = Math.Min(scaleX, scaleY);

                // Center on page
                float scaledWidth = sourceWidth * scale;
                float scaledHeight = sourceHeight * scale;
                float xOffset = margin + (targetWidth - (2 * margin) - scaledWidth) / 2;
                float yOffset = margin + (targetHeight - (2 * margin) - scaledHeight) / 2;

                // Apply simple transformation - NO ROTATION
                canvas.SaveState();
                canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                canvas.AddXObjectAt(copiedPage, 0, 0);
                canvas.RestoreState();
                canvas.Release();

                Console.WriteLine($"   ✅ Copied page {i} (Scale: {scale:F2}, Offset: {xOffset:F1},{yOffset:F1})");
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

                if (System.IO.Directory.Exists(uploadDirectory))
                {
                    // Delete ALL files in the upload directory
                    var files = System.IO.Directory.GetFiles(uploadDirectory);
                    foreach (var file in files)
                    {
                        try
                        {
                            System.IO.File.Delete(file);
                            Console.WriteLine($"🗑️ Deleted: {Path.GetFileName(file)}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ Error deleting {file}: {ex.Message}");
                        }
                    }

                    // Optionally delete the directory itself
                    try
                    {
                        System.IO.Directory.Delete(uploadDirectory);
                        Console.WriteLine($"🗑️ Deleted directory: {uploadDirectory}");
                    }
                    catch { }
                }

                // Clear session data
                HttpContext.Session.Remove(request.SessionId + "_uploadedPdfs");
                HttpContext.Session.Remove(request.SessionId + "_mergedPdfs");

                return Json(new { success = true, message = "All uploaded PDF files have been removed." });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in RemoveMergedPdf: {ex}");
                return Json(new { success = false, message = $"Error removing PDF: {ex.Message}" });
            }
        }


        //// ✅ NEW: Merge with page order preservation
        private async Task<string> MergeExcelWithUploadedPdfsWithPageOrder(
            string excelPdfPath,
            string uploadedPdfPath,
            List<PageOrderInfoWithRotation> pageOrderData)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"preview_merged_with_order_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Merging Excel PDF with uploaded PDF with page order...");

                using (var writer = new PdfWriter(outputPath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    // Create a combined list of all pages with their order
                    var allPages = new List<MergedPageInfo>();

                    // Add Excel PDF pages
                    if (System.IO.File.Exists(excelPdfPath))
                    {
                        using (var reader = new PdfReader(excelPdfPath))
                        using (var excelPdf = new PdfDocument(reader))
                        {
                            for (int i = 1; i <= excelPdf.GetNumberOfPages(); i++)
                            {
                                allPages.Add(new MergedPageInfo
                                {
                                    Source = "excel",
                                    PageNumber = i,
                                    OriginalPage = i,
                                    PdfDoc = excelPdf,
                                    IsExcel = true
                                });
                            }
                        }
                    }

                    // Add uploaded PDF pages
                    if (System.IO.File.Exists(uploadedPdfPath))
                    {
                        using (var reader = new PdfReader(uploadedPdfPath))
                        using (var uploadedPdf = new PdfDocument(reader))
                        {
                            for (int i = 1; i <= uploadedPdf.GetNumberOfPages(); i++)
                            {
                                allPages.Add(new MergedPageInfo
                                {
                                    Source = "uploaded",
                                    PageNumber = i,
                                    OriginalPage = i + 1000, // Offset to distinguish from Excel pages
                                    PdfDoc = uploadedPdf,
                                    IsExcel = false
                                });
                            }
                        }
                    }

                    Console.WriteLine($"📊 Total pages to merge: {allPages.Count}");

                    // Apply page order from pageOrderData if available
                    if (pageOrderData != null && pageOrderData.Any())
                    {
                        // Sort by current order
                        var orderedPages = new List<MergedPageInfo>();

                        foreach (var pageInfo in pageOrderData.OrderBy(p => p.CurrentOrder))
                        {
                            if (pageInfo.Visible)
                            {
                                var page = allPages.FirstOrDefault(p =>
                                    (p.IsExcel && p.OriginalPage == pageInfo.OriginalPage) ||
                                    (!p.IsExcel && p.OriginalPage == pageInfo.OriginalPage));

                                if (page != null)
                                {
                                    orderedPages.Add(page);
                                }
                            }
                        }

                        // Add remaining pages that aren't in pageOrderData
                        var remainingPages = allPages.Where(p =>
                            !orderedPages.Any(op => op.OriginalPage == p.OriginalPage && op.Source == p.Source));
                        orderedPages.AddRange(remainingPages);

                        allPages = orderedPages;
                    }

                    // Now add pages to merged PDF in correct order
                    foreach (var pageInfo in allPages)
                    {
                        try
                        {
                            // Create new page with appropriate orientation
                            PageSize pageSize = PageSize.A4;
                            if (pageOrderData != null)
                            {
                                var pageOrderInfo = pageOrderData.FirstOrDefault(p =>
                                    p.OriginalPage == pageInfo.OriginalPage);
                                if (pageOrderInfo != null)
                                {
                                    pageSize = pageOrderInfo.Orientation == "landscape" ?
                                        PageSize.A4.Rotate() : PageSize.A4;
                                }
                            }

                            var newPage = mergedPdf.AddNewPage(pageSize);
                            var sourcePage = pageInfo.PdfDoc.GetPage(pageInfo.PageNumber);
                            var copiedPage = sourcePage.CopyAsFormXObject(mergedPdf);
                            var canvas = new PdfCanvas(newPage);

                            // Apply rotation if specified
                            int rotation = 0;
                            if (pageOrderData != null)
                            {
                                var pageOrderInfo = pageOrderData.FirstOrDefault(p =>
                                    p.OriginalPage == pageInfo.OriginalPage);
                                if (pageOrderInfo != null)
                                {
                                    rotation = pageOrderInfo.Rotation;
                                }
                            }

                            // Calculate positioning
                            float margin = 20;
                            float targetWidth = pageSize.GetWidth();
                            float targetHeight = pageSize.GetHeight();
                            float availableWidth = targetWidth - (2 * margin);
                            float availableHeight = targetHeight - (2 * margin);

                            var sourcePageSize = sourcePage.GetPageSize();
                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            // Calculate scaling
                            float scaleX = availableWidth / sourceWidth;
                            float scaleY = availableHeight / sourceHeight;
                            float scale = Math.Min(scaleX, scaleY);

                            // Apply transformations
                            canvas.SaveState();

                            // Move to center
                            canvas.ConcatMatrix(1, 0, 0, 1,
                                margin + (availableWidth - sourceWidth * scale) / 2 + sourceWidth * scale / 2,
                                margin + (availableHeight - sourceHeight * scale) / 2 + sourceHeight * scale / 2);

                            // Apply rotation
                            if (rotation != 0)
                            {
                                canvas.ConcatMatrix(
                                    (float)Math.Cos(rotation * Math.PI / 180),
                                    (float)Math.Sin(rotation * Math.PI / 180),
                                    (float)-Math.Sin(rotation * Math.PI / 180),
                                    (float)Math.Cos(rotation * Math.PI / 180),
                                    0, 0);
                            }

                            // Move back and apply scaling
                            canvas.ConcatMatrix(1, 0, 0, 1, -sourceWidth * scale / 2, -sourceHeight * scale / 2);
                            canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);

                            // Draw content
                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();
                            canvas.Release();

                            Console.WriteLine($"✅ Added {pageInfo.Source} page {pageInfo.PageNumber} (Rotation: {rotation}°)");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"❌ Error adding page: {ex.Message}");
                            continue;
                        }
                    }

                    mergedPdf.Close();
                }

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in MergeExcelWithUploadedPdfsWithPageOrder: {ex}");
                return excelPdfPath; // Return original if merge fails
            }
        }


        [HttpPost]
        public async Task<IActionResult> GeneratePdfPreviewWithFitToPage([FromBody] PdfPreviewWithFitToPageRequest request)
        {
            try
            {
                Console.WriteLine($"=== GeneratePdfPreviewWithFitToPage Called ===");

                // ✅ SET TIMEOUT for large operations
                var timeoutTask = Task.Delay(TimeSpan.FromMinutes(2)); // 2 minutes timeout
                var previewTask = GeneratePreviewInternal(request);

                var completedTask = await Task.WhenAny(previewTask, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    Console.WriteLine($"❌ Preview generation timed out after 2 minutes");
                    return Json(new
                    {
                        success = false,
                        message = "Preview generation is taking too long. Please try with smaller PDF files."
                    });
                }

                return await previewTask;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GeneratePdfPreviewWithFitToPage: {ex}");
                return Json(new { success = false, message = $"Preview generation failed: {ex.Message}" });
            }
        }

        private async Task<IActionResult> GeneratePreviewInternal(PdfPreviewWithFitToPageRequest request)
        {
            var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");

            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
            {
                return Json(new { success = false, message = "File not found." });
            }

            // ✅ Show progress in console
            Console.WriteLine($"🔄 Step 1/4: Converting Excel to PDF...");

            var outputFileName = $"preview_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            var result = await ConvertToPdfWithColorPreservation(
                filePath,
                outputFileName,
                request.SelectedSheets);

            if (!result.Success || !System.IO.File.Exists(result.PdfFilePath))
            {
                return Json(new { success = false, message = result.Message });
            }

            Console.WriteLine($"✅ Step 1/4: Excel PDF created");

            // Convert page order data
            List<PageOrderInfoWithRotation> controllerPageOrderData = null;

            if (request.PageOrderData != null && request.PageOrderData.Any())
            {
                controllerPageOrderData = request.PageOrderData.Select(p => new PageOrderInfoWithRotation
                {
                    OriginalPage = p.OriginalPage,
                    CurrentOrder = p.CurrentOrder,
                    Visible = p.Visible,
                    Orientation = p.Orientation ?? "portrait",
                    Rotation = p.Rotation
                }).ToList();
            }

            // Step 2: Check for merged PDFs
            Console.WriteLine($"🔄 Step 2/4: Checking merged PDFs...");

            string finalPdfPath = result.PdfFilePath;

            if (request.IncludeMergedPdfs)
            {
                var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();

                if (mergedPdfs.Any())
                {
                    var latestMergedPdf = mergedPdfs
                        .OrderByDescending(m => m.CreatedAt)
                        .FirstOrDefault();

                    if (latestMergedPdf != null && System.IO.File.Exists(latestMergedPdf.FilePath))
                    {
                        Console.WriteLine($"🔄 Step 2/4: Merging with uploaded PDF ({latestMergedPdf.TotalPages} pages)...");

                        finalPdfPath = await MergeExcelWithLatestMergedPdf(result.PdfFilePath, latestMergedPdf.FilePath);

                        Console.WriteLine($"✅ Step 2/4: Merge completed");
                    }
                }
            }

            // Step 3: Apply FitToPage
            Console.WriteLine($"🔄 Step 3/4: Applying FitToPage...");

            if (controllerPageOrderData != null && controllerPageOrderData.Any())
            {
                var processedPath = await ApplyOnlyFitToPage(finalPdfPath, controllerPageOrderData);

                // Cleanup old file
                if (finalPdfPath != result.PdfFilePath && System.IO.File.Exists(finalPdfPath))
                {
                    System.IO.File.Delete(finalPdfPath);
                }

                finalPdfPath = processedPath;
            }

            Console.WriteLine($"✅ Step 3/4: FitToPage applied");

            // Step 4: ✅ PDF Compression and URL Generation
            Console.WriteLine($"🔄 Step 4/4: Compressing PDF and generating URL...");

            //var pdfBytes = await System.IO.File.ReadAllBytesAsync(finalPdfPath);
            //var pdfBase64 = Convert.ToBase64String(pdfBytes);

            // ✅ Compress PDF if too large
            finalPdfPath = await _pdfCompressionService.CompressPdfIfLarge(finalPdfPath, maxSizeMB: 10);

            // ✅ Generate unique filename
            var uniqueFileName = $"preview_{Guid.NewGuid()}.pdf";
            var webFilePath = Path.Combine(_previewsDirectory, uniqueFileName);

            // ✅ Copy to web directory
            System.IO.File.Copy(finalPdfPath, webFilePath, true);

            // Get final page count
            int totalPagesFinal = 0;
            try
            {
                using (var reader = new PdfReader(finalPdfPath))
                using (var pdfDoc = new PdfDocument(reader))
                {
                    totalPagesFinal = pdfDoc.GetNumberOfPages();
                }
            }
            catch
            {
                totalPagesFinal = 1;
            }

            // ✅ Generate URL
            var pdfUrl = Url.Content($"~/previews/{uniqueFileName}");

            // ✅ Store preview PDF in session
            //var previewPdfInfo = new PreviewPdfInfo
            //{
            //    Base64Data = pdfBase64,
            //    FileName = outputFileName,
            //    PageCount = totalPagesFinal,
            //    CreatedAt = DateTime.Now,
            //    PageOrderData = controllerPageOrderData,
            //    SelectedSheets = request.SelectedSheets
            //};

            // ✅ Store preview PDF info in session WITHOUT Base64
            var previewPdfInfo = new PreviewPdfInfo
            {
                Base64Data = null, // ✅ Base64 नहीं भेजेंगे
                PdfUrl = pdfUrl, // ✅ URL स्टोर करें
                FileName = uniqueFileName,
                PageCount = totalPagesFinal,
                CreatedAt = DateTime.Now,
                PageOrderData = controllerPageOrderData,
                SelectedSheets = request.SelectedSheets,
                FilePath = webFilePath, // ✅ सर्वर पर फाइल पथ भी स्टोर करें
                // ✅ NEW: Store the initial page order
                CurrentPageOrder = controllerPageOrderData ?? new List<PageOrderInfoWithRotation>()
            };

            HttpContext.Session.SetString(request.SessionId + "_previewPdfInfo",
                JsonSerializer.Serialize(previewPdfInfo));

            //Console.WriteLine($"✅ Step 4/4: Preview ready ({totalPagesFinal} pages)");

            Console.WriteLine($"✅ Step 4/4: Preview ready ({totalPagesFinal} pages) at URL: {pdfUrl}");

            // Cleanup temporary files
            CleanupTempFiles(result.PdfFilePath, finalPdfPath);

            // Cleanup
            //System.IO.File.Delete(result.PdfFilePath);
            if (finalPdfPath != result.PdfFilePath && System.IO.File.Exists(finalPdfPath))
            {
                System.IO.File.Delete(finalPdfPath);
            }

            //return Json(new
            //{
            //    success = true,
            //    pdfData = pdfBase64,
            //    fileName = outputFileName,
            //    totalPages = totalPagesFinal,
            //    message = $"Preview generated successfully with {totalPagesFinal} pages"
            //});

            return Json(new
            {
                success = true,
                pdfUrl = pdfUrl, // ✅ Base64 की जगह URL भेजें
                fileName = outputFileName,
                totalPages = totalPagesFinal,
                message = $"Preview generated successfully with {totalPagesFinal} pages"
            });
        }


        private void CleanupTempFiles(params string[] filePaths)
        {
            foreach (var filePath in filePaths)
            {
                try
                {
                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Error deleting temp file {filePath}: {ex.Message}");
                }
            }
        }


        //[HttpPost]
        //public async Task<IActionResult> DownloadFromPreview([FromBody] DownloadFromPreviewRequest request)
        //{
        //    try
        //    {
        //        Console.WriteLine($"=== DownloadFromPreview Called ===");
        //        Console.WriteLine($"Session ID: {request.SessionId}");
        //        Console.WriteLine($"Apply Actions: {request.ApplyPageOrderData?.Count ?? 0} actions");

        //        // ✅ STEP 1: Get stored preview PDF from session
        //        var previewPdfJson = HttpContext.Session.GetString(request.SessionId + "_previewPdfInfo");

        //        if (string.IsNullOrEmpty(previewPdfJson))
        //        {
        //            return Json(new { success = false, message = "Preview PDF not found. Please generate preview first." });
        //        }

        //        var previewPdfInfo = JsonSerializer.Deserialize<PreviewPdfInfo>(previewPdfJson);

        //        //if (previewPdfInfo == null || string.IsNullOrEmpty(previewPdfInfo.Base64Data))
        //        if (previewPdfInfo == null || string.IsNullOrEmpty(previewPdfInfo.PdfUrl))
        //        {
        //            return Json(new { success = false, message = "Preview PDF data is invalid." });
        //        }

        //        //Console.WriteLine($"📊 Found stored preview PDF: {previewPdfInfo.FileName} with {previewPdfInfo.PageCount} pages");

        //        Console.WriteLine($"📊 Found stored preview PDF at URL: {previewPdfInfo.PdfUrl}");

        //        //// ✅ STEP 2: Convert Base64 to PDF file
        //        //var pdfBytes = Convert.FromBase64String(previewPdfInfo.Base64Data);
        //        //var tempPdfPath = Path.Combine(Path.GetTempPath(), $"preview_{Guid.NewGuid()}.pdf");
        //        //await System.IO.File.WriteAllBytesAsync(tempPdfPath, pdfBytes);

        //        //// ✅ STEP 3: Apply actions directly to the preview PDF
        //        //string finalPdfPath = tempPdfPath;

        //        //if (request.ApplyPageOrderData != null && request.ApplyPageOrderData.Any())
        //        //{
        //        //    Console.WriteLine($"🔄 Applying {request.ApplyPageOrderData.Count} actions to preview PDF...");

        //        //    // Use requested actions OR fallback to stored actions
        //        //    var pageOrderData = request.ApplyPageOrderData ?? previewPdfInfo.PageOrderData;

        //        //    if (pageOrderData != null && pageOrderData.Any())
        //        //    {
        //        //        // Apply only FitToPage (same as preview) - NO LibreOffice conversion
        //        //        finalPdfPath = await ApplyOnlyFitToPage(tempPdfPath, pageOrderData);
        //        //        System.IO.File.Delete(tempPdfPath);

        //        //        Console.WriteLine($"✅ Actions applied to preview PDF: {finalPdfPath}");
        //        //    }
        //        //}

        //        //// ✅ STEP 4: Read final PDF
        //        //var finalPdfBytes = await System.IO.File.ReadAllBytesAsync(finalPdfPath);
        //        //var finalPdfBase64 = Convert.ToBase64String(finalPdfBytes);

        //        //// Get final page count
        //        //int totalPages = 0;
        //        //try
        //        //{
        //        //    using (var reader = new PdfReader(finalPdfPath))
        //        //    using (var pdfDoc = new PdfDocument(reader))
        //        //    {
        //        //        totalPages = pdfDoc.GetNumberOfPages();
        //        //        Console.WriteLine($"📊 Final Download PDF: {totalPages} pages");
        //        //    }
        //        //}
        //        //catch (Exception ex)
        //        //{
        //        //    Console.WriteLine($"⚠️ Error getting page count: {ex.Message}");
        //        //    totalPages = 1;
        //        //}

        //        //// ✅ STEP 5: Cleanup
        //        //if (System.IO.File.Exists(tempPdfPath))
        //        //    System.IO.File.Delete(tempPdfPath);
        //        //if (System.IO.File.Exists(finalPdfPath) && finalPdfPath != tempPdfPath)
        //        //    System.IO.File.Delete(finalPdfPath);

        //        //return Json(new
        //        //{
        //        //    success = true,
        //        //    pdfData = finalPdfBase64,
        //        //    fileName = $"final_{previewPdfInfo.FileName}",
        //        //    totalPages = totalPages,
        //        //    message = "PDF downloaded from preview (no re-conversion)"
        //        //});

        //        // ✅ Check if file exists on server
        //        if (!System.IO.File.Exists(previewPdfInfo.FilePath))
        //        {
        //            return Json(new { success = false, message = "Preview PDF file not found on server." });
        //        }

        //        // ✅ Read the file
        //        var fileBytes = await System.IO.File.ReadAllBytesAsync(previewPdfInfo.FilePath);

        //        // ✅ Return as file download
        //        return File(fileBytes, "application/pdf", $"final_{previewPdfInfo.FileName}");


        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"❌ Error in DownloadFromPreview: {ex}");
        //        return Json(new { success = false, message = $"Download failed: {ex.Message}" });
        //    }
        //}

        [HttpPost]
        public async Task<IActionResult> DownloadFromPreview([FromBody] DownloadFromPreviewRequest request)
        {
            try
            {
                Console.WriteLine($"=== DownloadFromPreview Called ===");
                Console.WriteLine($"Session ID: {request.SessionId}");
                Console.WriteLine($"Apply Actions: {request.ApplyPageOrderData?.Count ?? 0} actions");

                // ✅ STEP 1: Get stored preview PDF from session
                var previewPdfJson = HttpContext.Session.GetString(request.SessionId + "_previewPdfInfo");

                if (string.IsNullOrEmpty(previewPdfJson))
                {
                    return Json(new { success = false, message = "Preview PDF not found. Please generate preview first." });
                }

                var previewPdfInfo = JsonSerializer.Deserialize<PreviewPdfInfo>(previewPdfJson);

                if (previewPdfInfo == null || string.IsNullOrEmpty(previewPdfInfo.FilePath))
                {
                    return Json(new { success = false, message = "Preview PDF data is invalid." });
                }

                Console.WriteLine($"📊 Found stored preview PDF: {previewPdfInfo.FileName} with {previewPdfInfo.PageCount} pages");

                // ✅ STEP 2: Check if file exists on server
                if (!System.IO.File.Exists(previewPdfInfo.FilePath))
                {
                    return Json(new { success = false, message = "Preview PDF file not found on server." });
                }

                // ✅ STEP 3: Apply modifications if provided
                string finalPdfPath = previewPdfInfo.FilePath;

                if (request.ApplyPageOrderData != null && request.ApplyPageOrderData.Any())
                {
                    Console.WriteLine($"🔄 Applying {request.ApplyPageOrderData.Count} actions to preview PDF...");

                    // Create temporary file path for processed PDF
                    var tempPdfPath = Path.Combine(Path.GetTempPath(), $"modified_{Guid.NewGuid()}.pdf");

                    try
                    {
                        // Apply modifications to the preview PDF
                        await ApplyModificationsToPdf(previewPdfInfo.FilePath, tempPdfPath, request.ApplyPageOrderData);
                        finalPdfPath = tempPdfPath;

                        Console.WriteLine($"✅ Modifications applied successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error applying modifications: {ex.Message}");
                        // Fallback to original if modification fails
                        finalPdfPath = previewPdfInfo.FilePath;
                    }
                }

                // ✅ STEP 4: Read the final PDF
                var fileBytes = await System.IO.File.ReadAllBytesAsync(finalPdfPath);

                // ✅ STEP 5: Cleanup temp file if created
                if (finalPdfPath != previewPdfInfo.FilePath && System.IO.File.Exists(finalPdfPath))
                {
                    try
                    {
                        System.IO.File.Delete(finalPdfPath);
                    }
                    catch { }
                }

                // ✅ STEP 6: Return as file download
                return File(fileBytes, "application/pdf", $"final_{previewPdfInfo.FileName}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in DownloadFromPreview: {ex}");
                return Json(new { success = false, message = $"Download failed: {ex.Message}" });
            }
        }

        // ✅ NEW: Helper method to apply modifications
        private async Task ApplyModificationsToPdf(string sourcePdfPath, string destPdfPath,
            List<PageOrderInfoWithRotation> pageOrderData)
        {
            try
            {
                Console.WriteLine($"🔄 Applying page modifications...");

                using (var reader = new PdfReader(sourcePdfPath))
                using (var writer = new PdfWriter(destPdfPath))
                using (var destPdf = new PdfDocument(writer))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    // Sort pages by current order
                    var orderedPages = pageOrderData
                        .Where(p => p.Visible)
                        .OrderBy(p => p.CurrentOrder)
                        .ToList();

                    Console.WriteLine($"📄 Processing {orderedPages.Count} pages in new order");

                    foreach (var pageInfo in orderedPages)
                    {
                        int sourcePageNum = pageInfo.OriginalPage;

                        if (sourcePageNum > 0 && sourcePageNum <= sourcePdf.GetNumberOfPages())
                        {
                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
                            var sourcePageSize = sourcePage.GetPageSize();

                            // Get orientation and rotation
                            string orientation = pageInfo.Orientation ?? "portrait";
                            PageSize targetPageSize = orientation == "landscape" ? PageSize.A4.Rotate() : PageSize.A4;
                            int rotation = pageInfo.Rotation;

                            // Create new page
                            var newPage = destPdf.AddNewPage(targetPageSize);
                            var copiedPage = sourcePage.CopyAsFormXObject(destPdf);
                            var canvas = new PdfCanvas(newPage);

                            // Calculate scaling and positioning (same as in preview)
                            float margin = 20;
                            float targetWidth = targetPageSize.GetWidth();
                            float targetHeight = targetPageSize.GetHeight();
                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            float scale = 0.90f; // Same scale as preview
                            float scaledWidth = sourceWidth * scale;
                            float scaledHeight = sourceHeight * scale;
                            float xOffset = margin + (targetWidth - (2 * margin) - scaledWidth) / 2;
                            float yOffset = margin + (targetHeight - (2 * margin) - scaledHeight) / 2;

                            // Apply transformations
                            canvas.SaveState();

                            if (rotation != 0)
                            {
                                canvas.ConcatMatrix(1, 0, 0, 1,
                                    xOffset + scaledWidth / 2,
                                    yOffset + scaledHeight / 2);

                                canvas.ConcatMatrix((float)Math.Cos(rotation * Math.PI / 180),
                                                    (float)Math.Sin(rotation * Math.PI / 180),
                                                    (float)-Math.Sin(rotation * Math.PI / 180),
                                                    (float)Math.Cos(rotation * Math.PI / 180),
                                                    0, 0);

                                canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
                                canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);
                            }
                            else
                            {
                                canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                            }

                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();

                            Console.WriteLine($"✅ Page {sourcePageNum}: Reordered to position {pageInfo.CurrentOrder} " +
                                             $"(Orientation: {orientation}, Rotation: {rotation}°)");
                        }
                    }

                    destPdf.Close();
                    sourcePdf.Close();
                }

                Console.WriteLine($"✅ Modifications applied successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplyModificationsToPdf: {ex}");
                throw;
            }
        }

        //    private async Task ApplyModificationsToPdf(string sourcePdfPath, string destPdfPath,
        //List<PageOrderInfoWithRotation> pageOrderData)
        //    {
        //        try
        //        {
        //            Console.WriteLine($"🔄 Applying page modifications to: {sourcePdfPath}");

        //            using (var reader = new PdfReader(sourcePdfPath))
        //            {
        //                // ✅ IMPORTANT: Set this to prevent auto-rotation
        //                reader.SetMemorySavingMode(false);
        //                reader.SetUnethicalReading(true);

        //                using (var writer = new PdfWriter(destPdfPath))
        //                using (var destPdf = new PdfDocument(writer))
        //                using (var sourcePdf = new PdfDocument(reader))
        //                {
        //                    int sourceTotalPages = sourcePdf.GetNumberOfPages();
        //                    Console.WriteLine($"📄 Source PDF has {sourceTotalPages} pages");

        //                    // Sort pages by current order
        //                    var orderedPages = pageOrderData
        //                        .Where(p => p.Visible)
        //                        .OrderBy(p => p.CurrentOrder)
        //                        .ToList();

        //                    Console.WriteLine($"📄 Processing {orderedPages.Count} pages in new order");

        //                    foreach (var pageInfo in orderedPages)
        //                    {
        //                        int sourcePageNum = pageInfo.OriginalPage;

        //                        if (sourcePageNum > 0 && sourcePageNum <= sourceTotalPages)
        //                        {
        //                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
        //                            var sourcePageSize = sourcePage.GetPageSize();
        //                            var originalRotation = sourcePage.GetRotation();

        //                            Console.WriteLine($"   Page {sourcePageNum}: Original rotation = {originalRotation}°");

        //                            // ✅ FIX: Use original rotation unless user explicitly changed it
        //                            string orientation = pageInfo.Orientation ?? "portrait";
        //                            int userRotation = pageInfo.Rotation;

        //                            // ✅ If user hasn't applied rotation, use original rotation
        //                            int finalRotation = userRotation != 0 ? userRotation : originalRotation;

        //                            // Determine page size based on orientation
        //                            PageSize targetPageSize;
        //                            if (orientation == "landscape")
        //                            {
        //                                targetPageSize = PageSize.A4.Rotate();
        //                            }
        //                            else if (finalRotation == 90 || finalRotation == 270)
        //                            {
        //                                // If content is rotated 90/270 degrees, use landscape orientation
        //                                targetPageSize = PageSize.A4.Rotate();
        //                            }
        //                            else
        //                            {
        //                                targetPageSize = PageSize.A4;
        //                            }

        //                            Console.WriteLine($"     User rotation: {userRotation}°, Final rotation: {finalRotation}°");

        //                            // Create new page
        //                            var newPage = destPdf.AddNewPage(targetPageSize);
        //                            var copiedPage = sourcePage.CopyAsFormXObject(destPdf);
        //                            var canvas = new PdfCanvas(newPage);

        //                            // Calculate scaling
        //                            float margin = 20;
        //                            float targetWidth = targetPageSize.GetWidth();
        //                            float targetHeight = targetPageSize.GetHeight();
        //                            float sourceWidth = sourcePageSize.GetWidth();
        //                            float sourceHeight = sourcePageSize.GetHeight();

        //                            float scale = 0.90f;
        //                            float scaledWidth = sourceWidth * scale;
        //                            float scaledHeight = sourceHeight * scale;

        //                            // Adjust for rotation
        //                            if (finalRotation == 90 || finalRotation == 270)
        //                            {
        //                                // Swap width/height for rotated pages
        //                                var temp = scaledWidth;
        //                                scaledWidth = scaledHeight;
        //                                scaledHeight = temp;
        //                            }

        //                            float xOffset = margin + (targetWidth - (2 * margin) - scaledWidth) / 2;
        //                            float yOffset = margin + (targetHeight - (2 * margin) - scaledHeight) / 2;

        //                            // Apply transformations
        //                            canvas.SaveState();

        //                            if (finalRotation != 0)
        //                            {
        //                                canvas.ConcatMatrix(1, 0, 0, 1,
        //                                    xOffset + scaledWidth / 2,
        //                                    yOffset + scaledHeight / 2);

        //                                canvas.ConcatMatrix((float)Math.Cos(finalRotation * Math.PI / 180),
        //                                                    (float)Math.Sin(finalRotation * Math.PI / 180),
        //                                                    (float)-Math.Sin(finalRotation * Math.PI / 180),
        //                                                    (float)Math.Cos(finalRotation * Math.PI / 180),
        //                                                    0, 0);

        //                                canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
        //                                canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);
        //                            }
        //                            else
        //                            {
        //                                canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
        //                            }

        //                            canvas.AddXObjectAt(copiedPage, 0, 0);
        //                            canvas.RestoreState();

        //                            Console.WriteLine($"✅ Page {sourcePageNum}: Position {pageInfo.CurrentOrder}, Rotation: {finalRotation}°");
        //                        }
        //                    }

        //                    destPdf.Close();
        //                    sourcePdf.Close();
        //                }
        //            }

        //            Console.WriteLine($"✅ Modifications applied successfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"❌ Error in ApplyModificationsToPdf: {ex}");
        //            throw;
        //        }
        //    }

        // New Request Model
        public class DownloadFromPreviewRequest
        {
            public string SessionId { get; set; }
            public List<PageOrderInfoWithRotation> ApplyPageOrderData { get; set; }
        }


        //private async Task<string> MergeExcelWithLatestMergedPdf(string excelPdfPath, string mergedPdfPath)
        //{
        //    var outputPath = Path.Combine(Path.GetTempPath(), $"excel_with_latest_merged_{Guid.NewGuid()}.pdf");

        //    try
        //    {
        //        Console.WriteLine("🔄 Merging Excel PDF with LATEST merged PDF only...");

        //        // Check file sizes
        //        var excelFileInfo = new FileInfo(excelPdfPath);
        //        var mergedFileInfo = new FileInfo(mergedPdfPath);

        //        Console.WriteLine($"📁 File sizes - Excel: {excelFileInfo.Length / 1024}KB, Merged: {mergedFileInfo.Length / 1024}KB");

        //        // If merged PDF is too large, use optimized merging
        //        if (mergedFileInfo.Length > 50 * 1024 * 1024) // > 50MB
        //        {
        //            Console.WriteLine($"⚠️ Large merged PDF detected ({mergedFileInfo.Length / 1024 / 1024}MB), using optimized merge...");
        //            return await OptimizedLargePdfMerge(excelPdfPath, mergedPdfPath);
        //        }

        //        using (var writer = new PdfWriter(outputPath))
        //        using (var mergedPdf = new PdfDocument(writer))
        //        {
        //            // 1. First add Excel PDF
        //            if (System.IO.File.Exists(excelPdfPath))
        //            {
        //                using (var reader = new PdfReader(excelPdfPath))
        //                using (var excelPdf = new PdfDocument(reader))
        //                {
        //                    var pageCount = excelPdf.GetNumberOfPages();
        //                    excelPdf.CopyPagesTo(1, pageCount, mergedPdf);
        //                    Console.WriteLine($"✅ Added Excel PDF: {pageCount} pages");
        //                }
        //            }

        //            // 2. Then add ONLY the latest merged PDF with progress tracking
        //            if (System.IO.File.Exists(mergedPdfPath))
        //            {
        //                using (var reader = new PdfReader(mergedPdfPath))
        //                using (var mergedSourcePdf = new PdfDocument(reader))
        //                {
        //                    var totalPages = mergedSourcePdf.GetNumberOfPages();
        //                    var batchSize = 50; // Process in batches

        //                    Console.WriteLine($"📄 Merging uploaded PDF: {totalPages} pages...");

        //                    for (int startPage = 1; startPage <= totalPages; startPage += batchSize)
        //                    {
        //                        int endPage = Math.Min(startPage + batchSize - 1, totalPages);
        //                        mergedSourcePdf.CopyPagesTo(startPage, endPage, mergedPdf);

        //                        // Show progress
        //                        if (startPage % 100 == 1 || startPage == 1)
        //                        {
        //                            Console.WriteLine($"   Processed {Math.Min(endPage, totalPages)}/{totalPages} pages");
        //                        }

        //                        // Small delay to prevent UI freeze
        //                        if (totalPages > 100)
        //                        {
        //                            await Task.Delay(10);
        //                        }
        //                    }

        //                    Console.WriteLine($"✅ Added latest merged PDF: {totalPages} pages");
        //                }
        //            }

        //            mergedPdf.Close();
        //        }

        //        Console.WriteLine($"✅ Merge with latest completed: {outputPath}");
        //        return outputPath;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"❌ Error in MergeExcelWithLatestMergedPdf: {ex}");
        //        return excelPdfPath; // Return original if merge fails
        //    }
        //}

        //private async Task<string> MergeExcelWithLatestMergedPdf(string excelPdfPath, string mergedPdfPath)
        //{
        //    var outputPath = Path.Combine(Path.GetTempPath(), $"excel_with_latest_merged_{Guid.NewGuid()}.pdf");

        //    try
        //    {
        //        Console.WriteLine("🔄 Merging Excel PDF with LATEST merged PDF (Simple)...");

        //        using (var writer = new PdfWriter(outputPath))
        //        using (var mergedPdf = new PdfDocument(writer))
        //        {
        //            // 1. First add Excel PDF
        //            if (System.IO.File.Exists(excelPdfPath))
        //            {
        //                using (var reader = new PdfReader(excelPdfPath))
        //                using (var excelPdf = new PdfDocument(reader))
        //                {
        //                    var pageCount = excelPdf.GetNumberOfPages();

        //                    Console.WriteLine($"📄 Adding Excel PDF: {pageCount} pages");

        //                    // ✅ SIMPLE: Copy without rotation logic
        //                    CopyPagesSimple(excelPdf, mergedPdf, 1, pageCount);
        //                    Console.WriteLine($"✅ Added Excel PDF");
        //                }
        //            }

        //            // 2. Then add ONLY the latest merged PDF
        //            if (System.IO.File.Exists(mergedPdfPath))
        //            {
        //                using (var reader = new PdfReader(mergedPdfPath))
        //                using (var mergedSourcePdf = new PdfDocument(reader))
        //                {
        //                    var totalPages = mergedSourcePdf.GetNumberOfPages();

        //                    Console.WriteLine($"📄 Adding merged PDF: {totalPages} pages");

        //                    // ✅ SIMPLE: Copy without rotation logic
        //                    CopyPagesSimple(mergedSourcePdf, mergedPdf, 1, totalPages);
        //                    Console.WriteLine($"✅ Added merged PDF");
        //                }
        //            }

        //            mergedPdf.Close();
        //        }

        //        Console.WriteLine($"✅ Simple merge completed: {outputPath}");
        //        return outputPath;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"❌ Error in MergeExcelWithLatestMergedPdf: {ex}");
        //        return excelPdfPath;
        //    }
        //}


        private async Task<string> MergeExcelWithLatestMergedPdf(string excelPdfPath, string mergedPdfPath)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"excel_with_latest_merged_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Merging Excel PDF with LATEST merged PDF (Preserve Original)...");

                using (var writer = new PdfWriter(outputPath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    // 1. First add Excel PDF
                    if (System.IO.File.Exists(excelPdfPath))
                    {
                        using (var reader = new PdfReader(excelPdfPath))
                        using (var excelPdf = new PdfDocument(reader))
                        {
                            var pageCount = excelPdf.GetNumberOfPages();

                            Console.WriteLine($"📄 Adding Excel PDF: {pageCount} pages");

                            // ✅ PRESERVE ORIGINAL: Copy with original orientation
                            CopyPagesWithOriginalOrientation(excelPdf, mergedPdf, 1, pageCount);
                            Console.WriteLine($"✅ Added Excel PDF with original orientation");
                        }
                    }

                    // 2. Then add ONLY the latest merged PDF
                    if (System.IO.File.Exists(mergedPdfPath))
                    {
                        using (var reader = new PdfReader(mergedPdfPath))
                        using (var mergedSourcePdf = new PdfDocument(reader))
                        {
                            var totalPages = mergedSourcePdf.GetNumberOfPages();

                            Console.WriteLine($"📄 Adding merged PDF: {totalPages} pages");

                            // ✅ PRESERVE ORIGINAL: Copy with original orientation
                            CopyPagesWithOriginalOrientation(mergedSourcePdf, mergedPdf, 1, totalPages);
                            Console.WriteLine($"✅ Added merged PDF with original orientation");
                        }
                    }

                    mergedPdf.Close();
                }

                Console.WriteLine($"✅ Merge completed (Original orientation preserved): {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in MergeExcelWithLatestMergedPdf: {ex}");
                return excelPdfPath;
            }
        }



        private async Task<string> OptimizedLargePdfMerge(string excelPdfPath, string mergedPdfPath)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"optimized_merge_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("⚡ Using optimized merge for large PDF...");

                // Create a temp directory for split files
                var tempDir = Path.Combine(Path.GetTempPath(), $"split_{Guid.NewGuid()}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    // Step 1: Split large PDF into smaller chunks
                    var splitFiles = await SplitLargePdf(mergedPdfPath, tempDir, 50); // 50 pages per chunk

                    Console.WriteLine($"📦 Split into {splitFiles.Count} chunks");

                    // Step 2: Merge Excel PDF first
                    using (var writer = new PdfWriter(outputPath))
                    using (var mergedPdf = new PdfDocument(writer))
                    {
                        // Add Excel PDF
                        if (System.IO.File.Exists(excelPdfPath))
                        {
                            using (var reader = new PdfReader(excelPdfPath))
                            using (var excelPdf = new PdfDocument(reader))
                            {
                                excelPdf.CopyPagesTo(1, excelPdf.GetNumberOfPages(), mergedPdf);
                                Console.WriteLine($"✅ Added Excel PDF: {excelPdf.GetNumberOfPages()} pages");
                            }
                        }

                        // Step 3: Add split files one by one
                        foreach (var splitFile in splitFiles)
                        {
                            try
                            {
                                using (var reader = new PdfReader(splitFile))
                                using (var splitPdf = new PdfDocument(reader))
                                {
                                    splitPdf.CopyPagesTo(1, splitPdf.GetNumberOfPages(), mergedPdf);
                                    Console.WriteLine($"✅ Added chunk: {Path.GetFileName(splitFile)}");
                                }

                                // Clean up split file
                                System.IO.File.Delete(splitFile);

                                // Yield control to prevent UI freeze
                                await Task.Delay(50);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"⚠️ Error processing chunk {splitFile}: {ex.Message}");
                                continue;
                            }
                        }

                        mergedPdf.Close();
                    }

                    return outputPath;
                }
                finally
                {
                    // Cleanup temp directory
                    try
                    {
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in OptimizedLargePdfMerge: {ex}");
                return excelPdfPath;
            }


        }

        private async Task<List<string>> SplitLargePdf(string pdfPath, string outputDir, int pagesPerChunk)
        {
            var splitFiles = new List<string>();

            try
            {
                using (var reader = new PdfReader(pdfPath))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    int totalPages = sourcePdf.GetNumberOfPages();
                    int chunkCount = (int)Math.Ceiling((double)totalPages / pagesPerChunk);

                    Console.WriteLine($"🔪 Splitting {totalPages} pages into {chunkCount} chunks...");

                    for (int chunkIndex = 0; chunkIndex < chunkCount; chunkIndex++)
                    {
                        int startPage = (chunkIndex * pagesPerChunk) + 1;
                        int endPage = Math.Min((chunkIndex + 1) * pagesPerChunk, totalPages);

                        var chunkPath = Path.Combine(outputDir, $"chunk_{chunkIndex + 1}.pdf");

                        using (var writer = new PdfWriter(chunkPath))
                        using (var chunkPdf = new PdfDocument(writer))
                        {
                            sourcePdf.CopyPagesTo(startPage, endPage, chunkPdf);
                            chunkPdf.Close();
                        }

                        splitFiles.Add(chunkPath);
                        Console.WriteLine($"   Created chunk {chunkIndex + 1}: pages {startPage}-{endPage}");

                        // Yield control
                        if (chunkIndex % 5 == 0)
                        {
                            await Task.Delay(10);
                        }
                    }
                }

                return splitFiles;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error splitting PDF: {ex}");
                throw;
            }
        }




        // ✅ New method to merge PDFs in correct order
        private async Task<string> MergeAllPdfsInOrder(string excelPdfPath, string mergedPdfPath)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"ordered_merge_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Merging PDFs in correct order...");

                using (var writer = new PdfWriter(outputPath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    // First add Excel PDF
                    if (System.IO.File.Exists(excelPdfPath))
                    {
                        using (var reader = new PdfReader(excelPdfPath))
                        using (var excelPdf = new PdfDocument(reader))
                        {
                            excelPdf.CopyPagesTo(1, excelPdf.GetNumberOfPages(), mergedPdf);
                            Console.WriteLine($"✅ Added Excel PDF: {excelPdf.GetNumberOfPages()} pages");
                        }
                    }

                    // Then add the merged PDF (which contains all uploaded PDFs)
                    if (System.IO.File.Exists(mergedPdfPath))
                    {
                        using (var reader = new PdfReader(mergedPdfPath))
                        using (var uploadedPdf = new PdfDocument(reader))
                        {
                            uploadedPdf.CopyPagesTo(1, uploadedPdf.GetNumberOfPages(), mergedPdf);
                            Console.WriteLine($"✅ Added merged uploaded PDF: {uploadedPdf.GetNumberOfPages()} pages");
                        }
                    }

                    mergedPdf.Close();
                }

                Console.WriteLine($"✅ Ordered merge completed: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in MergeAllPdfsInOrder: {ex}");
                return excelPdfPath; // Return original if merge fails
            }
        }


        // Add this simple merge method
        private async Task<string> SimpleMergePdfs(string excelPdfPath, string uploadedPdfPath)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"simple_merged_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Simple PDF merge...");

                using (var writer = new PdfWriter(outputPath))
                using (var mergedPdf = new PdfDocument(writer))
                {
                    // Add Excel PDF first
                    if (System.IO.File.Exists(excelPdfPath))
                    {
                        using (var reader = new PdfReader(excelPdfPath))
                        using (var excelPdf = new PdfDocument(reader))
                        {
                            excelPdf.CopyPagesTo(1, excelPdf.GetNumberOfPages(), mergedPdf);
                            Console.WriteLine($"✅ Added Excel PDF: {excelPdf.GetNumberOfPages()} pages");
                        }
                    }

                    // Add uploaded PDF
                    if (System.IO.File.Exists(uploadedPdfPath))
                    {
                        using (var reader = new PdfReader(uploadedPdfPath))
                        using (var uploadedPdf = new PdfDocument(reader))
                        {
                            uploadedPdf.CopyPagesTo(1, uploadedPdf.GetNumberOfPages(), mergedPdf);
                            Console.WriteLine($"✅ Added uploaded PDF: {uploadedPdf.GetNumberOfPages()} pages");
                        }
                    }

                    mergedPdf.Close();
                }

                Console.WriteLine($"✅ Simple merge completed: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in SimpleMergePdfs: {ex}");
                return excelPdfPath; // Return original if merge fails
            }
        }



        // Helper class for merged page info
        private class MergedPageInfo
        {
            public string Source { get; set; } // "excel" or "uploaded"
            public int PageNumber { get; set; }
            public int OriginalPage { get; set; }
            public PdfDocument PdfDoc { get; set; }
            public bool IsExcel { get; set; }
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


        // ✅ NEW: Get consistent page mapping between preview and download
        private List<PageOrderInfoWithRotation> AdjustPageOrderForDownload(
            List<PageOrderInfoWithRotation> previewPageOrder,
            int previewTotalPages,
            int downloadTotalPages)
        {
            try
            {
                Console.WriteLine($"🔄 Adjusting page order: Preview={previewTotalPages}, Download={downloadTotalPages}");

                if (previewTotalPages == downloadTotalPages)
                {
                    Console.WriteLine($"✅ Page counts match, no adjustment needed");
                    return previewPageOrder;
                }

                var adjustedOrder = new List<PageOrderInfoWithRotation>();

                // Simple proportional mapping
                foreach (var pageInfo in previewPageOrder.Where(p => p.Visible))
                {
                    // Calculate proportional page number
                    double proportion = (double)pageInfo.OriginalPage / previewTotalPages;
                    int adjustedPage = (int)Math.Ceiling(proportion * downloadTotalPages);

                    // Ensure within bounds
                    adjustedPage = Math.Max(1, Math.Min(adjustedPage, downloadTotalPages));

                    adjustedOrder.Add(new PageOrderInfoWithRotation
                    {
                        OriginalPage = adjustedPage,
                        CurrentOrder = pageInfo.CurrentOrder,
                        Visible = pageInfo.Visible,
                        Orientation = pageInfo.Orientation,
                        Rotation = pageInfo.Rotation
                    });

                    Console.WriteLine($"   Preview Page {pageInfo.OriginalPage} → Download Page {adjustedPage}");
                }

                return adjustedOrder;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error adjusting page order: {ex.Message}");
                return previewPageOrder; // Return original if adjustment fails
            }
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

                // ✅ CRITICAL: Wait for any pending operations on server side
                //await Task.Delay(2000); // Small delay to ensure client-side is ready

                var filePath = HttpContext.Session.GetString(request.SessionId + "_filePath");
                var originalFileName = HttpContext.Session.GetString(request.SessionId + "_fileName");

                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    return Json(new { success = false, message = "File not found. Please upload again." });
                }

                // ✅ STEP 1: Get the PREVIEW page count from session
                int previewTotalPages = 0;
                var previewPageCountJson = HttpContext.Session.GetString(request.SessionId + "_previewPageCount");
                if (!string.IsNullOrEmpty(previewPageCountJson))
                {
                    previewTotalPages = int.Parse(previewPageCountJson);
                    Console.WriteLine($"📊 Preview PDF had: {previewTotalPages} pages");
                }

                // Step 2: Get selected sheets
                var selectedSheets = request.SelectedSheets ?? new List<string>();

                // Step 3: Create initial PDF
                var outputFileName = $"document_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                var conversionResult = await ConvertToPdfWithColorPreservation(
                    filePath, outputFileName, selectedSheets);

                if (!conversionResult.Success || !System.IO.File.Exists(conversionResult.PdfFilePath))
                {
                    return Json(new { success = false, message = conversionResult.Message });
                }

                Console.WriteLine($"✅ Initial PDF created: {conversionResult.PdfFilePath}");

                // Get download PDF page count
                int downloadExcelPages = 0;
                using (var reader = new PdfReader(conversionResult.PdfFilePath))
                using (var pdfDoc = new PdfDocument(reader))
                {
                    downloadExcelPages = pdfDoc.GetNumberOfPages();
                }
                Console.WriteLine($"📊 Download Excel PDF: {downloadExcelPages} pages");

                // Step 4: Check for merged PDFs
                var mergedPdfsJson = HttpContext.Session.GetString(request.SessionId + "_mergedPdfs") ?? "[]";
                var mergedPdfs = JsonSerializer.Deserialize<List<MergedPdfInfo>>(mergedPdfsJson) ?? new List<MergedPdfInfo>();

                string finalPdfPath = conversionResult.PdfFilePath;
                int mergedPages = 0;

                // ✅ IMPORTANT: Use the LATEST merged PDF
                if (mergedPdfs.Any())
                {
                    var latestMergedPdf = mergedPdfs
                        .OrderByDescending(m => m.CreatedAt)
                        .FirstOrDefault();

                    if (latestMergedPdf != null && System.IO.File.Exists(latestMergedPdf.FilePath))
                    {
                        Console.WriteLine($"📊 Using LATEST merged PDF: {latestMergedPdf.FileName} ({latestMergedPdf.TotalPages} pages)");
                        mergedPages = latestMergedPdf.TotalPages;

                        // ✅ Use the SAME merge logic as preview
                        finalPdfPath = await MergeExcelWithLatestMergedPdf(conversionResult.PdfFilePath, latestMergedPdf.FilePath);
                        System.IO.File.Delete(conversionResult.PdfFilePath);
                        Console.WriteLine($"✅ Combined with latest uploaded PDF: {finalPdfPath}");
                    }
                    else
                    {
                        Console.WriteLine($"⚠️ Latest merged PDF not found or doesn't exist");
                    }
                }
                else
                {
                    Console.WriteLine($"ℹ️ No merged PDFs found");
                }

                // ✅ STEP 5: Calculate TOTAL download pages
                int downloadTotalPages = downloadExcelPages + mergedPages;
                Console.WriteLine($"📊 Total Download Pages: {downloadExcelPages} (Excel) + {mergedPages} (Merged) = {downloadTotalPages}");

                // ✅ STEP 6: ADJUST page order data for download PDF
                List<PageOrderInfoWithRotation> adjustedPageOrderData = request.PageOrderData ?? new List<PageOrderInfoWithRotation>();

                if (previewTotalPages > 0 && previewTotalPages != downloadTotalPages && request.PageOrderData != null)
                {
                    Console.WriteLine($"🔄 Adjusting page order from preview ({previewTotalPages}) to download ({downloadTotalPages})");
                    adjustedPageOrderData = AdjustPageOrderForDownload(
                        request.PageOrderData,
                        previewTotalPages,
                        downloadTotalPages);
                }

                // Step 7: Apply reordering, orientation and rotation
                string finalProcessedPath = await ApplySimpleReorderingOrientationAndRotation(
                    finalPdfPath,
                    adjustedPageOrderData,
                    request.OrientationData,
                    request.RotationData);

                // Step 8: Read final PDF
                var finalPdfBytes = await System.IO.File.ReadAllBytesAsync(finalProcessedPath);
                var finalPdfBase64 = Convert.ToBase64String(finalPdfBytes);

                // ✅ DEBUG: Get page count for verification
                try
                {
                    using (var reader = new PdfReader(finalProcessedPath))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        Console.WriteLine($"📊 Final Download PDF pages: {pdfDoc.GetNumberOfPages()}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Error counting final pages: {ex.Message}");
                }

                // Step 9: Cleanup
                if (System.IO.File.Exists(conversionResult.PdfFilePath))
                    System.IO.File.Delete(conversionResult.PdfFilePath);
                if (System.IO.File.Exists(finalPdfPath) && finalPdfPath != conversionResult.PdfFilePath)
                    System.IO.File.Delete(finalPdfPath);
                if (System.IO.File.Exists(finalProcessedPath))
                    System.IO.File.Delete(finalProcessedPath);

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


        private async Task<string> ApplyOnlyFitToPage(
            string pdfPath,
            List<PageOrderInfoWithRotation> pageOrderData)
        {
            var outputPath = Path.Combine(
                Path.GetTempPath(),
                $"preview_fittopage_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Applying FitToPage to existing PDF...");

                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(outputPath))
                using (var newPdf = new PdfDocument(writer))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    int totalSourcePages = sourcePdf.GetNumberOfPages();
                    Console.WriteLine($"📄 Source PDF: {totalSourcePages} pages");

                    // Use provided pageOrderData or show all
                    List<PageOrderInfoWithRotation> visiblePages;

                    if (pageOrderData != null && pageOrderData.Any())
                    {
                        visiblePages = pageOrderData
                            .Where(p => p.Visible)
                            .OrderBy(p => p.CurrentOrder)
                            .ToList();

                        // Remove duplicates
                        visiblePages = visiblePages
                            .GroupBy(p => p.OriginalPage)
                            .Select(g => g.First())
                            .ToList();
                    }
                    else
                    {
                        // Default: all pages visible
                        visiblePages = new List<PageOrderInfoWithRotation>();
                        for (int i = 1; i <= totalSourcePages; i++)
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
                    }

                    Console.WriteLine($"📄 Processing {visiblePages.Count} visible pages");

                    foreach (var pageInfo in visiblePages)
                    {
                        int sourcePageNum = pageInfo.OriginalPage;

                        if (sourcePageNum > 0 && sourcePageNum <= totalSourcePages)
                        {
                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
                            var sourcePageSize = sourcePage.GetPageSize();

                            // Get orientation and rotation
                            string orientation = pageInfo.Orientation ?? "portrait";
                            PageSize targetPageSize = orientation == "landscape" ? PageSize.A4.Rotate() : PageSize.A4;
                            int rotation = pageInfo.Rotation;

                            // Create new page
                            var newPage = newPdf.AddNewPage(targetPageSize);
                            var copiedPage = sourcePage.CopyAsFormXObject(newPdf);
                            var canvas = new PdfCanvas(newPage);

                            // Calculate scale to fit (90% of page with margin)
                            float margin = 20;
                            float targetWidth = targetPageSize.GetWidth();
                            float targetHeight = targetPageSize.GetHeight();
                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            float scale = 0.90f; // Fixed scale for consistency
                            float scaledWidth = sourceWidth * scale;
                            float scaledHeight = sourceHeight * scale;
                            float xOffset = margin + (targetWidth - (2 * margin) - scaledWidth) / 2;
                            float yOffset = margin + (targetHeight - (2 * margin) - scaledHeight) / 2;

                            // Apply transformations
                            canvas.SaveState();

                            if (rotation != 0)
                            {
                                canvas.ConcatMatrix(1, 0, 0, 1,
                                    xOffset + scaledWidth / 2,
                                    yOffset + scaledHeight / 2);

                                canvas.ConcatMatrix((float)Math.Cos(rotation * Math.PI / 180),
                                                    (float)Math.Sin(rotation * Math.PI / 180),
                                                    (float)-Math.Sin(rotation * Math.PI / 180),
                                                    (float)Math.Cos(rotation * Math.PI / 180),
                                                    0, 0);

                                canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
                                canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);
                            }
                            else
                            {
                                canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                            }

                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();

                            Console.WriteLine($"✅ Page {sourcePageNum}: Applied (Orientation: {orientation}, Rotation: {rotation}°)");
                        }
                    }

                    newPdf.Close();
                    sourcePdf.Close();
                }

                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplyOnlyFitToPage: {ex}");
                return pdfPath;
            }
        }





        // ✅ UPDATED: ApplySimpleReorderingOrientationAndRotation - Add debugging info
        private async Task<string> ApplySimpleReorderingOrientationAndRotation(
            string pdfPath,
            List<PageOrderInfoWithRotation> pageOrderData,
            Dictionary<int, string> orientationData,
            Dictionary<int, int> rotationData)
        {
            var outputPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"final_simple_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine("🔄 Applying SIMPLE page reordering, orientation and rotation...");

                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(outputPath))
                using (var newPdf = new PdfDocument(writer))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    int totalSourcePages = sourcePdf.GetNumberOfPages();
                    Console.WriteLine($"📄 Source PDF: {totalSourcePages} pages");

                    // Create default page order if not provided
                    List<PageOrderInfoWithRotation> visiblePages;

                    if (pageOrderData != null && pageOrderData.Any())
                    {
                        visiblePages = pageOrderData
                            .Where(p => p.Visible)
                            .OrderBy(p => p.CurrentOrder)
                            .ToList();

                        // ✅ DEBUG: Show what we're processing
                        //Console.WriteLine("📋 Page Order Data received:");
                        //foreach (var p in visiblePages)
                        //{
                        //    Console.WriteLine($"   - Order {p.CurrentOrder}: Page {p.OriginalPage} (Visible: {p.Visible}, Orientation: {p.Orientation}, Rotation: {p.Rotation}°)");
                        //}

                        // ✅ Check for duplicates
                        var uniquePages = new List<PageOrderInfoWithRotation>();
                        var seenPages = new HashSet<int>();

                        foreach (var page in visiblePages)
                        {
                            if (!seenPages.Contains(page.OriginalPage))
                            {
                                seenPages.Add(page.OriginalPage);
                                uniquePages.Add(page);
                            }
                            else
                            {
                                Console.WriteLine($"⚠️ Skipping duplicate page: {page.OriginalPage}");
                            }
                        }

                        if (visiblePages.Count != uniquePages.Count)
                        {
                            Console.WriteLine($"✅ Removed {visiblePages.Count - uniquePages.Count} duplicates");
                            visiblePages = uniquePages;
                        }
                    }
                    else
                    {
                        // Default: all pages visible
                        visiblePages = new List<PageOrderInfoWithRotation>();
                        for (int i = 1; i <= totalSourcePages; i++)
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
                    }

                    Console.WriteLine($"📄 Processing {visiblePages.Count} visible pages");

                    foreach (var pageInfo in visiblePages)
                    {
                        int sourcePageNum = pageInfo.OriginalPage;

                        // Check if page exists
                        if (sourcePageNum > 0 && sourcePageNum <= totalSourcePages)
                        {
                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
                            var sourcePageSize = sourcePage.GetPageSize();

                            // Get orientation
                            string orientation = pageInfo.Orientation ?? "portrait";
                            if (orientationData != null && orientationData.ContainsKey(sourcePageNum))
                            {
                                orientation = orientationData[sourcePageNum];
                            }

                            // Get rotation
                            int rotation = pageInfo.Rotation;
                            if (rotationData != null && rotationData.ContainsKey(sourcePageNum))
                            {
                                rotation = rotationData[sourcePageNum];
                            }

                            // Create page with orientation
                            PageSize targetPageSize = orientation == "landscape" ? PageSize.A4.Rotate() : PageSize.A4;
                            var newPage = newPdf.AddNewPage(targetPageSize);

                            // Copy content
                            var copiedPage = sourcePage.CopyAsFormXObject(newPdf);
                            var canvas = new PdfCanvas(newPage);

                            // ✅ SAME fitting calculation as preview
                            float margin = 20;
                            float targetWidth = targetPageSize.GetWidth();
                            float targetHeight = targetPageSize.GetHeight();
                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            // Calculate scale to fit
                            float scaleX = (targetWidth - (2 * margin)) / sourceWidth;
                            float scaleY = (targetHeight - (2 * margin)) / sourceHeight;
                            //float scale = Math.Min(scaleX, scaleY);
                            float scale = 0.90f;


                            // ✅ Ensure scale is reasonable (0.8 - 1.0 range)
                            //if (scale < 0.8f)
                            //{
                            //    Console.WriteLine($"⚠️ Adjusting low scale {scale:F2} to 0.80 for page {sourcePageNum}");
                            //    scale = 0.8f;
                            //}
                            //else if (scale > 1.0f)
                            //{
                            //    Console.WriteLine($"⚠️ Adjusting high scale {scale:F2} to 1.00 for page {sourcePageNum}");
                            //    scale = 1.0f;
                            //}


                            // ✅ SAME rotation adjustment as preview
                            //if (rotation != 0)
                            //{
                            //    margin = 15;
                            //    // Rotated pages need extra space
                            //    scale *= 0.95f; // 95% of calculated scale for rotated
                            //}

                            // Calculate centered position
                            float scaledWidth = sourceWidth * scale;
                            float scaledHeight = sourceHeight * scale;
                            float xOffset = margin + (targetWidth - (2 * margin) - scaledWidth) / 2;
                            float yOffset = margin + (targetHeight - (2 * margin) - scaledHeight) / 2;

                            // ✅ SAME transformation logic as preview
                            canvas.SaveState();

                            // Apply rotation at center
                            if (rotation != 0)
                            {
                                canvas.ConcatMatrix(1, 0, 0, 1,
                                    xOffset + scaledWidth / 2,
                                    yOffset + scaledHeight / 2);

                                canvas.ConcatMatrix((float)Math.Cos(rotation * Math.PI / 180),
                                                    (float)Math.Sin(rotation * Math.PI / 180),
                                                    (float)-Math.Sin(rotation * Math.PI / 180),
                                                    (float)Math.Cos(rotation * Math.PI / 180),
                                                    0, 0);

                                //canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);

                                canvas.ConcatMatrix(1, 0, 0, 1, -scaledWidth / 2, -scaledHeight / 2);
                                canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);
                            }
                            else
                            {
                                // No rotation - simple scaling and positioning
                                canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                            }

                            // Apply scaling and positioning
                            //if (rotation == 0)
                            //{
                            //    canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                            //}
                            //else
                            //{
                            //    // If rotated, scaling is already applied in the transformation
                            //    canvas.ConcatMatrix(scale, 0, 0, scale, 0, 0);
                            //}

                            // Draw content
                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();

                            Console.WriteLine($"✅ Page {sourcePageNum} → Position {pageInfo.CurrentOrder} (Orientation: {orientation}, Rotation: {rotation}°, Scale: {scale:F2})");
                        }
                        else
                        {
                            Console.WriteLine($"⚠️ Page {sourcePageNum} not found in source PDF (Total: {totalSourcePages})");

                            // Add empty placeholder page
                            PageSize targetPageSize = PageSize.A4;
                            var newPage = newPdf.AddNewPage(targetPageSize);
                            var canvas = new PdfCanvas(newPage);
                            canvas.BeginText()
                                  .MoveText(50, targetPageSize.GetHeight() - 100)
                                  .SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(), 12)
                                  .ShowText($"Page {sourcePageNum} - Missing in source")
                                  .EndText();
                            canvas.Release();

                            Console.WriteLine($"⚠️ Added placeholder for missing page {sourcePageNum}");
                        }
                    }

                    newPdf.Close();
                    sourcePdf.Close();
                }

                Console.WriteLine($"✅ Simple processed PDF created: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplySimpleReorderingOrientationAndRotation: {ex}");
                return pdfPath;
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
