using ExcelToPdfConverter.Models;
using ExcelToPdfConverter.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using iText.Kernel.Pdf;

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

                // Quick validation
                var validationResult = _previewService.QuickValidate(model.ExcelFile);
                ViewBag.ValidationResult = validationResult;

                // Generate preview
                var sessionId = Guid.NewGuid().ToString();
                var previewModel = _previewService.GeneratePreview(model.ExcelFile, sessionId);

                // Save file for conversion
                var filePath = _libreOfficeService.SaveUploadedFile(model.ExcelFile);

                // Store in session
                HttpContext.Session.SetString(sessionId + "_filePath", filePath);
                HttpContext.Session.SetString(sessionId + "_fileName", model.ExcelFile.FileName ?? "unknown");

                // ✅ Session ID ko ViewBag mein bhi store karein for client-side access
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

                // ✅ STEP 1: Create ordered sheets list based on drag & drop order
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

                // ✅ STEP 2: Process Excel file - CREATE NEW FILE with selected sheets and order
                Console.WriteLine($"🔄 Starting Excel file processing...");
                string processedFilePath;
                try
                {
                    // ✅ USE THE BETTER APPROACH - Create new file
                    processedFilePath = await _excelProcessingService.CreateProcessedExcelFileAsync(
                        filePath, orderedSheets, orderedSheets);

                    Console.WriteLine($"✅ Excel file processed successfully: {processedFilePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Excel processing failed, using original file: {ex.Message}");
                    processedFilePath = filePath; // Fallback to original file
                }

                // ✅ STEP 3: Create orientation mapping
                var orientationMap = new Dictionary<string, string>();
                if (selectedSheets != null && sheetOrientations != null && selectedSheets.Count == sheetOrientations.Count)
                {
                    orientationMap = selectedSheets.Zip(sheetOrientations, (s, o) => new { Sheet = s, Orientation = o })
                                                 .ToDictionary(x => x.Sheet, x => x.Orientation);
                    Console.WriteLine($"✅ Orientation mapping: {string.Join(", ", orientationMap)}");
                }

                var outputFileName = System.IO.Path.GetFileNameWithoutExtension(originalFileName ?? "converted") + ".pdf";

                // ✅ STEP 4: Convert PROCESSED Excel to PDF
                Console.WriteLine($"🔄 Starting Excel to PDF conversion...");
                var result = await _libreOfficeService.ConvertToPdfAsync(
                    processedFilePath, outputFileName, orderedSheets, orientationMap);

                // ✅ STEP 5: Cleanup processed Excel file
                if (processedFilePath != filePath) // Only cleanup if it's not the original file
                {
                    _excelProcessingService.CleanupProcessedFile(processedFilePath);
                }

                if (result.Success)
                {
                    Console.WriteLine($"✅ Excel to PDF conversion successful: {result.PdfFilePath}");

                    // ✅ STEP 6: Merge with existing PDFs
                    Console.WriteLine($"🔄 Starting PDF merge process...");
                    var finalResult = await MergeAllPdfsWithiText7(result.PdfFilePath, outputFileName, orderedSheets, orientationMap);

                    // Cleanup session
                    HttpContext.Session.Remove(sessionId + "_filePath");
                    HttpContext.Session.Remove(sessionId + "_fileName");

                    if (finalResult.Success)
                    {
                        Console.WriteLine($"✅ Final PDF created: {finalResult.PdfFilePath} ({finalResult.TotalPages} pages)");

                        var fileBytes = await System.IO.File.ReadAllBytesAsync(finalResult.PdfFilePath);

                        // Cleanup temporary files
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
                        // Fallback to original converted PDF
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
                    // Step 1: Add the newly converted Excel PDF
                    if (System.IO.File.Exists(newPdfPath))
                    {
                        using (var reader = new PdfReader(newPdfPath))
                        using (var sourceDoc = new PdfDocument(reader))
                        {
                            sourceDoc.CopyPagesTo(1, sourceDoc.GetNumberOfPages(), mergedPdfDoc);
                            Console.WriteLine($"✅ Added converted Excel PDF: {newPdfPath} ({sourceDoc.GetNumberOfPages()} pages)");
                        }
                    }

                    // Step 2: Add existing PDFs from directory
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

                // Fallback to original PDF
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


        [HttpPost]
        public IActionResult GetFileNames(string sessionId)
        {
            try
            {
                Console.WriteLine($"=== GetFileNames Called ===");
                Console.WriteLine($"Session ID: {sessionId}");

                var fileNamesModel = new FileNamesModel();

                // Get Excel file name from session
                var excelFileName = HttpContext.Session.GetString(sessionId + "_fileName");
                if (!string.IsNullOrEmpty(excelFileName))
                {
                    fileNamesModel.ExcelFileName = Path.GetFileNameWithoutExtension(excelFileName);
                    Console.WriteLine($"✅ Excel File: {fileNamesModel.ExcelFileName}");
                }

                // Get PDF file names from directory
                var pdfDirectory = @"D:\CIPL\SinghAndSons\pdf";
                if (Directory.Exists(pdfDirectory))
                {
                    var pdfFiles = Directory.GetFiles(pdfDirectory, "*.pdf")
                        .Select(Path.GetFileNameWithoutExtension)
                        .Where(name => !string.IsNullOrEmpty(name))
                        .ToList();

                    fileNamesModel.PdfFileNames = pdfFiles!;
                    fileNamesModel.TotalPdfFiles = pdfFiles.Count;

                    Console.WriteLine($"✅ Found {pdfFiles.Count} PDF files in directory");
                    foreach (var pdfFile in pdfFiles)
                    {
                        Console.WriteLine($"   📄 {pdfFile}");
                    }
                }
                else
                {
                    Console.WriteLine($"⚠️ PDF directory not found: {pdfDirectory}");
                }

                return Json(new
                {
                    success = true,
                    data = fileNamesModel
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GetFileNames: {ex}");
                return Json(new
                {
                    success = false,
                    message = $"Error getting file names: {ex.Message}"
                });
            }
        }
    }
}
