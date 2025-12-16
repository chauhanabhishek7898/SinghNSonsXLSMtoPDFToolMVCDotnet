using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using ExcelToPdfConverter.Models;
using SystemIO = System.IO; // Alias to avoid conflict

namespace ExcelToPdfConverter.Services
{
    public class PdfProcessingService
    {
        private readonly IWebHostEnvironment _environment;
        private readonly ILogger<PdfProcessingService> _logger;

        public PdfProcessingService(IWebHostEnvironment environment, ILogger<PdfProcessingService> logger)
        {
            _environment = environment;
            _logger = logger;
        }

        public async Task<string> ApplyFitToPageScaling(string pdfPath)
        {
            var outputPath = SystemIO.Path.Combine(SystemIO.Path.GetTempPath(), $"scaled_{Guid.NewGuid()}.pdf");

            try
            {
                Console.WriteLine($"🔄 Applying fitToPage scaling to PDF...");

                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(outputPath))
                using (var newPdf = new PdfDocument(writer))
                using (var sourcePdf = new PdfDocument(reader))
                {
                    int totalPages = sourcePdf.GetNumberOfPages();
                    Console.WriteLine($"📄 Processing {totalPages} pages");

                    for (int pageNum = 1; pageNum <= totalPages; pageNum++)
                    {
                        var sourcePage = sourcePdf.GetPage(pageNum);
                        var sourcePageSize = sourcePage.GetPageSize();

                        // Determine orientation
                        bool isLandscape = sourcePageSize.GetWidth() > sourcePageSize.GetHeight();
                        PageSize targetPageSize = isLandscape ? PageSize.A4.Rotate() : PageSize.A4;

                        // Create new page
                        var newPage = newPdf.AddNewPage(targetPageSize);

                        // Copy content with fitToPage scaling
                        var copiedPage = sourcePage.CopyAsFormXObject(newPdf);
                        var canvas = new PdfCanvas(newPage);

                        // Calculate scaling to fit page
                        float sourceWidth = sourcePageSize.GetWidth();
                        float sourceHeight = sourcePageSize.GetHeight();
                        float targetWidth = targetPageSize.GetWidth();
                        float targetHeight = targetPageSize.GetHeight();

                        // Apply margin
                        float margin = 20; // 20 points margin
                        float availableWidth = targetWidth - (2 * margin);
                        float availableHeight = targetHeight - (2 * margin);

                        // Calculate scale to fit
                        float scaleX = availableWidth / sourceWidth;
                        float scaleY = availableHeight / sourceHeight;
                        float scale = Math.Min(scaleX, scaleY);

                        // Calculate centered position
                        float scaledWidth = sourceWidth * scale;
                        float scaledHeight = sourceHeight * scale;
                        float xOffset = margin + (availableWidth - scaledWidth) / 2;
                        float yOffset = margin + (availableHeight - scaledHeight) / 2;

                        // Apply scaling and centering
                        canvas.SaveState();
                        canvas.ConcatMatrix(scale, 0, 0, scale, xOffset, yOffset);
                        canvas.AddXObjectAt(copiedPage, 0, 0);
                        canvas.RestoreState();

                        canvas.Release();

                        Console.WriteLine($"✅ Page {pageNum}: Scaled to fit {(isLandscape ? "Landscape" : "Portrait")} A4");
                    }

                    newPdf.Close();
                    sourcePdf.Close();
                }

                Console.WriteLine($"✅ FitToPage scaling applied: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplyFitToPageScaling: {ex}");
                // Return original if processing fails
                return pdfPath;
            }
        }

        public async Task<string> ApplyPageReorderingOrientationAndRotation(
            string pdfPath,
            List<PageOrderInfoWithRotation> pageOrderData,
            Dictionary<int, string> orientationData,
            Dictionary<int, int> rotationData)
        {
            var outputPath = SystemIO.Path.Combine(SystemIO.Path.GetTempPath(), $"final_processed_{Guid.NewGuid()}.pdf");

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
                            var sourcePage = sourcePdf.GetPage(sourcePageNum);
                            var sourcePageSize = sourcePage.GetPageSize();

                            // Get orientation and rotation
                            string orientation = pageInfo.Orientation;
                            if (orientationData != null && orientationData.ContainsKey(sourcePageNum))
                            {
                                orientation = orientationData[sourcePageNum];
                            }

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

                            // Get dimensions
                            float targetWidth = targetPageSize.GetWidth();
                            float targetHeight = targetPageSize.GetHeight();
                            float sourceWidth = sourcePageSize.GetWidth();
                            float sourceHeight = sourcePageSize.GetHeight();

                            // Calculate scaling with margins
                            float margin = 20;
                            float availableWidth = targetWidth - (2 * margin);
                            float availableHeight = targetHeight - (2 * margin);

                            // Calculate scale
                            float scaleX = availableWidth / sourceWidth;
                            float scaleY = availableHeight / sourceHeight;
                            float scale = Math.Min(scaleX, scaleY);

                            // Adjust for rotation
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

                            // Apply transformations
                            canvas.SaveState();

                            // Move to center
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

                            // Draw content
                            canvas.AddXObjectAt(copiedPage, 0, 0);
                            canvas.RestoreState();

                            canvas.Release();

                            Console.WriteLine($"✅ Page {sourcePageNum} → Order {pageInfo.CurrentOrder} ({orientation}, {rotation}°)");
                        }
                    }

                    newPdf.Close();
                    sourcePdf.Close();
                }

                Console.WriteLine($"✅ PDF with all modifications created: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ApplyPageReorderingOrientationAndRotation: {ex}");
                throw;
            }
        }
    }
}