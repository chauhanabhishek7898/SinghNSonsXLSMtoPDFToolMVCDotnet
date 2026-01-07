using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Pdf.Canvas;
//using iText.Kernel.Pdf.Filter;

namespace ExcelToPdfConverter.Services
{
    public class PdfCompressionService
    {
        private readonly ILogger<PdfCompressionService> _logger;

        public PdfCompressionService(ILogger<PdfCompressionService> logger)
        {
            _logger = logger;
        }

        public async Task<string> CompressPdfIfLarge(string pdfPath, int maxSizeMB = 10)
        {
            try
            {
                var fileInfo = new FileInfo(pdfPath);
                var fileSizeMB = fileInfo.Length / (1024.0 * 1024.0);

                _logger.LogInformation($"PDF size: {fileSizeMB:F2} MB, Max allowed: {maxSizeMB} MB");

                if (fileSizeMB <= maxSizeMB)
                {
                    return pdfPath; // No compression needed
                }

                _logger.LogInformation($"Compressing PDF: {pdfPath}");

                var compressedPath = Path.Combine(Path.GetTempPath(), $"compressed_{Guid.NewGuid()}.pdf");
                await CompressPdfWithImages(pdfPath, compressedPath, maxSizeMB);

                // Check if compression was successful
                var compressedInfo = new FileInfo(compressedPath);
                if (compressedInfo.Exists && compressedInfo.Length < fileInfo.Length)
                {
                    _logger.LogInformation($"Compression successful: {fileSizeMB:F2} MB -> {compressedInfo.Length / (1024.0 * 1024.0):F2} MB");
                    File.Delete(pdfPath);
                    return compressedPath;
                }

                return pdfPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error compressing PDF: {pdfPath}");
                return pdfPath; // Return original if compression fails
            }
        }

        private async Task CompressPdfWithImages(string sourcePath, string destPath, int maxSizeMB)
        {
            using (var reader = new PdfReader(sourcePath))
            using (var writer = new PdfWriter(destPath, new WriterProperties()
                .SetCompressionLevel(CompressionConstants.BEST_COMPRESSION)
                .SetFullCompressionMode(true)))
            using (var pdfDoc = new PdfDocument(reader, writer))
            {
                int totalPages = pdfDoc.GetNumberOfPages();

                for (int i = 1; i <= totalPages; i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var resources = page.GetResources();

                    // Optimize images
                    await OptimizePageImages(resources);

                    // Update progress
                    if (i % 10 == 0)
                    {
                        _logger.LogInformation($"Processing page {i}/{totalPages}");
                    }
                }

                pdfDoc.Close();
            }
        }

        private async Task OptimizePageImages(iText.Kernel.Pdf.PdfResources resources)
        {
            try
            {
                var xObject = resources.GetResource(iText.Kernel.Pdf.PdfName.XObject);
                if (xObject == null) return;

                foreach (var key in xObject.KeySet())
                {
                    var pdfObject = xObject.Get(key);
                    if (pdfObject is iText.Kernel.Pdf.PdfStream pdfStream)
                    {
                        var subtype = pdfStream.Get(iText.Kernel.Pdf.PdfName.Subtype);
                        if (subtype != null && subtype.ToString() == "/Image")
                        {
                            // Apply image compression
                            await ApplyImageCompression(pdfStream);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error optimizing images: {ex.Message}");
            }
        }

        private async Task ApplyImageCompression(iText.Kernel.Pdf.PdfStream pdfStream)
        {
            try
            {
                // Check current image size
                var currentData = pdfStream.GetBytes();
                if (currentData.Length < 1024 * 100) // Skip small images (< 100KB)
                    return;

                // Apply DCTDecode filter (JPEG compression)
                var filters = pdfStream.Get(iText.Kernel.Pdf.PdfName.Filter);
                if (filters == null)
                {
                    pdfStream.SetData(currentData, true);
                    pdfStream.Put(iText.Kernel.Pdf.PdfName.Filter, iText.Kernel.Pdf.PdfName.DCTDecode);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error applying image compression: {ex.Message}");
            }
        }

        // ✅ Quick compression for large files
        public async Task<string> QuickCompressPdf(string pdfPath)
        {
            var outputPath = Path.Combine(Path.GetTempPath(), $"quick_compress_{Guid.NewGuid()}.pdf");

            try
            {
                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(outputPath, new WriterProperties()
                    .SetCompressionLevel(CompressionConstants.DEFAULT_COMPRESSION)))
                using (var pdfDoc = new PdfDocument(reader, writer))
                {
                    // Just copy with default compression
                    pdfDoc.Close();
                }

                return outputPath;
            }
            catch
            {
                return pdfPath;
            }
        }
    }
}