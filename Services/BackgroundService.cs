using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ExcelToPdfConverter.Services
{
    public class PdfCleanupService : BackgroundService
    {
        private readonly ILogger<PdfCleanupService> _logger;
        private readonly IWebHostEnvironment _environment;

        public PdfCleanupService(ILogger<PdfCleanupService> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("PDF Cleanup Service started.");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await CleanupOldPreviewFiles();
                    await Task.Delay(TimeSpan.FromHours(1), stoppingToken); // हर 1 घंटे में चलाएं
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in PDF cleanup service");
                    await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
                }
            }
        }

        private async Task CleanupOldPreviewFiles()
        {
            try
            {
                var previewsDir = Path.Combine(_environment.WebRootPath, "previews");
                if (!Directory.Exists(previewsDir))
                    return;

                var files = Directory.GetFiles(previewsDir, "*.pdf");
                int deletedCount = 0;

                foreach (var file in files)
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.CreationTime < DateTime.Now.AddHours(-2)) // 2 घंटे से पुरानी फाइलें
                        {
                            fileInfo.Delete();
                            deletedCount++;
                            _logger.LogInformation($"Deleted old preview file: {fileInfo.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"Error deleting file: {file}");
                    }
                }

                if (deletedCount > 0)
                {
                    _logger.LogInformation($"Cleaned up {deletedCount} old preview files.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in preview files cleanup");
            }

            await Task.CompletedTask;
        }
    }
}