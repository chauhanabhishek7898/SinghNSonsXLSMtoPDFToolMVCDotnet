using ExcelToPdfConverter.Models;
using System.Diagnostics;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Utils;

namespace ExcelToPdfConverter.Services
{
    public class LibreOfficeService
    {
        private readonly string _libreOfficePath;
        private readonly string _tempDirectory;
        private readonly string _outputDirectory;
        private readonly IWebHostEnvironment _environment;

        public LibreOfficeService(IWebHostEnvironment environment)
        {
            _environment = environment;
            _libreOfficePath = GetLibreOfficePath();
            _tempDirectory = System.IO.Path.Combine(_environment.WebRootPath, "App_Data", "Temp");
            _outputDirectory = System.IO.Path.Combine(_environment.WebRootPath, "App_Data", "Output");
            Directory.CreateDirectory(_tempDirectory);
            Directory.CreateDirectory(_outputDirectory);
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

        public async Task<ConversionResult> ConvertToPdfAsync(string inputFilePath, string outputFileName,
            List<string>? selectedSheets = null, Dictionary<string, string>? sheetOrientations = null)
        {
            var result = new ConversionResult();
            try
            {
                Console.WriteLine($"=== Starting PDF Conversion ===");
                Console.WriteLine($"Input file: {inputFilePath}");
                Console.WriteLine($"Output file name: {outputFileName}");
                Console.WriteLine($"Selected sheets: {(selectedSheets != null ? string.Join(", ", selectedSheets) : "All")}");
                Console.WriteLine($"Sheet orientations: {(sheetOrientations != null ? string.Join(", ", sheetOrientations) : "Default")}");
                Console.WriteLine($"File exists: {System.IO.File.Exists(inputFilePath)}");

                // Use user's home directory for output
                var outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var outputFilePath = System.IO.Path.Combine(outputDirectory, outputFileName);

                Console.WriteLine($"Output directory: {outputDirectory}");
                Console.WriteLine($"Full output path: {outputFilePath}");

                // Build LibreOffice arguments with sheet selection
                var arguments = BuildLibreOfficeArguments(inputFilePath, outputDirectory, selectedSheets);
                Console.WriteLine($"LibreOffice arguments: {arguments}");

                var processStartInfo = new ProcessStartInfo
                {
                    FileName = _libreOfficePath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = outputDirectory
                };

                Console.WriteLine($"LibreOffice path: {_libreOfficePath}");

                using (var process = new Process())
                {
                    process.StartInfo = processStartInfo;
                    Console.WriteLine("Starting LibreOffice process...");
                    process.Start();

                    // Read outputs
                    string output = await process.StandardOutput.ReadToEndAsync();
                    string error = await process.StandardError.ReadToEndAsync();

                    Console.WriteLine($"LibreOffice Output: {output}");
                    if (!string.IsNullOrEmpty(error))
                        Console.WriteLine($"LibreOffice Error: {error}");

                    // Wait with timeout
                    bool processExited = process.WaitForExit(120000); // 2 minutes
                    Console.WriteLine($"Process exited: {processExited}, Exit code: {process.ExitCode}");

                    if (processExited && process.ExitCode == 0)
                    {
                        // Check multiple possible output locations
                        var inputFileName = System.IO.Path.GetFileNameWithoutExtension(inputFilePath);
                        var possibleOutputPaths = new[]
                        {
                            outputFilePath,
                            System.IO.Path.Combine(outputDirectory, inputFileName + ".pdf"),
                            System.IO.Path.Combine(System.IO.Path.GetDirectoryName(inputFilePath) ?? "", inputFileName + ".pdf")
                        };

                        string? foundPath = null;
                        foreach (var path in possibleOutputPaths)
                        {
                            if (System.IO.File.Exists(path))
                            {
                                foundPath = path;
                                Console.WriteLine($"✅ PDF found at: {path}");
                                break;
                            }
                        }

                        if (foundPath != null)
                        {
                            result.Success = true;
                            result.Message = "Conversion successful";
                            result.PdfFilePath = foundPath;
                            result.FileName = System.IO.Path.GetFileName(foundPath);
                        }
                        else
                        {
                            result.Success = false;
                            result.Message = "Conversion completed but PDF file not found";
                            Console.WriteLine("❌ PDF file not found in any expected location");
                        }
                    }
                    else
                    {
                        result.Success = false;
                        result.Message = $"Process failed or timed out. Exit code: {process.ExitCode}";
                        Console.WriteLine($"❌ Process failed: {result.Message}");
                    }
                }

                // Cleanup input file only if conversion successful
                if (result.Success && System.IO.File.Exists(inputFilePath))
                {
                    System.IO.File.Delete(inputFilePath);
                    Console.WriteLine($"Cleaned up input file: {inputFilePath}");
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error during conversion: {ex.Message}";
                Console.WriteLine($"❌ Conversion error: {ex}");
            }

            Console.WriteLine($"=== Conversion Result: {result.Success} ===");
            return result;
        }

        private string BuildLibreOfficeArguments(string inputFilePath, string outputDirectory, List<string>? selectedSheets)
        {
            var arguments = new List<string>
            {
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                "--convert-to pdf:calc_pdf_Export",
                $"--outdir \"{outputDirectory}\""
            };

            arguments.Add($"\"{inputFilePath}\"");
            return string.Join(" ", arguments);
        }

        private async Task<string> ApplyPageOrientations(string pdfPath, Dictionary<string, string> sheetOrientations, List<string> selectedSheets)
        {
            string? tempOutputPath = null;

            try
            {
                Console.WriteLine("🔄 Applying page orientations using iText7...");

                if (!System.IO.File.Exists(pdfPath))
                {
                    Console.WriteLine("❌ PDF file not found for orientation application");
                    return pdfPath;
                }

                if (sheetOrientations == null || !sheetOrientations.Any())
                {
                    Console.WriteLine("ℹ️ No orientations specified, skipping");
                    return pdfPath;
                }

                tempOutputPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"oriented_{Guid.NewGuid()}.pdf");
                Console.WriteLine($"📁 Temp output path: {tempOutputPath}");

                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(tempOutputPath))
                using (var newPdfDoc = new PdfDocument(writer))
                using (var sourcePdfDoc = new PdfDocument(reader))
                {
                    var numberOfPages = sourcePdfDoc.GetNumberOfPages();
                    Console.WriteLine($"📄 Processing {numberOfPages} pages");

                    // ✅ NEW: Improved page-to-sheet mapping for multi-page sheets
                    var sheetPageMapping = CreateSheetPageMapping(selectedSheets, numberOfPages);

                    Console.WriteLine("📋 Sheet to Page Mapping:");
                    foreach (var mapping in sheetPageMapping)
                    {
                        Console.WriteLine($"   {mapping.Key}: Pages {mapping.Value.Start}-{mapping.Value.End}");
                    }

                    // Process each sheet and apply orientation to ALL its pages
                    foreach (var mapping in sheetPageMapping)
                    {
                        string sheetName = mapping.Key;
                        var pageRange = mapping.Value;
                        string orientation = GetOrientationForSheet(sheetName, sheetOrientations);

                        Console.WriteLine($"🎯 Applying '{orientation}' to {sheetName} (Pages {pageRange.Start}-{pageRange.End})");

                        // ✅ Apply same orientation to ALL pages of this sheet
                        for (int pageNum = pageRange.Start; pageNum <= pageRange.End; pageNum++)
                        {
                            ApplyOrientationToPage(sourcePdfDoc, newPdfDoc, pageNum, orientation, sheetName);
                        }
                    }

                    newPdfDoc.Close();
                    sourcePdfDoc.Close();
                }

                // Replace original file
                System.IO.File.Delete(pdfPath);
                System.IO.File.Move(tempOutputPath, pdfPath);

                Console.WriteLine("✅ Page orientations applied successfully to all sheets");
                return pdfPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error applying orientations: {ex.Message}");

                // Cleanup temp file
                if (tempOutputPath != null && System.IO.File.Exists(tempOutputPath))
                {
                    try { System.IO.File.Delete(tempOutputPath); } catch { }
                }

                return pdfPath;
            }
        }

        // ✅ NEW: Better page mapping that handles multi-page sheets
        private Dictionary<string, PageRange> CreateSheetPageMapping(List<string> selectedSheets, int totalPages)
        {
            var mapping = new Dictionary<string, PageRange>();

            if (selectedSheets != null && selectedSheets.Count > 0)
            {
                // Simple distribution: each sheet gets equal pages
                // For more advanced mapping, you might need page break detection
                int pagesPerSheet = (int)Math.Ceiling((double)totalPages / selectedSheets.Count);

                for (int i = 0; i < selectedSheets.Count; i++)
                {
                    int startPage = (i * pagesPerSheet) + 1;
                    int endPage = Math.Min((i + 1) * pagesPerSheet, totalPages);

                    if (startPage <= totalPages)
                    {
                        mapping[selectedSheets[i]] = new PageRange { Start = startPage, End = endPage };
                    }
                }
            }
            else
            {
                // Fallback: treat each page as separate sheet
                for (int i = 1; i <= totalPages; i++)
                {
                    mapping[$"Page_{i}"] = new PageRange { Start = i, End = i };
                }
            }

            return mapping;
        }

        // ✅ NEW: Clean orientation application without 90° rotation complexity
        private void ApplyOrientationToPage(PdfDocument sourceDoc, PdfDocument targetDoc, int pageNum, string orientation, string sheetName)
        {
            try
            {
                var sourcePage = sourceDoc.GetPage(pageNum);
                var sourcePageSize = sourcePage.GetPageSize();

                Console.WriteLine($"   📄 Processing Page {pageNum} - {sheetName}");
                Console.WriteLine($"      Source: {sourcePageSize.GetWidth()} x {sourcePageSize.GetHeight()}");

                // ✅ SIMPLE: Determine target page size based on orientation
                PageSize targetPageSize = orientation == "Landscape" ? PageSize.A4.Rotate() : PageSize.A4;

                Console.WriteLine($"      Target: {targetPageSize.GetWidth()} x {targetPageSize.GetHeight()} ({orientation})");

                // Create new page with target size
                var newPage = targetDoc.AddNewPage(targetPageSize);

                // Copy content from source page
                var copiedPage = sourcePage.CopyAsFormXObject(targetDoc);

                // Create canvas and add content (centered)
                var canvas = new PdfCanvas(newPage);

                // Calculate centering offsets
                float xOffset = (targetPageSize.GetWidth() - sourcePageSize.GetWidth()) / 2;
                float yOffset = (targetPageSize.GetHeight() - sourcePageSize.GetHeight()) / 2;

                // Add content centered on page
                canvas.AddXObjectAt(copiedPage, xOffset, yOffset);

                Console.WriteLine($"      ✅ Applied {orientation} (Centered at: {xOffset:F1}, {yOffset:F1})");

                canvas.Release();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error processing page {pageNum}: {ex.Message}");
                throw;
            }
        }

        // ✅ SIMPLIFIED: Get orientation for sheet
        private string GetOrientationForSheet(string sheetName, Dictionary<string, string> sheetOrientations)
        {
            if (sheetOrientations.ContainsKey(sheetName))
                return sheetOrientations[sheetName];

            return "Portrait"; // Default fallback
        }

        // Helper class for page ranges
        public class PageRange
        {
            public int Start { get; set; }
            public int End { get; set; }
        }

        private string GetDefaultOrientation(string? sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                return "Portrait";

            // Default orientations based on your requirements
            var portraitSheets = new List<string>
            {
                "Sheet1", "Sheet2", "Sheet3", "Sheet6"  // 1st, 2nd, 3rd, 6th sheets - Portrait
            };

            var landscapeSheets = new List<string>
            {
                "Sheet4", "Sheet5", "Sheet7", "Sheet8"  // 4th, 5th, 7th, 8th sheets - Landscape
            };

            // Check if sheet name contains pattern (like "Sheet1", "Sheet2", etc.)
            if (portraitSheets.Any(ps => sheetName.Contains(ps, StringComparison.OrdinalIgnoreCase) ||
                                        sheetName.Equals(ps, StringComparison.OrdinalIgnoreCase)))
                return "Portrait";

            if (landscapeSheets.Any(ls => sheetName.Contains(ls, StringComparison.OrdinalIgnoreCase) ||
                                         sheetName.Equals(ls, StringComparison.OrdinalIgnoreCase)))
                return "Landscape";

            // Default for other sheets
            return "Portrait";
        }

        private bool ShouldRotate90Degrees(string? sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                return false;

            // Sheets that should be rotated 90 degrees when in Portrait
            var rotateSheets = new List<string>
            {
                "AOC",
                "DIFF",
                "FORM B",
                "Wage Reg"
                // Add other sheet names that need 90-degree rotation
            };

            return rotateSheets.Any(rs => sheetName.Contains(rs, StringComparison.OrdinalIgnoreCase));
        }

        private string? GetSheetNameForPage(List<string>? selectedSheets, int pageNum, int totalPages)
        {
            try
            {
                if (selectedSheets == null || !selectedSheets.Any())
                    return null;

                if (pageNum <= selectedSheets.Count)
                {
                    return selectedSheets[pageNum - 1];
                }

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in GetSheetNameForPage: {ex.Message}");
                return null;
            }
        }

        public string SaveUploadedFile(IFormFile file)
        {
            var fileName = Guid.NewGuid() + System.IO.Path.GetExtension(file.FileName);
            var filePath = System.IO.Path.Combine(_tempDirectory, fileName);
            using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
            {
                file.CopyTo(stream);
            }
            Console.WriteLine($"File saved for conversion: {filePath}");
            return filePath;
        }

        public void CleanupOldFiles(int hoursOld = 24)
        {
            try
            {
                CleanupDirectory(_tempDirectory, hoursOld);
                CleanupDirectory(_outputDirectory, hoursOld);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cleanup error: {ex.Message}");
            }
        }

        private void CleanupDirectory(string directory, int hoursOld)
        {
            if (!Directory.Exists(directory)) return;

            var cutoffTime = DateTime.Now.AddHours(-hoursOld);
            var files = Directory.GetFiles(directory);
            foreach (var file in files)
            {
                try
                {
                    var fileInfo = new System.IO.FileInfo(file);
                    if (fileInfo.LastWriteTime < cutoffTime)
                    {
                        fileInfo.Delete();
                        Console.WriteLine($"Cleaned up old file: {file}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting file {file}: {ex.Message}");
                }
            }
        }
    }
}
