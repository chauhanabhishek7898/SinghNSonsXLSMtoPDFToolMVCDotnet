using ExcelToPdfConverter.Models;
using System.Diagnostics;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Xobject;

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
                            // Apply orientations if specified using iText7
                            if (sheetOrientations != null && sheetOrientations.Any())
                            {
                                Console.WriteLine($"🔄 Applying page orientations and rotations...");
                                foundPath = await ApplyPageOrientations(foundPath, sheetOrientations, selectedSheets);
                            }

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
                "--convert-to pdf",
                $"--outdir \"{outputDirectory}\""
            };

            if (selectedSheets != null && selectedSheets.Any())
            {
                Console.WriteLine($"Sheet selection specified: {string.Join(", ", selectedSheets)}");
            }

            arguments.Add($"\"{inputFilePath}\"");
            return string.Join(" ", arguments);
        }

        private async Task<string> ApplyPageOrientations(string pdfPath, Dictionary<string, string> sheetOrientations, List<string>? selectedSheets)
        {
            string? tempOutputPath = null;

            try
            {
                Console.WriteLine("🔄 Applying page orientations and rotations using iText7...");

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

                    for (int pageNum = 1; pageNum <= numberOfPages; pageNum++)
                    {
                        string? sheetName = GetSheetNameForPage(selectedSheets, pageNum, numberOfPages);
                        string orientation = GetDefaultOrientation(sheetName); // Get default based on sheet name

                        // Override with user selection if available
                        if (sheetName != null && sheetOrientations.ContainsKey(sheetName))
                        {
                            orientation = sheetOrientations[sheetName];
                            Console.WriteLine($"🎯 Page {pageNum} ({sheetName}): User selected {orientation}");
                        }
                        else
                        {
                            Console.WriteLine($"📄 Page {pageNum} ({sheetName ?? "Unknown"}): Using default {orientation}");
                        }

                        // Get source page
                        var sourcePage = sourcePdfDoc.GetPage(pageNum);
                        var sourcePageSize = sourcePage.GetPageSize();

                        // Determine target page size and rotation
                        PageSize targetPageSize;
                        float rotation = 0f;

                        if (orientation == "Portrait")
                        {
                            // For AOC and similar sheets, apply 90-degree rotation
                            if (ShouldRotate90Degrees(sheetName))
                            {
                                targetPageSize = PageSize.A4;
                                rotation = 90f; // Rotate 90 degrees clockwise
                                Console.WriteLine($"🔄 Page {pageNum} ({sheetName}): Applying 90° rotation for Portrait");
                            }
                            else
                            {
                                targetPageSize = PageSize.A4;
                                rotation = 0f;
                            }
                        }
                        else // Landscape
                        {
                            targetPageSize = PageSize.A4.Rotate();
                            rotation = 0f;
                        }

                        // Create new page with target size
                        var newPage = newPdfDoc.AddNewPage(targetPageSize);

                        // Copy content from source page
                        var copiedPage = sourcePage.CopyAsFormXObject(newPdfDoc);

                        // Create canvas and apply transformations
                        var canvas = new PdfCanvas(newPage);

                        if (rotation != 0f)
                        {
                            // Apply rotation transformation
                            if (rotation == 90f)
                            {
                                // Rotate 90 degrees clockwise and position correctly
                                canvas.ConcatMatrix(0, 1, -1, 0, targetPageSize.GetWidth(), 0);
                            }
                        }

                        // Add the copied content
                        canvas.AddXObjectAt(copiedPage, 0, 0);
                    }

                    newPdfDoc.Close();
                    sourcePdfDoc.Close();
                }

                // Replace original file
                System.IO.File.Delete(pdfPath);
                System.IO.File.Move(tempOutputPath, pdfPath);

                Console.WriteLine("✅ Page orientations and rotations applied successfully");
                return pdfPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error applying orientations with iText7: {ex.Message}");
                Console.WriteLine($"❌ Stack trace: {ex.StackTrace}");

                // Cleanup temp file
                if (tempOutputPath != null && System.IO.File.Exists(tempOutputPath))
                {
                    try { System.IO.File.Delete(tempOutputPath); } catch { }
                }

                return pdfPath;
            }
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


//using ExcelToPdfConverter.Models;
//using System.Diagnostics;
//using iText.Kernel.Pdf;
//using iText.Kernel.Geom;
//using iText.Layout;
//using iText.Kernel.Pdf.Canvas;

//namespace ExcelToPdfConverter.Services
//{
//    public class LibreOfficeService
//    {
//        private readonly string _libreOfficePath;
//        private readonly string _tempDirectory;
//        private readonly string _outputDirectory;
//        private readonly IWebHostEnvironment _environment;

//        public LibreOfficeService(IWebHostEnvironment environment)
//        {
//            _environment = environment;
//            _libreOfficePath = GetLibreOfficePath();
//            _tempDirectory = System.IO.Path.Combine(_environment.WebRootPath, "App_Data", "Temp");
//            _outputDirectory = System.IO.Path.Combine(_environment.WebRootPath, "App_Data", "Output");
//            Directory.CreateDirectory(_tempDirectory);
//            Directory.CreateDirectory(_outputDirectory);
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

//        public async Task<ConversionResult> ConvertToPdfAsync(string inputFilePath, string outputFileName,
//            List<string>? selectedSheets = null, Dictionary<string, string>? sheetOrientations = null)
//        {
//            var result = new ConversionResult();
//            try
//            {
//                Console.WriteLine($"=== Starting PDF Conversion ===");
//                Console.WriteLine($"Input file: {inputFilePath}");
//                Console.WriteLine($"Output file name: {outputFileName}");
//                Console.WriteLine($"Selected sheets: {(selectedSheets != null ? string.Join(", ", selectedSheets) : "All")}");
//                Console.WriteLine($"Sheet orientations: {(sheetOrientations != null ? string.Join(", ", sheetOrientations) : "Default")}");
//                Console.WriteLine($"File exists: {System.IO.File.Exists(inputFilePath)}");

//                // Use user's home directory for output
//                var outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
//                var outputFilePath = System.IO.Path.Combine(outputDirectory, outputFileName);

//                Console.WriteLine($"Output directory: {outputDirectory}");
//                Console.WriteLine($"Full output path: {outputFilePath}");

//                // Build LibreOffice arguments with sheet selection
//                var arguments = BuildLibreOfficeArguments(inputFilePath, outputDirectory, selectedSheets);
//                Console.WriteLine($"LibreOffice arguments: {arguments}");

//                var processStartInfo = new ProcessStartInfo
//                {
//                    FileName = _libreOfficePath,
//                    Arguments = arguments,
//                    UseShellExecute = false,
//                    CreateNoWindow = true,
//                    RedirectStandardOutput = true,
//                    RedirectStandardError = true,
//                    WindowStyle = ProcessWindowStyle.Hidden,
//                    WorkingDirectory = outputDirectory
//                };

//                Console.WriteLine($"LibreOffice path: {_libreOfficePath}");

//                using (var process = new Process())
//                {
//                    process.StartInfo = processStartInfo;
//                    Console.WriteLine("Starting LibreOffice process...");
//                    process.Start();

//                    // Read outputs
//                    string output = await process.StandardOutput.ReadToEndAsync();
//                    string error = await process.StandardError.ReadToEndAsync();

//                    Console.WriteLine($"LibreOffice Output: {output}");
//                    if (!string.IsNullOrEmpty(error))
//                        Console.WriteLine($"LibreOffice Error: {error}");

//                    // Wait with timeout
//                    bool processExited = process.WaitForExit(120000); // 2 minutes
//                    Console.WriteLine($"Process exited: {processExited}, Exit code: {process.ExitCode}");

//                    if (processExited && process.ExitCode == 0)
//                    {
//                        // Check multiple possible output locations
//                        var inputFileName = System.IO.Path.GetFileNameWithoutExtension(inputFilePath);
//                        var possibleOutputPaths = new[]
//                        {
//                            outputFilePath,
//                            System.IO.Path.Combine(outputDirectory, inputFileName + ".pdf"),
//                            System.IO.Path.Combine(System.IO.Path.GetDirectoryName(inputFilePath) ?? "", inputFileName + ".pdf")
//                        };

//                        string? foundPath = null;
//                        foreach (var path in possibleOutputPaths)
//                        {
//                            if (System.IO.File.Exists(path))
//                            {
//                                foundPath = path;
//                                Console.WriteLine($"✅ PDF found at: {path}");
//                                break;
//                            }
//                        }

//                        if (foundPath != null)
//                        {
//                            // Apply orientations if specified using iText7
//                            if (sheetOrientations != null && sheetOrientations.Any())
//                            {
//                                foundPath = await ApplyPageOrientations(foundPath, sheetOrientations, selectedSheets);
//                            }

//                            result.Success = true;
//                            result.Message = "Conversion successful";
//                            result.PdfFilePath = foundPath;
//                            result.FileName = System.IO.Path.GetFileName(foundPath);
//                        }
//                        else
//                        {
//                            result.Success = false;
//                            result.Message = "Conversion completed but PDF file not found";
//                            Console.WriteLine("❌ PDF file not found in any expected location");
//                        }
//                    }
//                    else
//                    {
//                        result.Success = false;
//                        result.Message = $"Process failed or timed out. Exit code: {process.ExitCode}";
//                        Console.WriteLine($"❌ Process failed: {result.Message}");
//                    }
//                }

//                // Cleanup input file only if conversion successful
//                if (result.Success && System.IO.File.Exists(inputFilePath))
//                {
//                    System.IO.File.Delete(inputFilePath);
//                    Console.WriteLine($"Cleaned up input file: {inputFilePath}");
//                }
//            }
//            catch (Exception ex)
//            {
//                result.Success = false;
//                result.Message = $"Error during conversion: {ex.Message}";
//                Console.WriteLine($"❌ Conversion error: {ex}");
//            }

//            Console.WriteLine($"=== Conversion Result: {result.Success} ===");
//            return result;
//        }

//        private string BuildLibreOfficeArguments(string inputFilePath, string outputDirectory, List<string>? selectedSheets)
//        {
//            var arguments = new List<string>
//            {
//                "--headless",
//                "--norestore",
//                "--nofirststartwizard",
//                "--convert-to pdf",
//                $"--outdir \"{outputDirectory}\""
//            };

//            if (selectedSheets != null && selectedSheets.Any())
//            {
//                Console.WriteLine($"Sheet selection specified: {string.Join(", ", selectedSheets)}");
//            }

//            arguments.Add($"\"{inputFilePath}\"");
//            return string.Join(" ", arguments);
//        }

//        private async Task<string> ApplyPageOrientations(string pdfPath, Dictionary<string, string> sheetOrientations, List<string>? selectedSheets)
//        {
//            string? tempOutputPath = null;

//            try
//            {
//                Console.WriteLine("🔄 Applying page orientations using iText7...");

//                if (!System.IO.File.Exists(pdfPath))
//                {
//                    Console.WriteLine("❌ PDF file not found for orientation application");
//                    return pdfPath;
//                }

//                if (sheetOrientations == null || !sheetOrientations.Any())
//                {
//                    Console.WriteLine("ℹ️ No orientations specified, skipping");
//                    return pdfPath;
//                }

//                tempOutputPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"oriented_{Guid.NewGuid()}.pdf");
//                Console.WriteLine($"📁 Temp output path: {tempOutputPath}");

//                // iText7 approach - create new PDF with correct page sizes
//                using (var reader = new PdfReader(pdfPath))
//                using (var writer = new PdfWriter(tempOutputPath))
//                using (var newPdfDoc = new PdfDocument(writer))
//                using (var sourcePdfDoc = new PdfDocument(reader))
//                {
//                    var numberOfPages = sourcePdfDoc.GetNumberOfPages();
//                    Console.WriteLine($"📄 Processing {numberOfPages} pages");

//                    for (int pageNum = 1; pageNum <= numberOfPages; pageNum++)
//                    {
//                        string? sheetName = GetSheetNameForPage(selectedSheets, pageNum, numberOfPages);
//                        string orientation = "Landscape"; // default

//                        if (sheetName != null && sheetOrientations.ContainsKey(sheetName))
//                        {
//                            orientation = sheetOrientations[sheetName];
//                            Console.WriteLine($"🎯 Page {pageNum} ({sheetName}): Applying {orientation} orientation");
//                        }
//                        else
//                        {
//                            Console.WriteLine($"📄 Page {pageNum} ({sheetName ?? "Unknown"}): Using default {orientation} orientation");
//                        }

//                        // Determine page size based on orientation
//                        PageSize pageSize = orientation == "Portrait" ? PageSize.A4 : PageSize.A4.Rotate();

//                        // Create new page with correct size
//                        var newPage = newPdfDoc.AddNewPage(pageSize);

//                        // Copy content from original page to new page
//                        var sourcePage = sourcePdfDoc.GetPage(pageNum);
//                        var copiedPage = sourcePage.CopyAsFormXObject(newPdfDoc);

//                        // Add content to new page
//                        var canvas = new PdfCanvas(newPage);
//                        canvas.AddXObjectAt(copiedPage, 0, 0);
//                    }

//                    newPdfDoc.Close();
//                    sourcePdfDoc.Close();
//                }

//                // Replace original file
//                System.IO.File.Delete(pdfPath);
//                System.IO.File.Move(tempOutputPath, pdfPath);

//                Console.WriteLine("✅ Page orientations applied successfully using iText7");
//                return pdfPath;
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error applying orientations with iText7: {ex.Message}");
//                Console.WriteLine($"❌ Stack trace: {ex.StackTrace}");

//                // Cleanup temp file
//                if (tempOutputPath != null && System.IO.File.Exists(tempOutputPath))
//                {
//                    try { System.IO.File.Delete(tempOutputPath); } catch { }
//                }

//                return pdfPath;
//            }
//        }

//        private string? GetSheetNameForPage(List<string>? selectedSheets, int pageNum, int totalPages)
//        {
//            try
//            {
//                if (selectedSheets == null || !selectedSheets.Any())
//                    return null;

//                if (pageNum <= selectedSheets.Count)
//                {
//                    return selectedSheets[pageNum - 1];
//                }

//                return null;
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error in GetSheetNameForPage: {ex.Message}");
//                return null;
//            }
//        }

//        public string SaveUploadedFile(IFormFile file)
//        {
//            var fileName = Guid.NewGuid() + System.IO.Path.GetExtension(file.FileName);
//            var filePath = System.IO.Path.Combine(_tempDirectory, fileName);
//            using (var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
//            {
//                file.CopyTo(stream);
//            }
//            Console.WriteLine($"File saved for conversion: {filePath}");
//            return filePath;
//        }

//        public void CleanupOldFiles(int hoursOld = 24)
//        {
//            try
//            {
//                CleanupDirectory(_tempDirectory, hoursOld);
//                CleanupDirectory(_outputDirectory, hoursOld);
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"Cleanup error: {ex.Message}");
//            }
//        }

//        private void CleanupDirectory(string directory, int hoursOld)
//        {
//            if (!Directory.Exists(directory)) return;

//            var cutoffTime = DateTime.Now.AddHours(-hoursOld);
//            var files = Directory.GetFiles(directory);
//            foreach (var file in files)
//            {
//                try
//                {
//                    var fileInfo = new System.IO.FileInfo(file);
//                    if (fileInfo.LastWriteTime < cutoffTime)
//                    {
//                        fileInfo.Delete();
//                        Console.WriteLine($"Cleaned up old file: {file}");
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine($"Error deleting file {file}: {ex.Message}");
//                }
//            }
//        }
//    }
//}



//using ExcelToPdfConverter.Models;
//using System.Diagnostics;
//using iTextSharp.text;
//using iTextSharp.text.pdf;


//namespace ExcelToPdfConverter.Services
//{
//    public class LibreOfficeService
//    {
//        private readonly string _libreOfficePath;
//        private readonly string _tempDirectory;
//        private readonly string _outputDirectory;
//        private readonly IWebHostEnvironment _environment;

//        public LibreOfficeService(IWebHostEnvironment environment)
//        {
//            _environment = environment;
//            _libreOfficePath = GetLibreOfficePath();

//            _tempDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "Temp");
//            _outputDirectory = Path.Combine(_environment.WebRootPath, "App_Data", "Output");

//            Directory.CreateDirectory(_tempDirectory);
//            Directory.CreateDirectory(_outputDirectory);
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
//                if (File.Exists(path))
//                {
//                    Console.WriteLine($"LibreOffice found at: {path}");
//                    return path;
//                }
//            }

//            throw new Exception("LibreOffice not found. Please install LibreOffice from https://www.libreoffice.org/download/download-libreoffice/");
//        }

//        public async Task<ConversionResult> ConvertToPdfAsync(string inputFilePath, string outputFileName,
//            List<string> selectedSheets = null, Dictionary<string, string> sheetOrientations = null)
//        {
//            var result = new ConversionResult();

//            try
//            {
//                Console.WriteLine($"=== Starting PDF Conversion ===");
//                Console.WriteLine($"Input file: {inputFilePath}");
//                Console.WriteLine($"Output file name: {outputFileName}");
//                Console.WriteLine($"Selected sheets: {(selectedSheets != null ? string.Join(", ", selectedSheets) : "All")}");
//                Console.WriteLine($"Sheet orientations: {(sheetOrientations != null ? string.Join(", ", sheetOrientations) : "Default")}");
//                Console.WriteLine($"File exists: {File.Exists(inputFilePath)}");

//                // Use user's home directory for output
//                var outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
//                var outputFilePath = Path.Combine(outputDirectory, outputFileName);

//                Console.WriteLine($"Output directory: {outputDirectory}");
//                Console.WriteLine($"Full output path: {outputFilePath}");

//                // Build LibreOffice arguments with sheet selection and orientation
//                var arguments = BuildLibreOfficeArguments(inputFilePath, outputDirectory, selectedSheets, sheetOrientations);

//                Console.WriteLine($"LibreOffice arguments: {arguments}");

//                var processStartInfo = new ProcessStartInfo
//                {
//                    FileName = _libreOfficePath,
//                    Arguments = arguments,
//                    UseShellExecute = false,
//                    CreateNoWindow = true,
//                    RedirectStandardOutput = true,
//                    RedirectStandardError = true,
//                    WindowStyle = ProcessWindowStyle.Hidden,
//                    WorkingDirectory = outputDirectory
//                };

//                Console.WriteLine($"LibreOffice path: {_libreOfficePath}");

//                using (var process = new Process())
//                {
//                    process.StartInfo = processStartInfo;

//                    Console.WriteLine("Starting LibreOffice process...");
//                    process.Start();

//                    // Read outputs
//                    string output = await process.StandardOutput.ReadToEndAsync();
//                    string error = await process.StandardError.ReadToEndAsync();

//                    Console.WriteLine($"LibreOffice Output: {output}");
//                    if (!string.IsNullOrEmpty(error))
//                        Console.WriteLine($"LibreOffice Error: {error}");

//                    // Wait with timeout
//                    bool processExited = process.WaitForExit(120000); // 2 minutes
//                    Console.WriteLine($"Process exited: {processExited}, Exit code: {process.ExitCode}");

//                    if (processExited && process.ExitCode == 0)
//                    {
//                        // Check multiple possible output locations
//                        var inputFileName = Path.GetFileNameWithoutExtension(inputFilePath);
//                        var possibleOutputPaths = new[]
//                        {
//                            outputFilePath,
//                            Path.Combine(outputDirectory, inputFileName + ".pdf"),
//                            Path.Combine(Path.GetDirectoryName(inputFilePath), inputFileName + ".pdf")
//                        };

//                        string foundPath = null;
//                        foreach (var path in possibleOutputPaths)
//                        {
//                            if (File.Exists(path))
//                            {
//                                foundPath = path;
//                                Console.WriteLine($"✅ PDF found at: {path}");
//                                break;
//                            }
//                        }

//                        if (foundPath != null)
//                        {
//                            result.Success = true;
//                            result.Message = "Conversion successful";
//                            result.PdfFilePath = foundPath;
//                            result.FileName = Path.GetFileName(foundPath);
//                        }
//                        else
//                        {
//                            result.Success = false;
//                            result.Message = "Conversion completed but PDF file not found";
//                            Console.WriteLine("❌ PDF file not found in any expected location");

//                            // Debug: List all PDF files
//                            var allPdfFiles = Directory.GetFiles(outputDirectory, "*.pdf");
//                            Console.WriteLine($"PDF files in {outputDirectory}: {string.Join(", ", allPdfFiles)}");
//                        }
//                    }
//                    else
//                    {
//                        result.Success = false;
//                        result.Message = $"Process failed or timed out. Exit code: {process.ExitCode}";
//                        Console.WriteLine($"❌ Process failed: {result.Message}");
//                    }
//                }

//                // Cleanup input file only if conversion successful
//                if (result.Success && File.Exists(inputFilePath))
//                {
//                    File.Delete(inputFilePath);
//                    Console.WriteLine($"Cleaned up input file: {inputFilePath}");
//                }
//            }
//            catch (Exception ex)
//            {
//                result.Success = false;
//                result.Message = $"Error during conversion: {ex.Message}";
//                Console.WriteLine($"❌ Conversion error: {ex}");
//            }

//            Console.WriteLine($"=== Conversion Result: {result.Success} ===");
//            return result;
//        }

//        private string BuildLibreOfficeArguments(string inputFilePath, string outputDirectory,
//            List<string> selectedSheets, Dictionary<string, string> sheetOrientations)
//        {
//            var arguments = new List<string>
//            {
//                "--headless",
//                "--norestore",
//                "--nofirststartwizard",
//                "--convert-to pdf",
//                $"--outdir \"{outputDirectory}\""
//            };

//            // Add sheet selection if specified
//            if (selectedSheets != null && selectedSheets.Any())
//            {
//                arguments.Add($"\"{inputFilePath}\"");
//            }
//            else
//            {
//                arguments.Add($"\"{inputFilePath}\"");
//            }

//            return string.Join(" ", arguments);
//        }

//        // NEW: PDF Merge Method using iTextSharp
//        public async Task<ConversionResult> MergePdfFilesAsync(string newPdfPath, string outputFileName)
//        {
//            var result = new ConversionResult();

//            try
//            {
//                var pdfDirectory = @"D:\CIPL\SinghAndSons\pdf";
//                var mergedPdfPath = Path.Combine(Path.GetTempPath(), $"merged_{Guid.NewGuid()}.pdf");
//                var mergedFileName = $"merged_{Path.GetFileNameWithoutExtension(outputFileName)}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

//                Console.WriteLine($"Starting PDF merge process...");
//                Console.WriteLine($"New PDF: {newPdfPath}");
//                Console.WriteLine($"PDF Directory: {pdfDirectory}");
//                Console.WriteLine($"Merged PDF will be saved at: {mergedPdfPath}");

//                using (var mergedDocument = new Document())
//                using (var mergedPdfWriter = new PdfCopy(mergedDocument, new FileStream(mergedPdfPath, FileMode.Create)))
//                {
//                    mergedDocument.Open();

//                    // Step 1: Add existing PDFs from directory
//                    if (Directory.Exists(pdfDirectory))
//                    {
//                        var existingPdfFiles = Directory.GetFiles(pdfDirectory, "*.pdf")
//                            .OrderBy(f => f)
//                            .ToList();

//                        Console.WriteLine($"Found {existingPdfFiles.Count} existing PDF files");

//                        foreach (var existingPdf in existingPdfFiles)
//                        {
//                            try
//                            {
//                                using (var existingReader = new PdfReader(existingPdf))
//                                {
//                                    for (int page = 1; page <= existingReader.NumberOfPages; page++)
//                                    {
//                                        var importedPage = mergedPdfWriter.GetImportedPage(existingReader, page);
//                                        mergedPdfWriter.AddPage(importedPage);
//                                    }
//                                    Console.WriteLine($"✅ Added existing PDF: {Path.GetFileName(existingPdf)} ({existingReader.NumberOfPages} pages)");
//                                }
//                            }
//                            catch (Exception ex)
//                            {
//                                Console.WriteLine($"❌ Error adding existing PDF {existingPdf}: {ex.Message}");
//                            }
//                        }
//                    }
//                    else
//                    {
//                        Console.WriteLine($"PDF directory not found: {pdfDirectory}");
//                    }

//                    // Step 2: Add the newly converted PDF
//                    if (File.Exists(newPdfPath))
//                    {
//                        try
//                        {
//                            using (var newPdfReader = new PdfReader(newPdfPath))
//                            {
//                                for (int page = 1; page <= newPdfReader.NumberOfPages; page++)
//                                {
//                                    var importedPage = mergedPdfWriter.GetImportedPage(newPdfReader, page);
//                                    mergedPdfWriter.AddPage(importedPage);
//                                }
//                                Console.WriteLine($"✅ Added new converted PDF: {Path.GetFileName(newPdfPath)} ({newPdfReader.NumberOfPages} pages)");
//                            }
//                        }
//                        catch (Exception ex)
//                        {
//                            Console.WriteLine($"❌ Error adding new PDF {newPdfPath}: {ex.Message}");
//                            throw;
//                        }
//                    }
//                    else
//                    {
//                        Console.WriteLine($"❌ New PDF file not found: {newPdfPath}");
//                        throw new FileNotFoundException($"New PDF file not found: {newPdfPath}");
//                    }

//                    mergedDocument.Close();
//                }

//                // Verify the merged file was created
//                if (File.Exists(mergedPdfPath))
//                {
//                    var fileInfo = new FileInfo(mergedPdfPath);
//                    Console.WriteLine($"✅ Merged PDF created successfully: {fileInfo.Length} bytes");

//                    result.Success = true;
//                    result.PdfFilePath = mergedPdfPath;
//                    result.FileName = mergedFileName;
//                    result.Message = "PDF files merged successfully";

//                    // Cleanup the original converted PDF
//                    try
//                    {
//                        if (File.Exists(newPdfPath))
//                        {
//                            File.Delete(newPdfPath);
//                            Console.WriteLine($"Cleaned up original PDF: {newPdfPath}");
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        Console.WriteLine($"Warning: Could not cleanup original PDF: {ex.Message}");
//                    }
//                }
//                else
//                {
//                    throw new Exception("Merged PDF file was not created");
//                }
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"❌ Error merging PDFs: {ex}");
//                result.Success = false;
//                result.Message = $"PDF merge failed: {ex.Message}";

//                // Fallback: return the original converted PDF
//                if (File.Exists(newPdfPath))
//                {
//                    result.PdfFilePath = newPdfPath;
//                    result.FileName = outputFileName;
//                }
//            }

//            return result;
//        }

//        public string SaveUploadedFile(IFormFile file)
//        {
//            var fileName = Guid.NewGuid() + Path.GetExtension(file.FileName);
//            var filePath = Path.Combine(_tempDirectory, fileName);

//            using (var stream = new FileStream(filePath, FileMode.Create))
//            {
//                file.CopyTo(stream);
//            }

//            Console.WriteLine($"File saved for conversion: {filePath}");
//            return filePath;
//        }

//        public void CleanupOldFiles(int hoursOld = 24)
//        {
//            try
//            {
//                CleanupDirectory(_tempDirectory, hoursOld);
//                CleanupDirectory(_outputDirectory, hoursOld);
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"Cleanup error: {ex.Message}");
//            }
//        }

//        private void CleanupDirectory(string directory, int hoursOld)
//        {
//            if (!Directory.Exists(directory)) return;

//            var cutoffTime = DateTime.Now.AddHours(-hoursOld);
//            var files = Directory.GetFiles(directory);

//            foreach (var file in files)
//            {
//                try
//                {
//                    var fileInfo = new FileInfo(file);
//                    if (fileInfo.LastWriteTime < cutoffTime)
//                    {
//                        fileInfo.Delete();
//                        Console.WriteLine($"Cleaned up old file: {file}");
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine($"Error deleting file {file}: {ex.Message}");
//                }
//            }
//        }
//    }
//}
