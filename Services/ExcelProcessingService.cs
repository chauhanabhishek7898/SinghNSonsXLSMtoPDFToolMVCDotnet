using OfficeOpenXml;
using ExcelToPdfConverter.Models;

namespace ExcelToPdfConverter.Services
{
    public class ExcelProcessingService
    {
        public ExcelProcessingService()
        {
            // EPPlus license already set in Program.cs
        }

        public async Task<string> ProcessExcelFileAsync(string inputFilePath, List<string> sheetsToKeep, List<string> sheetOrder)
        {
            try
            {
                Console.WriteLine($"=== Starting Excel File Processing ===");
                Console.WriteLine($"Input file: {inputFilePath}");
                Console.WriteLine($"Sheets to keep: {string.Join(", ", sheetsToKeep)}");
                Console.WriteLine($"Sheet order: {string.Join(" → ", sheetOrder)}");

                // Output file path
                var outputFilePath = Path.Combine(Path.GetDirectoryName(inputFilePath)!,
                    $"processed_{Guid.NewGuid()}_{Path.GetFileName(inputFilePath)}");

                using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
                {
                    var workbook = package.Workbook;

                    Console.WriteLine($"Original sheets: {workbook.Worksheets.Count}");

                    // Step 1: Delete unchecked sheets
                    var sheetsToDelete = workbook.Worksheets
                        .Where(ws => !sheetsToKeep.Contains(ws.Name))
                        .Select(ws => ws.Name)
                        .ToList();

                    foreach (var sheetName in sheetsToDelete)
                    {
                        workbook.Worksheets.Delete(sheetName);
                        Console.WriteLine($"🗑️ Deleted sheet: {sheetName}");
                    }

                    // Step 2: Reorder sheets according to drag & drop
                    if (sheetOrder.Any() && sheetOrder.Count > 1)
                    {
                        ReorderSheetsSimple(workbook, sheetOrder);
                        Console.WriteLine($"🔄 Sheets reordered as: {string.Join(" → ", sheetOrder)}");
                    }

                    // Save the processed file
                    package.SaveAs(new FileInfo(outputFilePath));
                    Console.WriteLine($"✅ Processed Excel file saved: {outputFilePath}");
                    Console.WriteLine($"Final sheets: {workbook.Worksheets.Count}");
                }

                return outputFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error processing Excel file: {ex}");
                throw;
            }
        }

        // Simple and reliable method for EPPlus 8.x
        private void ReorderSheetsSimple(ExcelWorkbook workbook, List<string> newOrder)
        {
            try
            {
                Console.WriteLine("🔄 Reordering sheets using simple method...");

                // Collect all worksheets
                var worksheets = workbook.Worksheets.ToList();
                var worksheetDict = worksheets.ToDictionary(w => w.Name, w => w);

                Console.WriteLine($"📋 Worksheets found: {string.Join(", ", worksheetDict.Keys)}");

                // Pehle sab worksheets delete karein
                // Backwards loop karein kyunki delete karne se indexes change hote hain
                for (int i = worksheets.Count - 1; i >= 0; i--)
                {
                    var worksheet = worksheets[i];
                    workbook.Worksheets.Delete(worksheet.Name);
                    Console.WriteLine($"🗑️ Deleted: {worksheet.Name}");
                }

                // Naye order mein worksheets add karein
                foreach (var sheetName in newOrder)
                {
                    if (worksheetDict.ContainsKey(sheetName))
                    {
                        // EPPlus 8.x mein Copy method use karein
                        var newWorksheet = workbook.Worksheets.Add(sheetName, worksheetDict[sheetName]);
                        Console.WriteLine($"✅ Added sheet in order: {sheetName}");
                    }
                    else
                    {
                        Console.WriteLine($"⚠️ Sheet not found: {sheetName}");
                    }
                }

                // Remaining sheets add karein (jo newOrder mein nahi hain)
                foreach (var worksheet in worksheetDict)
                {
                    if (!newOrder.Contains(worksheet.Key))
                    {
                        var newWorksheet = workbook.Worksheets.Add(worksheet.Key, worksheet.Value);
                        Console.WriteLine($"✅ Added remaining sheet: {worksheet.Key}");
                    }
                }

                Console.WriteLine($"✅ Sheet reordering completed. Total sheets: {workbook.Worksheets.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in ReorderSheetsSimple: {ex.Message}");
                Console.WriteLine($"❌ Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        public void CleanupProcessedFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine($"🧹 Cleaned up processed file: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Error cleaning up file {filePath}: {ex.Message}");
            }
        }

        // ✅ BETTER APPROACH: Create new file with desired sheets and order
        public async Task<string> CreateProcessedExcelFileAsync(string inputFilePath, List<string> sheetsToKeep, List<string> sheetOrder)
        {
            try
            {
                Console.WriteLine($"=== Creating Processed Excel File ===");
                Console.WriteLine($"Input: {inputFilePath}");
                Console.WriteLine($"Sheets to keep: {string.Join(", ", sheetsToKeep)}");
                Console.WriteLine($"Order: {string.Join(" → ", sheetOrder)}");

                var outputFilePath = Path.Combine(
                    Path.GetDirectoryName(inputFilePath)!,
                    $"processed_{DateTime.Now:yyyyMMdd_HHmmss}_{Path.GetFileName(inputFilePath)}"
                );

                using (var sourcePackage = new ExcelPackage(new FileInfo(inputFilePath)))
                using (var targetPackage = new ExcelPackage())
                {
                    var sourceWorkbook = sourcePackage.Workbook;
                    var targetWorkbook = targetPackage.Workbook;

                    Console.WriteLine($"Source sheets: {sourceWorkbook.Worksheets.Count}");

                    // Step 1: Add sheets in desired order
                    foreach (var sheetName in sheetOrder)
                    {
                        var sourceWorksheet = sourceWorkbook.Worksheets[sheetName];
                        if (sourceWorksheet != null && sheetsToKeep.Contains(sheetName))
                        {
                            targetWorkbook.Worksheets.Add(sheetName, sourceWorksheet);
                            Console.WriteLine($"✅ Added sheet: {sheetName}");
                        }
                    }

                    // Step 2: Add any remaining sheets that are in sheetsToKeep but not in order
                    foreach (var worksheet in sourceWorkbook.Worksheets)
                    {
                        if (sheetsToKeep.Contains(worksheet.Name) && !sheetOrder.Contains(worksheet.Name))
                        {
                            targetWorkbook.Worksheets.Add(worksheet.Name, worksheet);
                            Console.WriteLine($"✅ Added remaining sheet: {worksheet.Name}");
                        }
                    }

                    // Save the new file
                    targetPackage.SaveAs(new FileInfo(outputFilePath));
                    Console.WriteLine($"✅ Processed file created: {outputFilePath}");
                    Console.WriteLine($"Final sheets: {targetWorkbook.Worksheets.Count}");
                }

                return outputFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating processed file: {ex}");
                throw;
            }
        }
    }
}
