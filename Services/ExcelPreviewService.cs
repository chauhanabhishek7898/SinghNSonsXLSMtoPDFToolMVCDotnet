using ExcelToPdfConverter.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Drawing.Imaging;

namespace ExcelToPdfConverter.Services
{
    public class ExcelPreviewService
    {
        public ExcelPreviewService()
        {
            // License already set in Program.cs
        }

        // ✅ NEW: Sheet orientation analysis class
        public class SheetOrientationInfo
        {
            public string SheetName { get; set; } = string.Empty;
            public string SuggestedOrientation { get; set; } = "Portrait";
            public double WidthToHeightRatio { get; set; }
            public int TotalColumns { get; set; }
            public int TotalRows { get; set; }
            public bool HasWideContent { get; set; }
            public double AverageColumnWidth { get; set; }
        }




        public PreviewModel GeneratePreview(IFormFile excelFile, string sessionId)
        {
            var previewModel = new PreviewModel
            {
                OriginalFileName = Path.GetFileNameWithoutExtension(excelFile.FileName) ?? "Unknown",
                SessionId = sessionId,
                Worksheets = new List<WorksheetPreview>(),
                FileSelections = new List<FileSelection>(),
                AllNameErrors = new List<NameError>(),
                AllInvoiceDates = new List<InvoiceDate>(),
                 // ✅ NEW: Orientation analysis storage
                SuggestedOrientations = new Dictionary<string, string>(),
                SheetOrientationAnalysis = new Dictionary<string, SheetOrientationInfo>()
            };

            try
            {
                Console.WriteLine($"Generating preview for: {excelFile.FileName}");

                using (var stream = excelFile.OpenReadStream())
                {
                    using (var package = new ExcelPackage(stream))
                    {
                        Console.WriteLine($"Workbook has {package.Workbook.Worksheets.Count} worksheets");

                        // ✅ STEP 1: FIRST Analyze all sheets for automatic orientation detection
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension != null)
                            {
                                //var orientationInfo = AnalyzeSheetOrientation(worksheet);
                                //previewModel.SheetOrientationAnalysis[worksheet.Name] = orientationInfo;
                                //previewModel.SuggestedOrientations[worksheet.Name] = orientationInfo.SuggestedOrientation;

                                //Console.WriteLine($"📊 Orientation Analysis - {worksheet.Name}: {orientationInfo.SuggestedOrientation} " +
                                //                $"(Ratio: {orientationInfo.WidthToHeightRatio:F2}, Columns: {orientationInfo.TotalColumns})");
                            }
                        }

                        // FIRST: Find all invoice dates using the robust method
                        var allInvoiceDates = new List<InvoiceDate>();
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null) continue;
                            var invoiceDates = FindInvoiceDatesInWorksheet(worksheet);
                            allInvoiceDates.AddRange(invoiceDates);
                            Console.WriteLine($"Found {invoiceDates.Count} invoice dates in {worksheet.Name}");
                        }

                        // THEN: Create worksheet previews
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null)
                            {
                                Console.WriteLine($"Worksheet {worksheet.Name} has no data, skipping.");
                                continue;
                            }

                            Console.WriteLine($"Processing worksheet: {worksheet.Name}");
                            var worksheetPreview = CreateWorksheetPreview(worksheet);
                            previewModel.Worksheets.Add(worksheetPreview);

                            // Add to global collections
                            previewModel.AllNameErrors.AddRange(worksheetPreview.NameErrors);

                            // Get invoice dates for this specific sheet from the pre-collected list
                            var sheetInvoiceDates = allInvoiceDates
                                .Where(id => id.SheetName == worksheet.Name)
                                .ToList();

                            // ✅ USE AUTOMATIC DETECTED ORIENTATION for file selection
                            string suggestedOrientation = previewModel.SuggestedOrientations.ContainsKey(worksheet.Name)
                                ? previewModel.SuggestedOrientations[worksheet.Name]
                                : "Portrait";

                            // Create file selection entry - use the accurate invoice dates count
                            var fileSelection = new FileSelection
                            {
                                FileName = previewModel.OriginalFileName,
                                SheetName = worksheet.Name,
                                SortOrder = worksheet.Index,
                                HasNameErrors = worksheetPreview.NameErrors.Count > 0,
                                HasInvoiceDates = sheetInvoiceDates.Count > 0,
                                NameErrors = worksheetPreview.NameErrors,
                                InvoiceDates = sheetInvoiceDates,
                                //Orientation = suggestedOrientation
                            };
                            previewModel.FileSelections.Add(fileSelection);

                            Console.WriteLine($"FileSelection for {worksheet.Name}: Orientation = {suggestedOrientation}");

                            Console.WriteLine($"FileSelection for {worksheet.Name}: HasInvoiceDates = {fileSelection.HasInvoiceDates}, Count = {fileSelection.InvoiceDates.Count}");
                        }

                        // Set the global AllInvoiceDates collection
                        previewModel.AllInvoiceDates = allInvoiceDates;
                    }
                }

                Console.WriteLine($"Preview generated for {previewModel.Worksheets.Count} worksheets");
                Console.WriteLine($"Found {previewModel.AllNameErrors.Count} name errors");
                Console.WriteLine($"Found {previewModel.AllInvoiceDates.Count} invoice dates");

                // Debug: Check FileSelections
                Console.WriteLine("=== FILE SELECTIONS DEBUG ===");
                foreach (var fs in previewModel.FileSelections)
                {
                    Console.WriteLine($"{fs.SheetName}: HasInvoiceDates = {fs.HasInvoiceDates}, InvoiceDates.Count = {fs.InvoiceDates.Count}");
                }
                Console.WriteLine("=============================");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating preview: {ex}");
                throw new Exception($"Error reading Excel file: {ex.Message}", ex);
            }

            return previewModel;
        }

        // ✅ NEW: Automatic orientation detection based on content
        private SheetOrientationInfo AnalyzeSheetOrientation(ExcelWorksheet worksheet)
        {
            var info = new SheetOrientationInfo
            {
                SheetName = worksheet.Name
            };

            if (worksheet.Dimension == null)
                return info;

            // Calculate content dimensions
            int usedColumns = worksheet.Dimension.End.Column;
            int usedRows = worksheet.Dimension.End.Row;

            info.TotalColumns = usedColumns;
            info.TotalRows = usedRows;

            // Analyze column widths to determine if content is wide
            double totalWidth = 0;
            int wideColumns = 0;
            int analyzedColumns = 0;

            for (int col = 1; col <= Math.Min(usedColumns, 50); col++) // First 50 columns check
            {
                var width = worksheet.Column(col).Width;
                if (width > 0) // Only consider columns with content
                {
                    totalWidth += width;
                    analyzedColumns++;

                    if (width > 20) // Wide column threshold (adjust as needed)
                        wideColumns++;
                }
            }

            if (analyzedColumns > 0)
            {
                info.AverageColumnWidth = totalWidth / analyzedColumns;

                // Calculate approximate width-to-height ratio
                // Assuming average row height of 15 points
                double estimatedWidth = usedColumns * info.AverageColumnWidth;
                double estimatedHeight = usedRows * 15;

                info.WidthToHeightRatio = estimatedWidth / (estimatedHeight > 0 ? estimatedHeight : 1);
                info.HasWideContent = wideColumns > (analyzedColumns * 0.3) || info.WidthToHeightRatio > 1.3;

                // Determine suggested orientation based on analysis
                if (info.HasWideContent || info.WidthToHeightRatio > 1.5)
                {
                    info.SuggestedOrientation = "Landscape";
                }
                else if (info.WidthToHeightRatio < 0.8)
                {
                    info.SuggestedOrientation = "Portrait";
                }
                else
                {
                    // For moderate ratios, use default based on column count
                    info.SuggestedOrientation = usedColumns > 8 ? "Landscape" : "Portrait";
                }
            }
            else
            {
                // Default to Portrait if no columns analyzed
                info.SuggestedOrientation = "Portrait";
                info.WidthToHeightRatio = 1.0;
            }

            return info;
        }

        public ValidationResult QuickValidate(IFormFile excelFile)
        {
            var validationResult = new ValidationResult
            {
                FileName = excelFile.FileName,
                NameErrors = new List<NameError>(),
                InvoiceDates = new List<InvoiceDate>()
            };

            try
            {
                using (var stream = excelFile.OpenReadStream())
                {
                    using (var package = new ExcelPackage(stream))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null) continue;

                            // Find name errors
                            var nameErrors = FindNameErrorsInWorksheet(worksheet);
                            validationResult.NameErrors.AddRange(nameErrors);

                            // Find invoice dates
                            var invoiceDates = FindInvoiceDatesInWorksheet(worksheet);
                            validationResult.InvoiceDates.AddRange(invoiceDates);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Quick validation error: {ex.Message}");
            }

            return validationResult;
        }

        private List<NameError> FindNameErrorsInWorksheet(ExcelWorksheet worksheet)
        {
            var nameErrors = new List<NameError>();

            if (worksheet.Dimension == null) return nameErrors;

            int maxRows = Math.Min(200, worksheet.Dimension.End.Row);
            int maxCols = Math.Min(50, worksheet.Dimension.End.Column);

            for (int row = 1; row <= maxRows; row++)
            {
                for (int col = 1; col <= maxCols; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var cellValue = GetCellValue(cell);

                    if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("#NAME?"))
                    {
                        nameErrors.Add(new NameError
                        {
                            SheetName = worksheet.Name,
                            Row = row,
                            Column = col,
                            ColumnName = GetColumnName(col)
                        });
                    }
                }
            }

            return nameErrors;
        }

        private List<InvoiceDate> FindInvoiceDatesInWorksheet(ExcelWorksheet worksheet)
        {
            var invoiceDates = new List<InvoiceDate>();

            if (worksheet.Dimension == null)
            {
                Console.WriteLine($"Worksheet {worksheet.Name} has no data, skipping.");
                return invoiceDates;
            }

            int maxRows = worksheet.Dimension.End.Row;
            int maxCols = worksheet.Dimension.End.Column;

            Console.WriteLine($"🔍 Searching for invoice dates in {worksheet.Name} - Rows: {maxRows}, Columns: {maxCols}");

            for (int row = 1; row <= maxRows; row++)
            {
                for (int col = 1; col <= maxCols; col++)
                {
                    var cell = worksheet.Cells[row, col];

                    // Skip merged cells that are not the top-left cell
                    if (cell.Merge && (cell.Start.Row != row || cell.Start.Column != col))
                        continue;

                    var cellValue = GetCellValue(cell);

                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // Check for invoice date text
                    if (IsInvoiceDateText(cellValue))
                    {
                        Console.WriteLine($"📍 Found 'INVOICE DATE' text at {GetColumnName(col)}{row}: '{cellValue}'");
                        Console.WriteLine($"   Merged: {cell.Merge}, Merge area: {cell.Start.Address}:{cell.End.Address}");

                        // Look for date in multiple possible locations
                        var dateValue = FindDateValueNearInvoiceDate(worksheet, row, col, maxRows, maxCols);

                        if (!string.IsNullOrEmpty(dateValue))
                        {
                            invoiceDates.Add(new InvoiceDate
                            {
                                SheetName = worksheet.Name,
                                Row = row,
                                Column = col,
                                ColumnName = GetColumnName(col),
                                InvoiceDateText = cellValue.Trim(),
                                DateValue = dateValue.Trim()
                            });

                            Console.WriteLine($"✅ CONFIRMED invoice date: '{cellValue.Trim()}' → '{dateValue.Trim()}'");
                        }
                        else
                        {
                            Console.WriteLine($"❌ No date found near 'INVOICE DATE' at {GetColumnName(col)}{row}");
                            // Check surrounding cells for debugging
                            CheckSurroundingCells(worksheet, row, col, maxRows, maxCols);
                        }
                    }
                }
            }

            Console.WriteLine($"📊 Total invoice dates found in {worksheet.Name}: {invoiceDates.Count}");
            return invoiceDates;
        }

        private string FindDateValueNearInvoiceDate(ExcelWorksheet worksheet, int row, int col, int maxRows, int maxCols)
        {
            var currentCell = worksheet.Cells[row, col];

            // CASE 1: If current cell is merged, find the next cell after merged area
            if (currentCell.Merge)
            {
                int nextColAfterMerge = currentCell.End.Column + 1;
                Console.WriteLine($"   Current cell merged, checking after merge at column {GetColumnName(nextColAfterMerge)}");

                // Try multiple columns after merged area (for cases like G14-H14 merged, then I,J,K,L merged with date)
                for (int offset = 1; offset <= 4; offset++)
                {
                    int checkCol = nextColAfterMerge + offset - 1;
                    if (checkCol <= maxCols)
                    {
                        var checkCell = worksheet.Cells[row, checkCol];
                        var checkValue = GetCellValue(checkCell);

                        if (!string.IsNullOrEmpty(checkValue) && IsPotentialDate(checkValue))
                        {
                            Console.WriteLine($"   ✅ Found date after merged area at {GetColumnName(checkCol)}: '{checkValue}'");
                            return checkValue;
                        }
                        else
                        {
                            Console.WriteLine($"   Checking {GetColumnName(checkCol)}: '{checkValue}'");
                        }
                    }
                }
            }

            // CASE 2: Try immediate right cell (for non-merged or simple cases)
            int rightCol = col + 1;
            if (rightCol <= maxCols)
            {
                var rightCell = worksheet.Cells[row, rightCol];
                var rightValue = GetCellValue(rightCell);

                if (!string.IsNullOrEmpty(rightValue) && IsPotentialDate(rightValue))
                {
                    Console.WriteLine($"   ✅ Found date in immediate right cell: '{rightValue}'");
                    return rightValue;
                }
            }

            // CASE 3: Try skipping one column (for AOC sheet scenario: A3-C3 merged, skip D3, check E3)
            int skipOneCol = col + 2;
            if (skipOneCol <= maxCols)
            {
                var skipCell = worksheet.Cells[row, skipOneCol];
                var skipValue = GetCellValue(skipCell);

                if (!string.IsNullOrEmpty(skipValue) && IsPotentialDate(skipValue))
                {
                    Console.WriteLine($"   ✅ Found date by skipping one column: '{skipValue}'");
                    return skipValue;
                }
            }

            // CASE 4: Try cell below (for some edge cases)
            if (row + 1 <= maxRows)
            {
                var belowCell = worksheet.Cells[row + 1, col];
                var belowValue = GetCellValue(belowCell);

                if (!string.IsNullOrEmpty(belowValue) && IsPotentialDate(belowValue))
                {
                    Console.WriteLine($"   ✅ Found date in cell below: '{belowValue}'");
                    return belowValue;
                }
            }

            return string.Empty;
        }

        private bool IsPotentialDate(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            var trimmedValue = value.Trim();

            // Check for common date patterns
            if (trimmedValue.Contains(",") && trimmedValue.Contains(" ")) // "Tuesday, September 2, 2025"
                return true;

            if (System.Text.RegularExpressions.Regex.IsMatch(trimmedValue, @"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}")) // DD/MM/YYYY patterns
                return true;

            if (System.Text.RegularExpressions.Regex.IsMatch(trimmedValue, @"\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                return true;

            // If it's not empty and not obviously a label, consider it a date
            var lowerValue = trimmedValue.ToLower();
            if (!lowerValue.Contains("invoice") &&
                !lowerValue.Contains("date") &&
                !lowerValue.Contains("no") &&
                !lowerValue.Contains("number") &&
                trimmedValue.Length > 3)
                return true;

            return false;
        }

        private void CheckSurroundingCells(ExcelWorksheet worksheet, int row, int col, int maxRows, int maxCols)
        {
            Console.WriteLine("   🔎 Checking surrounding cells:");

            // Check right cells (up to 5 columns)
            for (int i = 1; i <= 5; i++)
            {
                int rightCol = col + i;
                if (rightCol <= maxCols)
                {
                    var rightCell = worksheet.Cells[row, rightCol];
                    var rightValue = GetCellValue(rightCell);
                    Console.WriteLine($"     → Right {i} ({GetColumnName(rightCol)}{row}): '{rightValue}' {(IsPotentialDate(rightValue) ? "📅 POTENTIAL DATE" : "")}");
                }
            }

            // Check below cells
            for (int i = 1; i <= 2; i++)
            {
                int belowRow = row + i;
                if (belowRow <= maxRows)
                {
                    var belowCell = worksheet.Cells[belowRow, col];
                    var belowValue = GetCellValue(belowCell);
                    Console.WriteLine($"     ↓ Below {i} ({GetColumnName(col)}{belowRow}): '{belowValue}' {(IsPotentialDate(belowValue) ? "📅 POTENTIAL DATE" : "")}");
                }
            }
        }

        private bool IsInvoiceDateText(string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue)) return false;

            var normalizedValue = cellValue.Trim().ToUpper();

            // Exact matches for invoice date
            var exactPatterns = new[]
            {
                "INVOICE DATE",
                "INVOICEDATE",
                "INVOICE_DATE",
                "INVOICE  DATE",  // Double space
                "INVOICE DATE\r\n",
                "INVOICE DATE\n",
                "INVOICE DATE",
                "INVOICE  DATE"
            };

            // Contains check
            var containsPatterns = new[]
            {
                "INVOICE DATE",
                "INVOICEDATE"
            };

            return exactPatterns.Any(pattern => normalizedValue == pattern) ||
                   containsPatterns.Any(pattern => normalizedValue.Contains(pattern));
        }

        private WorksheetPreview CreateWorksheetPreview(ExcelWorksheet worksheet)
        {
            var worksheetPreview = new WorksheetPreview
            {
                Name = worksheet.Name,
                Index = worksheet.Index,
                TotalRows = Math.Min(100, worksheet.Dimension.End.Row),
                TotalColumns = Math.Min(30, worksheet.Dimension.End.Column),
                Cells = new List<List<CellPreview>>(),
                Images = new List<ImagePreview>(),
                NameErrors = new List<NameError>(),
                InvoiceDates = new List<InvoiceDate>()
            };

            // Read cells with formatting
            ReadCells(worksheet, worksheetPreview);

            // Find name errors
            worksheetPreview.NameErrors = FindNameErrorsInPreview(worksheetPreview);

            // Find invoice dates - use the same robust method but with preview limits
            worksheetPreview.InvoiceDates = FindInvoiceDatesInWorksheetForPreview(worksheet);

            // Read images
            ReadImages(worksheet, worksheetPreview);

            Console.WriteLine($"Worksheet {worksheet.Name} preview created with {worksheetPreview.InvoiceDates.Count} invoice dates");

            return worksheetPreview;
        }

        // Separate method for preview to handle limited range
        private List<InvoiceDate> FindInvoiceDatesInWorksheetForPreview(ExcelWorksheet worksheet)
        {
            var invoiceDates = new List<InvoiceDate>();

            if (worksheet.Dimension == null)
            {
                Console.WriteLine($"Worksheet {worksheet.Name} has no data for preview, skipping.");
                return invoiceDates;
            }

            int maxRows = Math.Min(100, worksheet.Dimension.End.Row);
            int maxCols = Math.Min(30, worksheet.Dimension.End.Column);

            Console.WriteLine($"🔍 [PREVIEW] Searching for invoice dates in {worksheet.Name} - Rows: 1-{maxRows}, Columns: 1-{maxCols}");

            for (int row = 1; row <= maxRows; row++)
            {
                for (int col = 1; col <= maxCols; col++)
                {
                    var cell = worksheet.Cells[row, col];

                    // Skip merged cells that are not the top-left cell
                    if (cell.Merge && (cell.Start.Row != row || cell.Start.Column != col))
                        continue;

                    var cellValue = GetCellValue(cell);

                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // Check for invoice date text
                    if (IsInvoiceDateText(cellValue))
                    {
                        // Look for date in multiple possible locations with preview limits
                        var dateValue = FindDateValueNearInvoiceDateForPreview(worksheet, row, col, maxRows, maxCols);

                        if (!string.IsNullOrEmpty(dateValue))
                        {
                            invoiceDates.Add(new InvoiceDate
                            {
                                SheetName = worksheet.Name,
                                Row = row,
                                Column = col,
                                ColumnName = GetColumnName(col),
                                InvoiceDateText = cellValue.Trim(),
                                DateValue = dateValue.Trim()
                            });

                            Console.WriteLine($"✅ [PREVIEW] Found invoice date in {worksheet.Name}: '{cellValue.Trim()}' → '{dateValue.Trim()}'");
                        }
                    }
                }
            }

            return invoiceDates;
        }

        private string FindDateValueNearInvoiceDateForPreview(ExcelWorksheet worksheet, int row, int col, int maxRows, int maxCols)
        {
            var currentCell = worksheet.Cells[row, col];

            // CASE 1: If current cell is merged, find the next cell after merged area
            if (currentCell.Merge)
            {
                int nextColAfterMerge = currentCell.End.Column + 1;

                // Try multiple columns after merged area
                for (int offset = 1; offset <= 4; offset++)
                {
                    int checkCol = nextColAfterMerge + offset - 1;
                    if (checkCol <= maxCols)
                    {
                        var checkCell = worksheet.Cells[row, checkCol];
                        var checkValue = GetCellValue(checkCell);

                        if (!string.IsNullOrEmpty(checkValue) && IsPotentialDate(checkValue))
                        {
                            return checkValue;
                        }
                    }
                }
            }

            // CASE 2: Try immediate right cell
            int rightCol = col + 1;
            if (rightCol <= maxCols)
            {
                var rightCell = worksheet.Cells[row, rightCol];
                var rightValue = GetCellValue(rightCell);

                if (!string.IsNullOrEmpty(rightValue) && IsPotentialDate(rightValue))
                {
                    return rightValue;
                }
            }

            // CASE 3: Try skipping one column
            int skipOneCol = col + 2;
            if (skipOneCol <= maxCols)
            {
                var skipCell = worksheet.Cells[row, skipOneCol];
                var skipValue = GetCellValue(skipCell);

                if (!string.IsNullOrEmpty(skipValue) && IsPotentialDate(skipValue))
                {
                    return skipValue;
                }
            }

            return string.Empty;
        }

        private List<NameError> FindNameErrorsInPreview(WorksheetPreview worksheetPreview)
        {
            var nameErrors = new List<NameError>();

            foreach (var row in worksheetPreview.Cells)
            {
                foreach (var cell in row)
                {
                    if (cell.IsNameError)
                    {
                        nameErrors.Add(new NameError
                        {
                            SheetName = worksheetPreview.Name,
                            Row = cell.Row,
                            Column = cell.Column,
                            ColumnName = cell.ColumnName
                        });
                    }
                }
            }

            return nameErrors;
        }

        private void ReadCells(ExcelWorksheet worksheet, WorksheetPreview worksheetPreview)
        {
            for (int row = 1; row <= worksheetPreview.TotalRows; row++)
            {
                var rowCells = new List<CellPreview>();
                for (int col = 1; col <= worksheetPreview.TotalColumns; col++)
                {
                    var cell = worksheet.Cells[row, col];

                    // Skip merged cells that are not the top-left cell
                    if (cell.Merge && (cell.Start.Row != row || cell.Start.Column != col))
                        continue;

                    var cellPreview = CreateCellPreview(cell, row, col);
                    rowCells.Add(cellPreview);
                }
                worksheetPreview.Cells.Add(rowCells);
            }
        }

        private CellPreview CreateCellPreview(ExcelRange cell, int row, int col)
        {
            var cellPreview = new CellPreview
            {
                Row = row,
                Column = col,
                Value = GetCellValue(cell),
                IsBold = cell.Style.Font.Bold,
                HorizontalAlignment = cell.Style.HorizontalAlignment.ToString() ?? "Left"
            };

            // Background color
            if (cell.Style.Fill.BackgroundColor.Rgb != null)
            {
                cellPreview.BackgroundColor = cell.Style.Fill.BackgroundColor.Rgb;
            }
            else
            {
                cellPreview.BackgroundColor = "FFFFFF";
            }

            // Text color
            if (cell.Style.Font.Color.Rgb != null)
            {
                cellPreview.TextColor = cell.Style.Font.Color.Rgb;
            }
            else
            {
                cellPreview.TextColor = "000000";
            }

            // Handle merged cells
            if (cell.Merge)
            {
                cellPreview.ColSpan = cell.End.Column - cell.Start.Column + 1;
                cellPreview.RowSpan = cell.End.Row - cell.Start.Row + 1;
            }

            return cellPreview;
        }

        private string GetColumnName(int column)
        {
            string columnName = "";
            while (column > 0)
            {
                int modulo = (column - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                column = (column - modulo) / 26;
            }
            return columnName;
        }

        private void ReadImages(ExcelWorksheet worksheet, WorksheetPreview worksheetPreview)
        {
            try
            {
                if (worksheet.Drawings == null || worksheet.Drawings.Count == 0)
                {
                    return;
                }

                foreach (var drawing in worksheet.Drawings)
                {
                    if (drawing is ExcelPicture picture)
                    {
                        try
                        {
                            var imagePreview = CreateImagePreview(picture);
                            if (imagePreview != null)
                            {
                                worksheetPreview.Images.Add(imagePreview);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing image {picture.Name}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading images from worksheet: {ex.Message}");
            }
        }

        private ImagePreview? CreateImagePreview(ExcelPicture picture)
        {
            try
            {
                byte[]? imageBytes = GetImageBytesFromExcelPicture(picture);

                if (imageBytes == null || imageBytes.Length == 0)
                {
                    return null;
                }

                string imageFormat = GetImageFormatFromBytes(imageBytes);

                return new ImagePreview
                {
                    Name = picture.Name ?? "Image",
                    Base64Data = Convert.ToBase64String(imageBytes),
                    Format = imageFormat,
                    Row = picture.From.Row,
                    Column = picture.From.Column
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating image preview: {ex.Message}");
                return null;
            }
        }

        private byte[]? GetImageBytesFromExcelPicture(ExcelPicture picture)
        {
            try
            {
                var excelImage = picture.Image;
                if (excelImage == null) return null;

                var imageType = excelImage.GetType();
                string[] possibleByteProperties = { "Bytes", "ImageBytes", "Data", "ImageData" };

                foreach (var propName in possibleByteProperties)
                {
                    var property = imageType.GetProperty(propName);
                    if (property != null)
                    {
                        var value = property.GetValue(excelImage);
                        if (value is byte[] bytes && bytes.Length > 0)
                        {
                            return bytes;
                        }
                    }
                }

                // Platform-specific code with conditional compilation
#if WINDOWS
        try
        {
            var imageProperty = imageType.GetProperty("Image");
            if (imageProperty?.GetValue(excelImage) is System.Drawing.Image drawingImage)
            {
                using (var ms = new MemoryStream())
                {
                    drawingImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    return ms.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Method 2 failed: {ex.Message}");
        }
#endif

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in GetImageBytesFromExcelPicture: {ex.Message}");
                return null;
            }
        }

        private string GetCellValue(ExcelRange cell)
        {
            if (cell.Value == null) return "";

            try
            {
                if (cell.Value is DateTime date)
                    return date.ToString("yyyy-MM-dd");
                else if (cell.Value is double || cell.Value is decimal)
                    return cell.Text ?? cell.Value.ToString() ?? "";
                else
                    return cell.Value.ToString() ?? "";
            }
            catch
            {
                return cell.Value.ToString() ?? "";
            }
        }

        private string GetImageFormatFromBytes(byte[] imageBytes)
        {
            if (imageBytes.Length < 8) return "png";

            if (imageBytes[0] == 0x89 && imageBytes[1] == 0x50 && imageBytes[2] == 0x4E && imageBytes[3] == 0x47)
                return "png";
            else if (imageBytes[0] == 0xFF && imageBytes[1] == 0xD8 && imageBytes[2] == 0xFF)
                return "jpeg";
            else if (imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46)
                return "gif";
            else if (imageBytes[0] == 0x42 && imageBytes[1] == 0x4D)
                return "bmp";
            else
                return "png";
        }
    }
}
