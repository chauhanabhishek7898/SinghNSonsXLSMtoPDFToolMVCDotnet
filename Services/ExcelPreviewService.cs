using ExcelToPdfConverter.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;

namespace ExcelToPdfConverter.Services
{
    public class ExcelPreviewService
    {
        public ExcelPreviewService()
        {
            // License already set in Program.cs
        }

        // ✅ ENHANCED: Better theme color detection
        private string GetThemeColorWithTint(eThemeSchemeColor? themeColor, double tint = 0)
        {
            if (themeColor == null) return "FFFFFF";

            // Enhanced theme colors mapping with tints
            var themeColorMap = new Dictionary<eThemeSchemeColor, string>
            {
                { eThemeSchemeColor.Text1, "000000" },
                { eThemeSchemeColor.Background1, "FFFFFF" },
                { eThemeSchemeColor.Text2, "1F497D" },
                { eThemeSchemeColor.Background2, "EEECE1" },
                { eThemeSchemeColor.Accent1, "4F81BD" },
                { eThemeSchemeColor.Accent2, "C0504D" },
                { eThemeSchemeColor.Accent3, "9BBB59" },
                { eThemeSchemeColor.Accent4, "8064A2" },
                { eThemeSchemeColor.Accent5, "4BACC6" },
                { eThemeSchemeColor.Accent6, "F79646" },
                { eThemeSchemeColor.Hyperlink, "0000FF" },
                { eThemeSchemeColor.FollowedHyperlink, "800080" }
            };

            if (!themeColorMap.TryGetValue(themeColor.Value, out string baseColor))
            {
                baseColor = "FFFFFF";
            }

            // Apply tint if present
            if (tint != 0)
            {
                baseColor = ApplyTintToColor(baseColor, tint);
            }

            return baseColor;
        }

        private string ApplyTintToColor(string hexColor, double tint)
        {
            if (string.IsNullOrEmpty(hexColor) || hexColor.Length != 6)
                return hexColor;

            int r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
            int g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
            int b = Convert.ToInt32(hexColor.Substring(4, 2), 16);

            if (tint > 0)
            {
                // Make lighter
                r = (int)(r + (255 - r) * tint);
                g = (int)(g + (255 - g) * tint);
                b = (int)(b + (255 - b) * tint);
            }
            else if (tint < 0)
            {
                // Make darker
                r = (int)(r * (1 + tint));
                g = (int)(g * (1 + tint));
                b = (int)(b * (1 + tint));
            }

            r = Math.Max(0, Math.Min(255, r));
            g = Math.Max(0, Math.Min(255, g));
            b = Math.Max(0, Math.Min(255, b));

            return $"{r:X2}{g:X2}{b:X2}";
        }

        private void ReadCellsWithColors(ExcelWorksheet worksheet, WorksheetPreview worksheetPreview)
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

                    var cellPreview = CreateCellPreviewWithEnhancedColors(cell, row, col);
                    rowCells.Add(cellPreview);
                }
                worksheetPreview.Cells.Add(rowCells);
            }
        }

        // ✅ IMPROVED: Cell color detection
        private CellPreview CreateCellPreviewWithEnhancedColors(ExcelRange cell, int row, int col)
        {
            var cellPreview = new CellPreview
            {
                Row = row,
                Column = col,
                Value = GetCellValue(cell),
                IsBold = cell.Style.Font.Bold,
                IsItalic = cell.Style.Font.Italic,
                Underline = cell.Style.Font.UnderLine,
                FontSize = (float)cell.Style.Font.Size,
                HorizontalAlignment = cell.Style.HorizontalAlignment.ToString() ?? "Left"
            };

            try
            {
                // ✅ ENHANCED: Background color detection with priority
                string backgroundColor = "FFFFFF";

                // Check in order of priority
                if (!string.IsNullOrEmpty(cell.Style.Fill.PatternColor?.Rgb))
                {
                    backgroundColor = cell.Style.Fill.PatternColor.Rgb;
                }
                else if (!string.IsNullOrEmpty(cell.Style.Fill.BackgroundColor?.Rgb))
                {
                    backgroundColor = cell.Style.Fill.BackgroundColor.Rgb;
                }
                else if (cell.Style.Fill.BackgroundColor?.Theme != null)
                {
                    backgroundColor = GetThemeColorWithTint(
                        cell.Style.Fill.BackgroundColor.Theme,
                        cell.Style.Fill.BackgroundColor.Tint
                    );
                }
                else if (cell.Style.Fill.BackgroundColor?.Indexed > 0)
                {
                    backgroundColor = IndexedColorToHex(cell.Style.Fill.BackgroundColor.Indexed);
                }

                // Remove alpha channel
                backgroundColor = CleanColorString(backgroundColor);
                cellPreview.BackgroundColor = backgroundColor;

                // ✅ ENHANCED: Text color detection
                string textColor = "000000";

                if (!string.IsNullOrEmpty(cell.Style.Font.Color?.Rgb))
                {
                    textColor = cell.Style.Font.Color.Rgb;
                }
                else if (cell.Style.Font.Color?.Theme != null)
                {
                    textColor = GetThemeColorWithTint(
                        cell.Style.Font.Color.Theme,
                        cell.Style.Font.Color.Tint
                    );
                }
                else if (cell.Style.Font.Color?.Indexed > 0)
                {
                    textColor = IndexedColorToHex(cell.Style.Font.Color.Indexed);
                }

                textColor = CleanColorString(textColor);
                cellPreview.TextColor = textColor;

                Console.WriteLine($"🎨 Cell {GetColumnName(col)}{row}: BG={backgroundColor}, Text={textColor}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Color error at {GetColumnName(col)}{row}: {ex.Message}");
                cellPreview.BackgroundColor = "FFFFFF";
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

        private string CleanColorString(string color)
        {
            if (string.IsNullOrEmpty(color)) return "FFFFFF";

            // Remove #
            color = color.Replace("#", "");

            // Handle ARGB (8 characters)
            if (color.Length == 8)
            {
                // Remove alpha channel (first 2 characters)
                color = color.Substring(2);
            }
            else if (color.Length == 6)
            {
                // Already RGB
                return color;
            }
            else if (color.Length == 3)
            {
                // Expand #RGB to #RRGGBB
                color = $"{color[0]}{color[0]}{color[1]}{color[1]}{color[2]}{color[2]}";
            }

            // Ensure uppercase
            return color.ToUpper();
        }

        // ✅ Indexed color to hex conversion
        private string IndexedColorToHex(int indexed)
        {
            var indexedColors = new Dictionary<int, string>
            {
                { 0, "000000" },  // Black
                { 1, "FFFFFF" },  // White
                { 2, "FF0000" },  // Red
                { 3, "00FF00" },  // Green
                { 4, "0000FF" },  // Blue
                { 5, "FFFF00" },  // Yellow
                { 6, "FF00FF" },  // Magenta
                { 7, "00FFFF" },  // Cyan
                { 8, "800000" },  // Dark Red
                { 9, "008000" },  // Dark Green
                { 10, "000080" }, // Dark Blue
                { 11, "808000" }, // Olive
                { 12, "800080" }, // Purple
                { 13, "008080" }, // Teal
                { 14, "C0C0C0" }, // Silver
                { 15, "808080" }, // Gray
                { 16, "9999FF" },
                { 17, "993366" },
                { 18, "FFFFCC" },
                { 19, "CCFFFF" },
                { 20, "660066" },
                { 21, "FF8080" },
                { 22, "0066CC" },
                { 23, "CCCCFF" },
                { 53, "993300" },
                { 55, "339966" },
                { 56, "CCCC99" },
                { 57, "A5A5A5" },
                { 58, "FFCC00" },
                { 59, "FFFF99" },
                { 60, "99CCFF" },
                { 61, "FF99CC" },
                { 62, "CC99FF" },
                { 63, "FFCC99" },
                { 64, "3366FF" },
                { 65, "33CCCC" },
                { 66, "99CC00" },
                { 67, "FFCC00" },
                { 68, "FF9900" },
                { 69, "FF6600" },
                { 70, "666699" },
                { 71, "969696" },
                { 72, "003366" },
                { 73, "339966" },
                { 74, "003300" },
                { 75, "333300" },
                { 76, "993300" },
                { 77, "993366" },
                { 78, "333399" },
                { 79, "333333" }
            };

            return indexedColors.ContainsKey(indexed) ? indexedColors[indexed] : "000000";
        }

        // ✅ Sheet orientation analysis class
        public class SheetOrientationInfo
        {
            public string SheetName { get; set; } = string.Empty;
            public string SuggestedOrientation { get; set; } = "Portrait";
            public double WidthToHeightRatio { get; set; }
            public int TotalColumns { get; set; }
            public int TotalRows { get; set; }
            public bool HasWideContent { get; set; }
            public double AverageColumnWidth { get; set; }
            public List<string> RecommendedOrientations { get; set; } = new List<string> { "Portrait", "Landscape" };
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
                InvoiceNumbers = new List<InvoiceNumber>(), // नई property
                DateRanges = new List<DateRange>(), // नई property
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

                        // FIRST: Find all invoice dates using the robust method
                        var allInvoiceDates = new List<InvoiceDate>();
                        var allInvoiceNumbers = new List<InvoiceNumber>(); // नया
                        var allDateRanges = new List<DateRange>(); // नया

                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null) continue;

                            var invoiceDates = FindInvoiceDatesInWorksheet(worksheet);
                            allInvoiceDates.AddRange(invoiceDates);

                            // नया: Find invoice numbers
                            var invoiceNumbers = FindInvoiceNumbersInWorksheet(worksheet);
                            allInvoiceNumbers.AddRange(invoiceNumbers);

                            // नया: Find date ranges
                            var dateRanges = FindDateRangesInWorksheet(worksheet);
                            allDateRanges.AddRange(dateRanges);

                            Console.WriteLine($"Found in {worksheet.Name}: Dates={invoiceDates.Count}, InvoiceNos={invoiceNumbers.Count}, DateRanges={dateRanges.Count}");
                        }

                        // THEN: Create worksheet previews with enhanced color reading
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null)
                            {
                                Console.WriteLine($"Worksheet {worksheet.Name} has no data, skipping.");
                                continue;
                            }

                            Console.WriteLine($"Processing worksheet: {worksheet.Name}");
                            var worksheetPreview = CreateWorksheetPreviewWithEnhancedColors(worksheet);

                            // Get specific data for this sheet
                            var sheetInvoiceDates = allInvoiceDates
                                .Where(id => id.SheetName == worksheet.Name)
                                .ToList();

                            var sheetInvoiceNumbers = allInvoiceNumbers
                                .Where(inv => inv.SheetName == worksheet.Name)
                                .ToList();

                            var sheetDateRanges = allDateRanges
                                .Where(dr => dr.SheetName == worksheet.Name)
                                .ToList();

                            // Assign to worksheet preview
                            worksheetPreview.InvoiceDates = sheetInvoiceDates;
                            worksheetPreview.InvoiceNumbers = sheetInvoiceNumbers;
                            worksheetPreview.DateRanges = sheetDateRanges;

                            previewModel.Worksheets.Add(worksheetPreview);

                            // Add to global collections
                            previewModel.AllNameErrors.AddRange(worksheetPreview.NameErrors);

                            // Create file selection entry
                            var fileSelection = new FileSelection
                            {
                                FileName = previewModel.OriginalFileName,
                                SheetName = worksheet.Name,
                                SortOrder = worksheet.Index,
                                HasNameErrors = worksheetPreview.NameErrors.Count > 0,
                                HasInvoiceDates = sheetInvoiceDates.Count > 0,
                                HasInvoiceNumbers = sheetInvoiceNumbers.Count > 0, // नया
                                HasDateRanges = sheetDateRanges.Count > 0, // नया
                                NameErrors = worksheetPreview.NameErrors,
                                InvoiceDates = sheetInvoiceDates,
                                InvoiceNumbers = sheetInvoiceNumbers, // नया
                                DateRanges = sheetDateRanges // नया
                            };
                            previewModel.FileSelections.Add(fileSelection);

                            Console.WriteLine($"FileSelection for {worksheet.Name}: InvoiceNos={fileSelection.InvoiceNumbers.Count}, DateRanges={fileSelection.DateRanges.Count}");
                        }

                        // Set the global collections
                        previewModel.AllInvoiceDates = allInvoiceDates;
                        previewModel.InvoiceNumbers = allInvoiceNumbers; // नया
                        previewModel.DateRanges = allDateRanges; // नया
                    }
                }

                Console.WriteLine($"Preview generated for {previewModel.Worksheets.Count} worksheets");
                Console.WriteLine($"Found: NameErrors={previewModel.AllNameErrors.Count}, InvoiceDates={previewModel.AllInvoiceDates.Count}, " +
                                $"InvoiceNumbers={previewModel.InvoiceNumbers.Count}, DateRanges={previewModel.DateRanges.Count}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating preview: {ex}");
                throw new Exception($"Error reading Excel file: {ex.Message}", ex);
            }

            return previewModel;
        }

        // ✅ NEW: Create worksheet preview with enhanced color reading
        private WorksheetPreview CreateWorksheetPreviewWithEnhancedColors(ExcelWorksheet worksheet)
        {
            var worksheetPreview = new WorksheetPreview
            {
                Name = worksheet.Name,
                Index = worksheet.Index,
                TotalRows = Math.Min(100, worksheet.Dimension.End.Row),
                TotalColumns = Math.Min(50, worksheet.Dimension.End.Column), // Increased to 50 for better coverage
                Cells = new List<List<CellPreview>>(),
                Images = new List<ImagePreview>(),
                NameErrors = new List<NameError>(),
                InvoiceDates = new List<InvoiceDate>(),
                InvoiceNumbers = new List<InvoiceNumber>(), // नया
                DateRanges = new List<DateRange>() // नया
            };

            // Read cells with enhanced color detection
            ReadCellsWithColors(worksheet, worksheetPreview);

            // Find name errors
            worksheetPreview.NameErrors = FindNameErrorsInPreview(worksheetPreview);

            // Read images
            ReadImages(worksheet, worksheetPreview);

            return worksheetPreview;
        }

        private List<InvoiceNumber> FindInvoiceNumbersInWorksheet(ExcelWorksheet worksheet)
        {
            var invoiceNumbers = new List<InvoiceNumber>();

            if (worksheet.Dimension == null)
            {
                Console.WriteLine($"Worksheet {worksheet.Name} has no data, skipping invoice numbers.");
                return invoiceNumbers;
            }

            int maxRows = worksheet.Dimension.End.Row;
            int maxCols = worksheet.Dimension.End.Column;

            Console.WriteLine($"🔍 Searching for invoice numbers in {worksheet.Name} - Rows: {maxRows}, Columns: {maxCols}");

            // ✅ CREDIT NOTE SHEET को SKIP करें
            if (worksheet.Name.Contains("Credit Note", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"⏭️ Skipping Credit Note worksheet for invoice number search");
                return invoiceNumbers; // Empty list return करें
            }

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

                    // CASE 1: Check if same cell has both label and number (e.g., "Invoice No.: SSJ/1194")
                    if (SameCellHasInvoiceNumber(cellValue, out string labelText, out string numberValue))
                    {
                        Console.WriteLine($"📍 Found invoice number in same cell at {GetColumnName(col)}{row}: '{cellValue}'");

                        invoiceNumbers.Add(new InvoiceNumber
                        {
                            SheetName = worksheet.Name,
                            Row = row,
                            Column = col,
                            ColumnName = GetColumnName(col),
                            LabelText = labelText,
                            Number = numberValue
                        });

                        Console.WriteLine($"✅ CONFIRMED invoice number in same cell: '{labelText}' → '{numberValue}'");
                    }
                    // CASE 2: Cell contains only invoice label (like invoice date)
                    else if (IsInvoiceNumberText(cellValue))
                    {
                        Console.WriteLine($"📍 Found invoice number text at {GetColumnName(col)}{row}: '{cellValue}'");

                        // Look for invoice number value nearby
                        var invoiceNumberValue = FindInvoiceNumberValueNearby(worksheet, row, col, maxRows, maxCols);

                        if (!string.IsNullOrEmpty(invoiceNumberValue))
                        {
                            invoiceNumbers.Add(new InvoiceNumber
                            {
                                SheetName = worksheet.Name,
                                Row = row,
                                Column = col,
                                ColumnName = GetColumnName(col),
                                LabelText = cellValue.Trim(),
                                Number = invoiceNumberValue.Trim()
                            });

                            Console.WriteLine($"✅ CONFIRMED invoice number: '{cellValue.Trim()}' → '{invoiceNumberValue.Trim()}'");
                        }
                        else
                        {
                            Console.WriteLine($"❌ No invoice number found near '{cellValue}' at {GetColumnName(col)}{row}");
                        }
                    }
                }
            }

            Console.WriteLine($"📊 Total invoice numbers found in {worksheet.Name}: {invoiceNumbers.Count}");
            return invoiceNumbers;
        }

        private bool SameCellHasInvoiceNumber(string cellValue, out string labelText, out string numberValue)
        {
            labelText = string.Empty;
            numberValue = string.Empty;

            if (string.IsNullOrEmpty(cellValue)) return false;

            var trimmedValue = cellValue.Trim();

            // Check if cell contains both invoice label and a number
            if (ContainsInvoiceLabelText(trimmedValue))
            {
                // Try to extract number from same cell
                // Patterns: "Invoice No.: SSJ/1194", "Invoice Number: SSI/1194", etc.

                // Split by common separators
                char[] separators = { ':', '：', '-', ' ', '\t' };

                foreach (var separator in separators)
                {
                    if (trimmedValue.Contains(separator))
                    {
                        var parts = trimmedValue.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 2)
                        {
                            // First part should be the label
                            var possibleLabel = parts[0].Trim();
                            var possibleNumber = parts[1].Trim();

                            if (ContainsInvoiceLabelText(possibleLabel) &&
                                !string.IsNullOrEmpty(possibleNumber) &&
                                possibleNumber.Length > 0)
                            {
                                labelText = possibleLabel;
                                numberValue = possibleNumber;
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private bool ContainsInvoiceLabelText(string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue)) return false;

            var lowerValue = cellValue.ToLower();

            return lowerValue.Contains("invoice number") ||
                   lowerValue.Contains("invoice no") ||
                   lowerValue.Contains("invoiceno") ||
                   lowerValue.Contains("invoice_no") ||
                   lowerValue.Contains("invoice no.") ||
                   lowerValue.Contains("inv no") ||
                   lowerValue.Contains("inv. no") ||
                   lowerValue.Contains("inv no.");
         
        }

        private bool IsInvoiceNumberText(string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue)) return false;

            var normalizedValue = cellValue.Trim().ToUpper();

            var exactPatterns = new[]
            {
        "INVOICE NUMBER",
        "INVOICENUMBER",
        "INVOICE_NUMBER",
        "INVOICE  NUMBER",
        "INVOICE NO",
        "INVOICENO",
        "INVOICE_NO",
        "INVOICE NO.",
        "INV NO",
        "INV. NO",
        "INV NO."
   
    };

            var containsPatterns = new[]
            {
        "INVOICE NUMBER",
        "INVOICE NO",
        "INV NO"
     
    };

            // Exact match check
            if (exactPatterns.Any(pattern => normalizedValue == pattern))
                return true;

            // Contains check (same as invoice date)
            foreach (var pattern in containsPatterns)
            {
                if (normalizedValue.Contains(pattern) &&
                    (normalizedValue.StartsWith(pattern) || normalizedValue.EndsWith(pattern) ||
                     normalizedValue.IndexOf(pattern) > 0))
                {
                    return true;
                }
            }

            return false;
        }


        private string FindInvoiceNumberValueNearby(ExcelWorksheet worksheet, int row, int col, int maxRows, int maxCols)
        {
            var currentCell = worksheet.Cells[row, col];

            // If current cell is merged, start from after merged area
            int startCol = currentCell.Merge ? currentCell.End.Column + 1 : col + 1;

            // Look for a valid value (skip colons and empty cells)
            for (int offset = 1; offset <= 2; offset++) // Check up to 4 cells to the right
            {
                int checkCol = startCol + offset - 1;
                if (checkCol <= maxCols)
                {
                    var checkCell = worksheet.Cells[row, checkCol];
                    var checkValue = GetCellValue(checkCell);

                    Console.WriteLine($"   Checking cell {GetColumnName(checkCol)}: '{checkValue}'");

                    // Skip empty cells, colons, and other non-value cells
                    if (!string.IsNullOrEmpty(checkValue) &&
                        !IsColonOrSeparator(checkValue)
                        //&&
                        //!IsInvoiceLabel(checkValue)
                        )
                    {
                        Console.WriteLine($"   ✅ Found valid invoice number at {GetColumnName(checkCol)}: '{checkValue}'");
                        return checkValue;
                    }
                }
            }

            return string.Empty;
        }

        private bool IsColonOrSeparator(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            var trimmed = value.Trim();

            // Check for colon, dash, or other separators
            return trimmed == ":" ||
                   trimmed == "：" ||
                   trimmed == "-" ||
                   trimmed == "–" ||
                   trimmed.Length <= 2; // Very short values are likely separators
        }



        // नया: Find Date Ranges in worksheet (01 Oct 2025 To 31 Oct 2025 format)
        private List<DateRange> FindDateRangesInWorksheet(ExcelWorksheet worksheet)
        {
            var dateRanges = new List<DateRange>();

            if (worksheet.Dimension == null)
            {
                Console.WriteLine($"Worksheet {worksheet.Name} has no data, skipping date ranges.");
                return dateRanges;
            }

            int maxRows = worksheet.Dimension.End.Row;
            int maxCols = worksheet.Dimension.End.Column;

            Console.WriteLine($"🔍 Searching for date ranges in {worksheet.Name} - Rows: {maxRows}, Columns: {maxCols}");

            // ✅ CREDIT NOTE SHEET को SKIP करें
            if (worksheet.Name.Contains("Credit Note", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"⏭️ Skipping Credit Note worksheet for date range search");
                return dateRanges; // Empty list return करें
            }

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

                    // Check for date range pattern: 01 Oct 2025 To 31 Oct 2025
                    if (IsDateRangePatternWithMonthNames(cellValue))
                    {
                        Console.WriteLine($"📍 Found date range at {GetColumnName(col)}{row}: '{cellValue}'");

                        dateRanges.Add(new DateRange
                        {
                            SheetName = worksheet.Name,
                            Row = row,
                            Column = col,
                            ColumnName = GetColumnName(col),
                            DateRangeText = cellValue.Trim()
                        });

                        Console.WriteLine($"✅ CONFIRMED date range: '{cellValue.Trim()}'");
                    }
                }
            }

            Console.WriteLine($"📊 Total date ranges found in {worksheet.Name}: {dateRanges.Count}");
            return dateRanges;
        }

        private bool IsDateRangePatternWithMonthNames(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            var trimmedValue = value.Trim();

            Console.WriteLine($"Checking for date range pattern: '{trimmedValue}'");

            // Pattern 1: 01 Oct 2025 To 31 Oct 2025 (with spaces and "To")
            var pattern1 = @"^(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})\s+To\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})$";

            // Pattern 2: 01 October 2025 To 31 October 2025 (full month names)
            var pattern2 = @"^(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\s+To\s+(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})$";

            // Pattern 3: 01-Oct-2025 To 31-Oct-2025 (with hyphens)
            var pattern3 = @"^(\d{1,2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*-(\d{4})\s+To\s+(\d{1,2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*-(\d{4})$";

            // Pattern 4: 01/Oct/2025 To 31/Oct/2025 (with slashes)
            var pattern4 = @"^(\d{1,2})/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*/(\d{4})\s+To\s+(\d{1,2})/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*/(\d{4})$";

            // Pattern 5: With "TO" in uppercase
            var pattern5 = @"^(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})\s+TO\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})$";

            // Pattern 1: dd-mm-yyyy TO dd-mm-yyyy (with single digit day/month)
            var pattern6 = @"^\d{1,2}-\d{1,2}-\d{4}\s+TO\s+\d{1,2}-\d{1,2}-\d{4}$";

            // Pattern 2: dd/mm/yyyy TO dd/mm/yyyy
            var pattern7 = @"^\d{1,2}/\d{1,2}/\d{4}\s+TO\s+\d{1,2}/\d{1,2}/\d{4}$";


            bool matches = Regex.IsMatch(trimmedValue, pattern1, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(trimmedValue, pattern2, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(trimmedValue, pattern3, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(trimmedValue, pattern4, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(trimmedValue, pattern5, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(trimmedValue, pattern6, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(trimmedValue, pattern7, RegexOptions.IgnoreCase);

            if (matches)
            {
                Console.WriteLine($"✅ Pattern matched: '{trimmedValue}'");

                // Debug: Show which pattern matched
                if (Regex.IsMatch(trimmedValue, pattern1, RegexOptions.IgnoreCase))
                    Console.WriteLine("   Matched pattern 1");
                if (Regex.IsMatch(trimmedValue, pattern2, RegexOptions.IgnoreCase))
                    Console.WriteLine("   Matched pattern 2");
                if (Regex.IsMatch(trimmedValue, pattern3, RegexOptions.IgnoreCase))
                    Console.WriteLine("   Matched pattern 3");
                if (Regex.IsMatch(trimmedValue, pattern4, RegexOptions.IgnoreCase))
                    Console.WriteLine("   Matched pattern 4");
                if (Regex.IsMatch(trimmedValue, pattern5, RegexOptions.IgnoreCase))
                    Console.WriteLine("   Matched pattern 5");
            }

            return matches;
        }

        
        public ValidationResult QuickValidate(IFormFile excelFile)
        {
            var validationResult = new ValidationResult
            {
                FileName = excelFile.FileName,
                NameErrors = new List<NameError>(),
                InvoiceDates = new List<InvoiceDate>(),
                InvoiceNumbers = new List<InvoiceNumber>(), // नया
                DateRanges = new List<DateRange>() // नया
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

                            // नया: Find invoice numbers
                            var invoiceNumbers = FindInvoiceNumbersInWorksheet(worksheet);
                            validationResult.InvoiceNumbers.AddRange(invoiceNumbers);

                            // नया: Find date ranges
                            var dateRanges = FindDateRangesInWorksheet(worksheet);
                            validationResult.DateRanges.AddRange(dateRanges);
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

            // ✅ CREDIT NOTE SHEET को SKIP करें
            if (worksheet.Name.Contains("Credit Note", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"⏭️ Skipping Credit Note worksheet for name error search");
                return nameErrors; // Empty list return करें
            }


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

            // ✅ CREDIT NOTE SHEET को SKIP करें
            if (worksheet.Name.Contains("Credit Note", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"⏭️ Skipping Credit Note worksheet for invoice date search");
                return invoiceDates; // Empty list return करें
            }

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
                            Console.WriteLine($"   ✅ Found date after merged area at {GetColumnName(checkCol)}: '{checkValue}'");
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
                    Console.WriteLine($"   ✅ Found date in immediate right cell: '{rightValue}'");
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
                    Console.WriteLine($"   ✅ Found date by skipping one column: '{skipValue}'");
                    return skipValue;
                }
            }

            return string.Empty;
        }

        private bool IsPotentialDate(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            var trimmedValue = value.Trim();

            // Check for common date patterns
            if (trimmedValue.Contains(",") && trimmedValue.Contains(" "))
                return true;

            if (Regex.IsMatch(trimmedValue, @"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}"))
                return true;

            if (Regex.IsMatch(trimmedValue, @"\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)", RegexOptions.IgnoreCase))
                return true;

            var lowerValue = trimmedValue.ToLower();
            if (!lowerValue.Contains("invoice") &&
                !lowerValue.Contains("date") &&
                !lowerValue.Contains("no") &&
                !lowerValue.Contains("number") &&
                trimmedValue.Length > 3)
                return true;

            return false;
        }

        private bool IsInvoiceDateText(string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue)) return false;

            var normalizedValue = cellValue.Trim().ToUpper();

            var exactPatterns = new[]
            {
                "INVOICE DATE",
                "INVOICEDATE",
                "INVOICE_DATE",
                "INVOICE  DATE",
                "INVOICE DATE\r\n",
                "INVOICE DATE\n"
            };

            var containsPatterns = new[]
            {
                "INVOICE DATE",
                "INVOICEDATE"
            };

            return exactPatterns.Any(pattern => normalizedValue == pattern) ||
                   containsPatterns.Any(pattern => normalizedValue.Contains(pattern));
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
