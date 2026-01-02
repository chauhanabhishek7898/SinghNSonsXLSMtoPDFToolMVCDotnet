using ExcelToPdfConverter.Services;
using Microsoft.AspNetCore.Http;

namespace ExcelToPdfConverter.Models
{
    public class ExcelUploadModel
    {
        public IFormFile? ExcelFile { get; set; }
    }

    public class PdfPreviewWithFitToPageRequest
    {
        public string SessionId { get; set; }
        public List<string> SelectedSheets { get; set; }
        public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
        public bool IncludeMergedPdfs { get; set; }
    }

    public class WorksheetPreview
    {
        public string Name { get; set; } = string.Empty;
        public int Index { get; set; }
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public List<List<CellPreview>> Cells { get; set; } = new List<List<CellPreview>>();
        public List<ImagePreview> Images { get; set; } = new List<ImagePreview>();
        public List<NameError> NameErrors { get; set; } = new List<NameError>();
        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
        public bool HasNameErrors => NameErrors.Count > 0;
        public bool HasInvoiceDates => InvoiceDates.Count > 0;
    }

    public class CellPreview
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Value { get; set; } = string.Empty;
        public string BackgroundColor { get; set; } = string.Empty;
        public string TextColor { get; set; } = string.Empty;
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool Underline { get; set; }
        public double FontSize { get; set; }
        public string HorizontalAlignment { get; set; } = string.Empty;
        public int ColSpan { get; set; } = 1;
        public int RowSpan { get; set; } = 1;
        public bool IsNameError => Value?.Contains("#NAME?") == true;
        public string ColumnName => GetColumnName(Column);

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
    }

    public class NameError
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string ColumnName { get; set; } = string.Empty;
        public string Location => $"{ColumnName}{Row}";
        public string FullLocation => $"Row {Row}, Column {ColumnName}";
        public string SheetName { get; set; } = string.Empty;
    }

    public class InvoiceDate
    {
        public string SheetName { get; set; } = string.Empty;
        public int Row { get; set; }
        public int Column { get; set; }
        public string ColumnName { get; set; } = string.Empty;
        public string InvoiceDateText { get; set; } = string.Empty;
        public string DateValue { get; set; } = string.Empty;
        public string Location => $"{ColumnName}{Row}";
        public string FullInfo => $"'{InvoiceDateText}' → '{DateValue}' at {Location}";
    }

    public class ImagePreview
    {
        public string Name { get; set; } = string.Empty;
        public string Base64Data { get; set; } = string.Empty;
        public string Format { get; set; } = string.Empty;
        public int Row { get; set; }
        public int Column { get; set; }
    }

    public class ConversionResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string PdfFilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public int TotalPages { get; set; }
    }

    public class ValidationResult
    {
        public string FileName { get; set; } = string.Empty;
        public List<NameError> NameErrors { get; set; } = new List<NameError>();
        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
        public bool HasNameErrors => NameErrors.Count > 0;
        public bool HasInvoiceDates => InvoiceDates.Count > 0;
    }

    public class PreviewModel
    {
        public string OriginalFileName { get; set; } = string.Empty;
        public List<WorksheetPreview> Worksheets { get; set; } = new List<WorksheetPreview>();
        public string SessionId { get; set; } = string.Empty;
        public List<NameError> AllNameErrors { get; set; } = new List<NameError>();
        public List<InvoiceDate> AllInvoiceDates { get; set; } = new List<InvoiceDate>();
        public List<FileSelection> FileSelections { get; set; } = new List<FileSelection>();
        public Dictionary<string, string> SuggestedOrientations { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, ExcelPreviewService.SheetOrientationInfo> SheetOrientationAnalysis { get; set; }
            = new Dictionary<string, ExcelPreviewService.SheetOrientationInfo>();
        public bool HasNameErrors => AllNameErrors.Count > 0;
        public bool HasInvoiceDates => AllInvoiceDates.Count > 0;
    }

    public class FileSelection
    {
        public string FileName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public bool IsSelected { get; set; } = true;
        public int SortOrder { get; set; }
        public bool HasNameErrors { get; set; }
        public bool HasInvoiceDates { get; set; }
        public List<NameError> NameErrors { get; set; } = new List<NameError>();
        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
        public string Orientation { get; set; } = "Landscape";
    }

    public class FileNamesModel
    {
        public string ExcelFileName { get; set; } = string.Empty;
        public List<string> PdfFileNames { get; set; } = new List<string>();
        public int TotalPdfFiles { get; set; }
    }

    public class PdfReorderingRequest
    {
        public string SessionId { get; set; }
        public List<string> SelectedSheets { get; set; }
        public List<PageOrderInfo> PageOrderData { get; set; }
        public Dictionary<int, int> RotationData { get; set; }
    }

    public class PdfConversionResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string FileName { get; set; }
        public byte[] FileData { get; set; }
    }

    public class PdfPreviewRequest
    {
        public string SessionId { get; set; }
        public List<string> SelectedSheets { get; set; }
    }

    public class PdfPreviewResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string PdfData { get; set; }
        public string FileName { get; set; }
        public string GeneratedTime { get; set; }
        public int TotalPages { get; set; }
    }

    public class MarginInfo
    {
        public float Top { get; set; } = 10;
        public float Bottom { get; set; } = 10;
        public float Left { get; set; } = 10;
        public float Right { get; set; } = 10;
    }

    public class PositionInfo
    {
        public string Horizontal { get; set; } = "center";
        public string Vertical { get; set; } = "middle";
    }

    public class PageOrderInfo
    {
        public int OriginalPage { get; set; }
        public int CurrentOrder { get; set; }
        public bool Visible { get; set; }
        public int Rotation { get; set; }
        public string Orientation { get; set; } = "portrait";
        public string FitMode { get; set; } = "fitToPage";
        public MarginInfo Margin { get; set; } = new MarginInfo();
        public PositionInfo Position { get; set; } = new PositionInfo();
    }

    public class PdfReorderingRequestWithAllSettings
    {
        public string SessionId { get; set; }
        public List<string> SelectedSheets { get; set; }
        public List<PageOrderInfo> PageOrderData { get; set; }
        public Dictionary<int, int> RotationData { get; set; }
        public Dictionary<int, string> OrientationData { get; set; }
        public Dictionary<int, string> FitModeData { get; set; }
        public Dictionary<int, MarginInfo> MarginData { get; set; }
        public Dictionary<int, PositionInfo> PositionData { get; set; }
    }

    public class PageOrderInfoWithRotation
    {
        public int OriginalPage { get; set; }
        public int CurrentOrder { get; set; }
        public bool Visible { get; set; }
        public string Orientation { get; set; } = "portrait";
        public int Rotation { get; set; } = 0;
    }

    public class PdfRequestWithRotation
    {
        public string SessionId { get; set; }
        public List<string> SelectedSheets { get; set; }
        public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
        public Dictionary<int, string> OrientationData { get; set; }
        public Dictionary<int, int> RotationData { get; set; }
    }

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
}



//using ExcelToPdfConverter.Services;
//using Microsoft.AspNetCore.Http;

//namespace ExcelToPdfConverter.Models
//{
//    public class ExcelUploadModel
//    {
//        public IFormFile? ExcelFile { get; set; }
//    }

//    public class PdfPreviewWithFitToPageRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
//        public bool IncludeMergedPdfs { get; set; }
//    }

//    public class WorksheetPreview
//    {
//        public string Name { get; set; } = string.Empty;
//        public int Index { get; set; }
//        public int TotalRows { get; set; }
//        public int TotalColumns { get; set; }
//        public List<List<CellPreview>> Cells { get; set; } = new List<List<CellPreview>>();
//        public List<ImagePreview> Images { get; set; } = new List<ImagePreview>();
//        public List<NameError> NameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
//        public bool HasNameErrors => NameErrors.Count > 0;
//        public bool HasInvoiceDates => InvoiceDates.Count > 0;
//    }

//    public class CellPreview
//    {
//        public int Row { get; set; }
//        public int Column { get; set; }
//        public string Value { get; set; } = string.Empty;
//        public string BackgroundColor { get; set; } = string.Empty;
//        public string TextColor { get; set; } = string.Empty;
//        public bool IsBold { get; set; }
//        public bool IsItalic { get; set; }
//        public bool Underline { get; set; }
//        public double FontSize { get; set; }
//        public string HorizontalAlignment { get; set; } = string.Empty;
//        public int ColSpan { get; set; } = 1;
//        public int RowSpan { get; set; } = 1;
//        public bool IsNameError => Value?.Contains("#NAME?") == true;
//        public string ColumnName => GetColumnName(Column);

//        private string GetColumnName(int column)
//        {
//            string columnName = "";
//            while (column > 0)
//            {
//                int modulo = (column - 1) % 26;
//                columnName = Convert.ToChar('A' + modulo) + columnName;
//                column = (column - modulo) / 26;
//            }
//            return columnName;
//        }
//    }

//    public class NameError
//    {
//        public int Row { get; set; }
//        public int Column { get; set; }
//        public string ColumnName { get; set; } = string.Empty;
//        public string Location => $"{ColumnName}{Row}";
//        public string FullLocation => $"Row {Row}, Column {ColumnName}";
//        public string SheetName { get; set; } = string.Empty;
//    }

//    public class InvoiceDate
//    {
//        public string SheetName { get; set; } = string.Empty;
//        public int Row { get; set; }
//        public int Column { get; set; }
//        public string ColumnName { get; set; } = string.Empty;
//        public string InvoiceDateText { get; set; } = string.Empty;
//        public string DateValue { get; set; } = string.Empty;
//        public string Location => $"{ColumnName}{Row}";
//        public string FullInfo => $"'{InvoiceDateText}' → '{DateValue}' at {Location}";
//    }

//    public class ImagePreview
//    {
//        public string Name { get; set; } = string.Empty;
//        public string Base64Data { get; set; } = string.Empty;
//        public string Format { get; set; } = string.Empty;
//        public int Row { get; set; }
//        public int Column { get; set; }
//    }

//    public class ConversionResult
//    {
//        public bool Success { get; set; }
//        public string Message { get; set; } = string.Empty;
//        public string PdfFilePath { get; set; } = string.Empty;
//        public string FileName { get; set; } = string.Empty;
//        public int TotalPages { get; set; }
//    }

//    public class ValidationResult
//    {
//        public string FileName { get; set; } = string.Empty;
//        public List<NameError> NameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
//        public bool HasNameErrors => NameErrors.Count > 0;
//        public bool HasInvoiceDates => InvoiceDates.Count > 0;
//    }

//    public class PreviewModel
//    {
//        public string OriginalFileName { get; set; } = string.Empty;
//        public List<WorksheetPreview> Worksheets { get; set; } = new List<WorksheetPreview>();
//        public string SessionId { get; set; } = string.Empty;
//        public List<NameError> AllNameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> AllInvoiceDates { get; set; } = new List<InvoiceDate>();
//        public List<FileSelection> FileSelections { get; set; } = new List<FileSelection>();
//        public Dictionary<string, string> SuggestedOrientations { get; set; } = new Dictionary<string, string>();
//        public Dictionary<string, ExcelPreviewService.SheetOrientationInfo> SheetOrientationAnalysis { get; set; }
//            = new Dictionary<string, ExcelPreviewService.SheetOrientationInfo>();
//        public bool HasNameErrors => AllNameErrors.Count > 0;
//        public bool HasInvoiceDates => AllInvoiceDates.Count > 0;
//    }

//    public class FileSelection
//    {
//        public string FileName { get; set; } = string.Empty;
//        public string SheetName { get; set; } = string.Empty;
//        public bool IsSelected { get; set; } = true;
//        public int SortOrder { get; set; }
//        public bool HasNameErrors { get; set; }
//        public bool HasInvoiceDates { get; set; }
//        public List<NameError> NameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
//        public string Orientation { get; set; } = "Landscape";
//    }

//    public class FileNamesModel
//    {
//        public string ExcelFileName { get; set; } = string.Empty;
//        public List<string> PdfFileNames { get; set; } = new List<string>();
//        public int TotalPdfFiles { get; set; }
//    }

//    public class PdfReorderingRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfo> PageOrderData { get; set; }
//        public Dictionary<int, int> RotationData { get; set; }
//    }

//    public class PdfConversionResponse
//    {
//        public bool Success { get; set; }
//        public string Message { get; set; }
//        public string FileName { get; set; }
//        public byte[] FileData { get; set; }
//    }

//    public class PdfPreviewRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//    }

//    public class PdfPreviewResponse
//    {
//        public bool Success { get; set; }
//        public string Message { get; set; }
//        public string PdfData { get; set; }
//        public string FileName { get; set; }
//        public string GeneratedTime { get; set; }
//        public int TotalPages { get; set; }
//    }

//    // Existing models के साथ ये नए models जोड़ें
//    public class MarginInfo
//    {
//        public float Top { get; set; } = 10;
//        public float Bottom { get; set; } = 10;
//        public float Left { get; set; } = 10;
//        public float Right { get; set; } = 10;
//    }

//    public class PositionInfo
//    {
//        public string Horizontal { get; set; } = "center";
//        public string Vertical { get; set; } = "middle";
//    }

//    public class PageOrderInfo
//    {
//        public int OriginalPage { get; set; }
//        public int CurrentOrder { get; set; }
//        public bool Visible { get; set; }
//        public int Rotation { get; set; }
//        public string Orientation { get; set; } = "portrait";
//        public string FitMode { get; set; } = "fitToPage";
//        public MarginInfo Margin { get; set; } = new MarginInfo();
//        public PositionInfo Position { get; set; } = new PositionInfo();
//    }

//    public class PdfReorderingRequestWithAllSettings
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfo> PageOrderData { get; set; }
//        public Dictionary<int, int> RotationData { get; set; }
//        public Dictionary<int, string> OrientationData { get; set; }
//        public Dictionary<int, string> FitModeData { get; set; }
//        public Dictionary<int, MarginInfo> MarginData { get; set; }
//        public Dictionary<int, PositionInfo> PositionData { get; set; }
//    }

//    public class PageOrderInfoWithRotation
//    {
//        public int OriginalPage { get; set; }
//        public int CurrentOrder { get; set; }
//        public bool Visible { get; set; }
//        public string Orientation { get; set; } = "portrait";
//        public int Rotation { get; set; } = 0;
//    }

//    public class PdfRequestWithRotation
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
//        public Dictionary<int, string> OrientationData { get; set; }
//        public Dictionary<int, int> RotationData { get; set; }
//    }

//    public class RemovePdfRequest
//    {
//        public string SessionId { get; set; }
//        public string FileName { get; set; }
//    }

//    public class MergePdfRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> FileNames { get; set; }
//    }

//    public class DownloadPdfRequest
//    {
//        public string SessionId { get; set; }
//        public string FileName { get; set; }
//    }

//    public class MergedPdfInfo
//    {
//        public string FileName { get; set; }
//        public string FilePath { get; set; }
//        public long FileSize { get; set; }
//        public int TotalPages { get; set; }
//        public DateTime CreatedAt { get; set; }
//    }

//    public class UpdatePdfFilesRequest
//    {
//        public string SessionId { get; set; }
//        public List<PdfFileInfo> UploadedFiles { get; set; }
//    }

//    public class PdfFileInfo
//    {
//        public string Name { get; set; }
//        public string Path { get; set; }
//        public long Size { get; set; }
//        public DateTime UploadTime { get; set; }
//    }
//}



//using ExcelToPdfConverter.Services;
//using Microsoft.AspNetCore.Http;

//namespace ExcelToPdfConverter.Models
//{
//    public class ExcelUploadModel
//    {
//        public IFormFile? ExcelFile { get; set; }
//    }


//    public class PdfPreviewWithFitToPageRequest
//{
//    public string SessionId { get; set; }
//    public List<string> SelectedSheets { get; set; }
//    public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
//    public bool IncludeMergedPdfs { get; set; }
//}
//    public class WorksheetPreview
//    {
//        public string Name { get; set; } = string.Empty;
//        public int Index { get; set; }
//        public int TotalRows { get; set; }
//        public int TotalColumns { get; set; }
//        public List<List<CellPreview>> Cells { get; set; } = new List<List<CellPreview>>();
//        public List<ImagePreview> Images { get; set; } = new List<ImagePreview>();
//        public List<NameError> NameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
//        public bool HasNameErrors => NameErrors.Count > 0;
//        public bool HasInvoiceDates => InvoiceDates.Count > 0;
//    }

//    public class CellPreview
//    {
//        public int Row { get; set; }
//        public int Column { get; set; }
//        public string Value { get; set; } = string.Empty;
//        public string BackgroundColor { get; set; } = string.Empty;
//        public string TextColor { get; set; } = string.Empty;
//        public bool IsBold { get; set; }
//        public bool IsItalic { get; set; }
//        public bool Underline { get; set; }
//        public double FontSize { get; set; }
//        public string HorizontalAlignment { get; set; } = string.Empty;
//        public int ColSpan { get; set; } = 1;
//        public int RowSpan { get; set; } = 1;
//        public bool IsNameError => Value?.Contains("#NAME?") == true;
//        public string ColumnName => GetColumnName(Column);

//        private string GetColumnName(int column)
//        {
//            string columnName = "";
//            while (column > 0)
//            {
//                int modulo = (column - 1) % 26;
//                columnName = Convert.ToChar('A' + modulo) + columnName;
//                column = (column - modulo) / 26;
//            }
//            return columnName;
//        }
//    }

//    public class NameError
//    {
//        public int Row { get; set; }
//        public int Column { get; set; }
//        public string ColumnName { get; set; } = string.Empty;
//        public string Location => $"{ColumnName}{Row}";
//        public string FullLocation => $"Row {Row}, Column {ColumnName}";
//        public string SheetName { get; set; } = string.Empty;
//    }

//    public class InvoiceDate
//    {
//        public string SheetName { get; set; } = string.Empty;
//        public int Row { get; set; }
//        public int Column { get; set; }
//        public string ColumnName { get; set; } = string.Empty;
//        public string InvoiceDateText { get; set; } = string.Empty;
//        public string DateValue { get; set; } = string.Empty;
//        public string Location => $"{ColumnName}{Row}";
//        public string FullInfo => $"'{InvoiceDateText}' → '{DateValue}' at {Location}";
//    }

//    public class ImagePreview
//    {
//        public string Name { get; set; } = string.Empty;
//        public string Base64Data { get; set; } = string.Empty;
//        public string Format { get; set; } = string.Empty;
//        public int Row { get; set; }
//        public int Column { get; set; }
//    }

//    public class ConversionResult
//    {
//        public bool Success { get; set; }
//        public string Message { get; set; } = string.Empty;
//        public string PdfFilePath { get; set; } = string.Empty;
//        public string FileName { get; set; } = string.Empty;
//        public int TotalPages { get; set; }
//    }

//    public class ValidationResult
//    {
//        public string FileName { get; set; } = string.Empty;
//        public List<NameError> NameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
//        public bool HasNameErrors => NameErrors.Count > 0;
//        public bool HasInvoiceDates => InvoiceDates.Count > 0;
//    }

//    public class PreviewModel
//    {
//        public string OriginalFileName { get; set; } = string.Empty;
//        public List<WorksheetPreview> Worksheets { get; set; } = new List<WorksheetPreview>();
//        public string SessionId { get; set; } = string.Empty;
//        public List<NameError> AllNameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> AllInvoiceDates { get; set; } = new List<InvoiceDate>();
//        public List<FileSelection> FileSelections { get; set; } = new List<FileSelection>();
//        public Dictionary<string, string> SuggestedOrientations { get; set; } = new Dictionary<string, string>();
//        public Dictionary<string, ExcelPreviewService.SheetOrientationInfo> SheetOrientationAnalysis { get; set; }
//            = new Dictionary<string, ExcelPreviewService.SheetOrientationInfo>();
//        public bool HasNameErrors => AllNameErrors.Count > 0;
//        public bool HasInvoiceDates => AllInvoiceDates.Count > 0;
//    }

//    public class FileSelection
//    {
//        public string FileName { get; set; } = string.Empty;
//        public string SheetName { get; set; } = string.Empty;
//        public bool IsSelected { get; set; } = true;
//        public int SortOrder { get; set; }
//        public bool HasNameErrors { get; set; }
//        public bool HasInvoiceDates { get; set; }
//        public List<NameError> NameErrors { get; set; } = new List<NameError>();
//        public List<InvoiceDate> InvoiceDates { get; set; } = new List<InvoiceDate>();
//        public string Orientation { get; set; } = "Landscape";
//    }

//    public class FileNamesModel
//    {
//        public string ExcelFileName { get; set; } = string.Empty;
//        public List<string> PdfFileNames { get; set; } = new List<string>();
//        public int TotalPdfFiles { get; set; }
//    }

//    public class PdfReorderingRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfo> PageOrderData { get; set; }
//        public Dictionary<int, int> RotationData { get; set; }
//    }

//    public class PdfConversionResponse
//    {
//        public bool Success { get; set; }
//        public string Message { get; set; }
//        public string FileName { get; set; }
//        public byte[] FileData { get; set; }
//    }

//    public class PdfPreviewRequest
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//    }

//    public class PdfPreviewResponse
//    {
//        public bool Success { get; set; }
//        public string Message { get; set; }
//        public string PdfData { get; set; }
//        public string FileName { get; set; }
//        public string GeneratedTime { get; set; }
//        public int TotalPages { get; set; }
//    }

//    // Existing models के साथ ये नए models जोड़ें
//    public class MarginInfo
//    {
//        public float Top { get; set; } = 10;
//        public float Bottom { get; set; } = 10;
//        public float Left { get; set; } = 10;
//        public float Right { get; set; } = 10;
//    }

//    public class PositionInfo
//    {
//        public string Horizontal { get; set; } = "center";
//        public string Vertical { get; set; } = "middle";
//    }

//    public class PageOrderInfo
//    {
//        public int OriginalPage { get; set; }
//        public int CurrentOrder { get; set; }
//        public bool Visible { get; set; }
//        public int Rotation { get; set; }
//        public string Orientation { get; set; } = "portrait";
//        public string FitMode { get; set; } = "fitToPage";
//        public MarginInfo Margin { get; set; } = new MarginInfo();
//        public PositionInfo Position { get; set; } = new PositionInfo();
//    }

//    public class PdfReorderingRequestWithAllSettings
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfo> PageOrderData { get; set; }
//        public Dictionary<int, int> RotationData { get; set; }
//        public Dictionary<int, string> OrientationData { get; set; }
//        public Dictionary<int, string> FitModeData { get; set; }
//        public Dictionary<int, MarginInfo> MarginData { get; set; }
//        public Dictionary<int, PositionInfo> PositionData { get; set; }
//    }

//    public class PageOrderInfoWithRotation
//    {
//        public int OriginalPage { get; set; }
//        public int CurrentOrder { get; set; }
//        public bool Visible { get; set; }
//        public string Orientation { get; set; } = "portrait";
//        public int Rotation { get; set; } = 0;
//    }

//    public class PdfRequestWithRotation
//    {
//        public string SessionId { get; set; }
//        public List<string> SelectedSheets { get; set; }
//        public List<PageOrderInfoWithRotation> PageOrderData { get; set; }
//        public Dictionary<int, string> OrientationData { get; set; }
//        public Dictionary<int, int> RotationData { get; set; }
//    }

//    public class RemovePdfRequest
//    {
//        public string SessionId { get; set; }
//        public string FileName { get; set; }
//    }

//    public class MergePdfRequest
//    {
//        public string SessionId { get; set; }
//    }

//    public class DownloadPdfRequest
//    {
//        public string SessionId { get; set; }
//        public string FileName { get; set; }
//    }

//    public class MergedPdfInfo
//    {
//        public string FileName { get; set; }
//        public string FilePath { get; set; }
//        public long FileSize { get; set; }
//        public int TotalPages { get; set; }
//        public DateTime CreatedAt { get; set; }
//    }


//}
