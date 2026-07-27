using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using KiotVietLabelPrinter.Models;
using System.Globalization;
using System.Text;

namespace KiotVietLabelPrinter.Services;

public class ExcelService
{
    private const int BarcodeColumnIndex = 5; // Cột F

    #region PUBLIC API CHO PROJECT MỚI

    public List<ProductRow> ReadProducts(string sourceFile)
    {
        using IWorkbook workbook = OpenWorkbook(sourceFile);
        ISheet sheet = workbook.GetSheetAt(0);

        List<ProductRow> products = new();
        IRow? header = sheet.GetRow(sheet.FirstRowNum);

        if (header == null)
            throw new Exception("File Excel không có hàng tiêu đề.");

        Dictionary<string, int> columns = BuildColumnMap(header);
        int productCodeColumn = RequireColumn(columns, "Mã hàng");
        int barcodeColumn = FindColumn(columns, "Mã vạch");
        int productNameColumn = RequireColumn(columns, "Tên hàng");
        int productNameWithAttrColumn = FindColumn(columns, "Tên hàng thuộc tính");
        int unitColumn = FindColumn(columns, "Đơn vị tính");
        int quantityColumn = FindColumn(columns, "Số lượng");
        int priceColumn = FindColumn(columns, "Giá bán");
        int descriptionColumn = FindColumn(columns, "Mô tả");

        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
        {
            IRow? row = sheet.GetRow(i);
            if (row == null)
                continue;

            string productCode = GetCellString(row, productCodeColumn);
            string barcode = GetCellString(row, barcodeColumn);
            string productName = GetCellString(row, productNameColumn);
            string productNameWithAttr = GetCellString(row, productNameWithAttrColumn);
            string unit = GetCellString(row, unitColumn);
            double quantity = GetCellDouble(row, quantityColumn);
            double price = GetCellDouble(row, priceColumn);
            string description = GetCellString(row, descriptionColumn);

            if (string.IsNullOrWhiteSpace(productCode) &&
                string.IsNullOrWhiteSpace(productName) &&
                string.IsNullOrWhiteSpace(productNameWithAttr))
            {
                continue;
            }

            products.Add(new ProductRow
            {
                ProductCode = productCode,
                Barcode = barcode,
                ProductName = productName,
                ProductNameWithAttr = productNameWithAttr,
                Unit = unit,
                Quantity = quantity <= 0 ? 1 : quantity,
                Price = price,
                Description = description
            });
        }

        return products;
    }

    private static Dictionary<string, int> BuildColumnMap(IRow header)
    {
        Dictionary<string, int> result = new(StringComparer.OrdinalIgnoreCase);

        for (int index = header.FirstCellNum; index < header.LastCellNum; index++)
        {
            if (index < 0) continue;
            string name = NormalizeHeader(GetCellString(header, index));
            if (!string.IsNullOrWhiteSpace(name) && !result.ContainsKey(name))
                result[name] = index;
        }

        return result;
    }

    private static int RequireColumn(Dictionary<string, int> columns, string name)
    {
        int index = FindColumn(columns, name);
        if (index >= 0) return index;

        throw new Exception(
            $"File Excel thiếu cột “{name}”.\n\n" +
            "Hãy xuất lại file Bảng giá sản phẩm từ KiotViet.");
    }

    private static int FindColumn(Dictionary<string, int> columns, string name)
    {
        return columns.TryGetValue(NormalizeHeader(name), out int index) ? index : -1;
    }

    private static string NormalizeHeader(string value)
    {
        string decomposed = (value ?? string.Empty)
            .Trim()
            .ToLowerInvariant()
            .Normalize(NormalizationForm.FormD);

        StringBuilder result = new();
        foreach (char character in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
                result.Append(character == 'đ' ? 'd' : character);
        }

        return string.Join(" ", result.ToString()
            .Normalize(NormalizationForm.FormC)
            .Replace("(", " ")
            .Replace(")", " ")
            .Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    /// <summary>
    /// Ghi file data tem FULL theo logic tool cũ:
    /// copy nguyên dữ liệu từ source sang target, không parse cột F
    /// </summary>
    public void WriteGenericLabelData(string sourceFile, string targetFile)
    {
        CopyToBarTenderData(sourceFile, targetFile, false, "");
    }

    /// <summary>
    /// Ghi file data tem BARCODE theo logic tool cũ:
    /// copy nguyên dữ liệu từ source sang target,
    /// riêng cột F parse mã + nối mã nhân viên
    /// </summary>
    public void WriteBarcodeLikeData(string sourceFile, string targetFile, string employeeCode)
    {
        CopyToBarTenderData(sourceFile, targetFile, true, employeeCode);
    }

    public void ExportProducts(string targetFile, IEnumerable<ProductRow> products)
    {
        using IWorkbook workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("Dữ liệu in tem");
        string[] headers =
        [
            "Mã hàng",
            "Tên hàng",
            "Tên in trên tem",
            "Đơn vị tính",
            "Số lượng",
            "Giá bán"
        ];

        IRow header = sheet.CreateRow(0);
        for (int column = 0; column < headers.Length; column++)
            header.CreateCell(column).SetCellValue(headers[column]);

        int rowIndex = 1;
        foreach (ProductRow product in products)
        {
            IRow row = sheet.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue(product.ProductCode);
            row.CreateCell(1).SetCellValue(product.ProductName);
            row.CreateCell(2).SetCellValue(product.ProductNameWithAttr);
            row.CreateCell(3).SetCellValue(product.Unit);
            row.CreateCell(4).SetCellValue(product.Quantity);
            row.CreateCell(5).SetCellValue(product.Price);
        }

        for (int column = 0; column < headers.Length; column++)
            sheet.AutoSizeColumn(column);

        SaveWorkbook(workbook, targetFile);
    }

    #endregion

    #region CORE LOGIC - GIỮ THEO TOOL CŨ

    public void CopyToBarTenderData(
        string sourceFile,
        string targetFile,
        bool isBarcode,
        string employeeCode = "")
    {
        using IWorkbook sourceWorkbook = OpenWorkbook(sourceFile);
        using IWorkbook targetWorkbook = OpenWorkbook(targetFile);

        ISheet sourceSheet = sourceWorkbook.GetSheetAt(0);
        ISheet targetSheet = targetWorkbook.GetSheetAt(0);

        // Xóa dữ liệu cũ, giữ header
        ClearSheetData(targetSheet);

        // Copy nguyên từng hàng/cột từ source sang target
        for (int i = 1; i <= sourceSheet.LastRowNum; i++)
        {
            IRow? sourceRow = sourceSheet.GetRow(i);
            if (sourceRow == null)
                continue;

            // bỏ qua dòng hoàn toàn rỗng
            if (IsRowEmpty(sourceRow))
                continue;

            IRow targetRow = targetSheet.GetRow(i) ?? targetSheet.CreateRow(i);

            for (int j = 0; j < sourceRow.LastCellNum; j++)
            {
                ICell? sourceCell = sourceRow.GetCell(j);
                if (sourceCell == null)
                    continue;

                ICell targetCell = targetRow.GetCell(j) ?? targetRow.CreateCell(j);

                string value = GetCellText(sourceCell);

                // TEM BARCODE: chỉ xử lý riêng cột F
                if (isBarcode && j == BarcodeColumnIndex)
                {
                    string parsedCode = BarcodeParser.Parse(value);

                    if (string.IsNullOrWhiteSpace(parsedCode))
                    {
                        // fallback về mã hàng cột C
                        parsedCode = sourceRow.GetCell(2)?.ToString()?.Trim() ?? "";
                    }

                    if (!string.IsNullOrWhiteSpace(employeeCode))
                    {
                        parsedCode = $"{parsedCode}-{employeeCode.Trim()}";
                    }

                    value = parsedCode;
                    targetCell.SetCellValue(value);
                    continue;
                }

                CopyCellValue(sourceCell, targetCell);
            }
        }

        SaveWorkbook(targetWorkbook, targetFile);
    }

    #endregion

    #region HELPERS

    private static IWorkbook OpenWorkbook(string filePath)
    {
        if (!File.Exists(filePath))
            throw new Exception($"Không tìm thấy file Excel:\n{filePath}");

        FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        try
        {
            // Tự nhận diện xls / xlsx theo nội dung thật của file
            return WorkbookFactory.Create(fs);
        }
        catch (Exception ex)
        {
            fs.Dispose();
            throw new Exception(
                $"Không đọc được file Excel.\n" +
                $"File: {Path.GetFileName(filePath)}\n" +
                $"Đường dẫn: {filePath}\n" +
                $"Chi tiết: {ex.Message}");
        }
    }

    private static void SaveWorkbook(IWorkbook workbook, string targetFile)
    {
        string? folder = Path.GetDirectoryName(targetFile);
        if (!string.IsNullOrWhiteSpace(folder) && !Directory.Exists(folder))
        {
            Directory.CreateDirectory(folder);
        }

        using FileStream output = new(targetFile, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
        workbook.Write(output);
    }

    private static void ClearSheetData(ISheet sheet)
    {
        for (int i = sheet.LastRowNum; i >= 1; i--)
        {
            IRow? row = sheet.GetRow(i);
            if (row != null)
                sheet.RemoveRow(row);
        }
    }

    private static bool IsRowEmpty(IRow row)
    {
        for (int i = row.FirstCellNum; i < row.LastCellNum; i++)
        {
            if (i < 0) continue;

            ICell? cell = row.GetCell(i);
            if (cell != null && !string.IsNullOrWhiteSpace(GetCellText(cell)))
                return false;
        }

        return true;
    }

    private static void CopyCellValue(ICell sourceCell, ICell targetCell)
    {
        switch (sourceCell.CellType)
        {
            case CellType.Numeric:
                targetCell.SetCellValue(sourceCell.NumericCellValue);
                break;

            case CellType.Boolean:
                targetCell.SetCellValue(sourceCell.BooleanCellValue);
                break;

            case CellType.Formula:
                // Với file data BarTender, giữ nguyên kết quả text an toàn hơn
                targetCell.SetCellValue(GetCellText(sourceCell));
                break;

            case CellType.Blank:
                targetCell.SetCellValue(string.Empty);
                break;

            default:
                targetCell.SetCellValue(GetCellText(sourceCell));
                break;
        }
    }

    private static string GetCellString(IRow row, int index)
    {
        if (index < 0) return "";
        return row.GetCell(index)?.ToString()?.Trim() ?? "";
    }

    private static double GetCellDouble(IRow row, int index)
    {
        if (index < 0) return 0;
        ICell? cell = row.GetCell(index);
        if (cell == null) return 0;

        if (cell.CellType == CellType.Numeric)
            return cell.NumericCellValue;

        if (double.TryParse(cell.ToString(), out double value))
            return value;

        return 0;
    }

    private static string GetCellText(ICell cell)
    {
        return cell.ToString()?.Trim() ?? "";
    }

    #endregion
}
