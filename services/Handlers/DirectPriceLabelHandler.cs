using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services.Interfaces;

namespace KiotVietLabelPrinter.Services.Handlers;

public class DirectPriceLabelHandler : ILabelHandler
{
    private readonly BarTenderService barTenderService = new();
    private readonly ExcelService excelService = new();

    public string HandlerType => "DIRECT_PRICE";

    public List<PreviewRow> BuildPreview(
        List<ProductRow> products,
        LabelDefinition label,
        string employeeCode)
    {
        return products.Select(item => new PreviewRow
        {
            ProductCode = item.ProductCode,
            ProductName = item.ProductNameWithAttr.Length > 0
                ? item.ProductNameWithAttr
                : item.ProductName,
            ProductNameWithAttr = item.ProductNameWithAttr,
            Unit = item.Unit,
            Quantity = item.Quantity,
            Price = item.Price
        }).ToList();
    }

    public void PrepareDataAndPrint(
        List<ProductRow> products,
        LabelDefinition label,
        string employeeCode)
    {
        if (string.IsNullOrWhiteSpace(label.TemplatePath) || !File.Exists(label.TemplatePath))
            throw new Exception($"Không tìm thấy file BarTender:\n{label.TemplatePath}");
        if (string.IsNullOrWhiteSpace(label.DataFilePath))
            throw new Exception("Mẫu chưa có file data trung gian.");

        // Ghi toàn bộ sản phẩm đã chọn vào file trung gian một lần.
        // BarTender sẽ đọc tất cả record trong database và in trọn lô.
        excelService.WriteDirectPriceDataFile(label.DataFilePath, products);
        barTenderService.Print(label.TemplatePath);
    }
}
