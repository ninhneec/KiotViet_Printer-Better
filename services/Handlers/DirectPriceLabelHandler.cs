using System.Globalization;
using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services.Interfaces;

namespace KiotVietLabelPrinter.Services.Handlers;

public class DirectPriceLabelHandler : ILabelHandler
{
    private readonly BarTenderService barTenderService = new();
    private static readonly CultureInfo VietnameseCulture = CultureInfo.GetCultureInfo("vi-VN");

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

        foreach (ProductRow item in products)
        {
            string displayName = string.IsNullOrWhiteSpace(item.ProductNameWithAttr)
                ? item.ProductName
                : item.ProductNameWithAttr;

            Dictionary<string, string> fields = new()
            {
                ["Tên hàng"] = displayName,
                ["Giá bán"] = item.Price.ToString("N0", VietnameseCulture),
                ["Đơn vị tính"] = item.Unit
            };

            int copies = Math.Max(1, (int)Math.Round(item.Quantity));
            for (int copy = 0; copy < copies; copy++)
                barTenderService.Print(label.TemplatePath, fields);
        }
    }
}
