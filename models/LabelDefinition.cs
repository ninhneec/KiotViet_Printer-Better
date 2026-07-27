namespace KiotVietLabelPrinter.Models;

public class LabelDefinition
{
    public string SourceExcelFile { get; set; } = "";
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string IconText { get; set; } = "";

    public bool IsEnabled { get; set; } = true;

    public string TemplatePath { get; set; } = "";
    public string DataFilePath { get; set; } = "";

    // Tên handler xử lý tem này
    // Ví dụ: FULL, BARCODE, GENERIC
    public string HandlerType { get; set; } = "GENERIC";

    // Có cần nhập mã nhân viên không
    public bool RequiresEmployeeCode { get; set; }

    // Có parse mã từ "Tên hàng (thuộc tính)" hay không
    public bool UseBarcodeParser { get; set; }

    // Có nối mã nhân viên vào cuối mã hay không
    public bool AppendEmployeeCode { get; set; }

    // Cột đích trong file data để ghi mã đã parse
    // ví dụ cột F = 5
    public int TargetNameColumnIndex { get; set; } = 5;

    [Newtonsoft.Json.JsonIgnore]
    public string DisplayName =>
        $"{(IsEnabled ? "●" : "○")}  {(string.IsNullOrWhiteSpace(Name) ? "Mẫu chưa đặt tên" : Name)}";

    public List<string> GetReadinessIssues()
    {
        List<string> issues = new();

        if (string.IsNullOrWhiteSpace(Name))
            issues.Add("chưa có tên");

        if (string.IsNullOrWhiteSpace(Code))
            issues.Add("chưa có mã");

        if (string.IsNullOrWhiteSpace(TemplatePath) || !File.Exists(TemplatePath))
            issues.Add("file .btw không tồn tại");

        if (!HandlerType.Equals("DIRECT_PRICE", StringComparison.OrdinalIgnoreCase) &&
            (string.IsNullOrWhiteSpace(DataFilePath) || !File.Exists(DataFilePath)))
            issues.Add("file dữ liệu không tồn tại");

        return issues;
    }
}
