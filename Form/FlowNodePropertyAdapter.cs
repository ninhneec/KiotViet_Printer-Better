using KiotVietLabelPrinter.Models;
using System.ComponentModel;

namespace KiotVietLabelPrinter.Forms;

public class FlowNodePropertyAdapter
{
    private readonly DataFlowNode node;
    public FlowNodePropertyAdapter(DataFlowNode node) => this.node = node;

    [Category("Khối"), DisplayName("Tên khối")]
    public string Title { get => node.Title; set => node.Title = value; }

    [Category("Nguồn Excel"), DisplayName("Đường dẫn file")]
    public string FilePath
    {
        get => node.Settings.GetValueOrDefault("FilePath", "");
        set => node.Settings["FilePath"] = value ?? "";
    }

    [Category("Nối dữ liệu"), DisplayName("Cột nối file 1")]
    public string LeftKey
    {
        get => node.Settings.GetValueOrDefault("LeftKey", "Mã hàng");
        set => node.Settings["LeftKey"] = value ?? "";
    }

    [Category("Nối dữ liệu"), DisplayName("Cột nối file 2")]
    public string RightKey
    {
        get => node.Settings.GetValueOrDefault("RightKey", "Mã hàng");
        set => node.Settings["RightKey"] = value ?? "";
    }

    [Category("Nối dữ liệu"), DisplayName("Cách nối"), TypeConverter(typeof(JoinTypeConverter))]
    public string JoinType
    {
        get => node.Settings.GetValueOrDefault("JoinType", "Left");
        set => node.Settings["JoinType"] = value ?? "Left";
    }

    [Category("Lọc dữ liệu"), DisplayName("Tên cột")]
    public string FilterColumn
    {
        get => node.Settings.GetValueOrDefault("Column", "");
        set => node.Settings["Column"] = value ?? "";
    }

    [Category("Lọc dữ liệu"), DisplayName("Điều kiện"), TypeConverter(typeof(FilterOperationConverter))]
    public string FilterOperation
    {
        get => node.Settings.GetValueOrDefault("Operation", "Có dữ liệu");
        set => node.Settings["Operation"] = value ?? "";
    }

    [Category("Lọc dữ liệu"), DisplayName("Giá trị so sánh")]
    public string FilterValue
    {
        get => node.Settings.GetValueOrDefault("Value", "");
        set => node.Settings["Value"] = value ?? "";
    }

    [Category("Chọn cột"), DisplayName("Ánh xạ cột")]
    [Description("Dạng: Tên đầu ra=Tên cột nguồn;Giá bán=Giá bán")]
    public string Mappings
    {
        get => node.Settings.GetValueOrDefault("Mappings", "");
        set => node.Settings["Mappings"] = value ?? "";
    }

    [Category("Kiểm tra"), DisplayName("Các cột bắt buộc")]
    [Description("Ngăn cách bằng dấu chấm phẩy.")]
    public string RequiredColumns
    {
        get => node.Settings.GetValueOrDefault("Required", "");
        set => node.Settings["Required"] = value ?? "";
    }

    [Browsable(false)]
    public DataFlowNodeType NodeType => node.Type;
}

public class JoinTypeConverter : StringConverter
{
    public override bool GetStandardValuesSupported(ITypeDescriptorContext? context) => true;
    public override bool GetStandardValuesExclusive(ITypeDescriptorContext? context) => true;
    public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext? context) =>
        new(new[] { "Left", "Inner" });
}

public class FilterOperationConverter : StringConverter
{
    public override bool GetStandardValuesSupported(ITypeDescriptorContext? context) => true;
    public override bool GetStandardValuesExclusive(ITypeDescriptorContext? context) => true;
    public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext? context) =>
        new(new[] { "Có dữ liệu", "Không có dữ liệu", "Bằng", "Chứa" });
}
