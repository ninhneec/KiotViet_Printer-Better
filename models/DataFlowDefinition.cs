namespace KiotVietLabelPrinter.Models;

public enum DataFlowNodeType
{
    ExcelSource,
    Join,
    Filter,
    SelectColumns,
    Validate,
    LabelOutput
}

public class DataFlowDefinition
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = "Flow tem mới";
    public string Description { get; set; } = "";
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
    public List<DataFlowNode> Nodes { get; set; } = new();
    public List<DataFlowConnection> Connections { get; set; } = new();
}

public class DataFlowNode
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public DataFlowNodeType Type { get; set; }
    public string Title { get; set; } = "";
    public int X { get; set; }
    public int Y { get; set; }
    public Dictionary<string, string> Settings { get; set; } = new();
}

public class DataFlowConnection
{
    public string FromNodeId { get; set; } = "";
    public string ToNodeId { get; set; } = "";
    public int ToInput { get; set; }
}

public class FlowTable
{
    public List<string> Columns { get; set; } = new();
    public List<Dictionary<string, string>> Rows { get; set; } = new();
}

