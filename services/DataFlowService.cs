using KiotVietLabelPrinter.Models;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using System.Globalization;
using System.Text;

namespace KiotVietLabelPrinter.Services;

public class DataFlowService
{
    private readonly string flowDirectory;

    public DataFlowService()
    {
        flowDirectory = Path.Combine(ConfigService.Instance.DataDirectory, "Flows");
        Directory.CreateDirectory(flowDirectory);
    }

    public IReadOnlyList<DataFlowDefinition> LoadAll()
    {
        return Directory.GetFiles(flowDirectory, "*.json")
            .Select(Load)
            .Where(flow => flow != null)
            .Cast<DataFlowDefinition>()
            .OrderBy(flow => flow.Name)
            .ToList();
    }

    public DataFlowDefinition? Load(string path)
    {
        try
        {
            return JsonConvert.DeserializeObject<DataFlowDefinition>(File.ReadAllText(path));
        }
        catch
        {
            return null;
        }
    }

    public string Save(DataFlowDefinition flow)
    {
        flow.UpdatedAt = DateTime.Now;
        string path = Path.Combine(flowDirectory, $"{SafeName(flow.Id)}.json");
        File.WriteAllText(path, JsonConvert.SerializeObject(flow, Formatting.Indented));
        return path;
    }

    public void Delete(DataFlowDefinition flow)
    {
        string path = Path.Combine(flowDirectory, $"{SafeName(flow.Id)}.json");
        if (File.Exists(path))
            File.Delete(path);
    }

    public FlowTable Execute(DataFlowDefinition flow)
    {
        ValidateGraph(flow);
        Dictionary<string, FlowTable> results = new();
        foreach (DataFlowNode node in TopologicalOrder(flow))
        {
            List<DataFlowConnection> incoming = flow.Connections
                .Where(connection => connection.ToNodeId == node.Id)
                .OrderBy(connection => connection.ToInput)
                .ToList();
            List<FlowTable> inputs = incoming.Select(connection => results[connection.FromNodeId]).ToList();
            results[node.Id] = ExecuteNode(node, inputs);
        }

        DataFlowNode output = flow.Nodes.FirstOrDefault(node => node.Type == DataFlowNodeType.LabelOutput)
            ?? TopologicalOrder(flow).Last();
        return results[output.Id];
    }

    public FlowTable ReadSource(string path) => ReadExcel(path);

    public DataFlowDefinition CreateSimpleFlow(
        string name,
        string file1,
        string? file2,
        string? leftKey,
        string? rightKey,
        string nameColumn,
        string priceColumn,
        string unitColumn)
    {
        DataFlowNode source1 = NewNode(DataFlowNodeType.ExcelSource, "File sản phẩm", 60, 180);
        source1.Settings["FilePath"] = file1;
        DataFlowNode select = NewNode(DataFlowNodeType.SelectColumns, "Lấy 3 cột cần in", 650, 180);
        select.Settings["Mappings"] =
            $"Tên hàng={nameColumn};Giá bán={priceColumn};Đơn vị tính={unitColumn}";
        DataFlowNode validate = NewNode(DataFlowNodeType.Validate, "Kiểm tra dữ liệu", 920, 180);
        validate.Settings["Required"] = "Tên hàng;Giá bán;Đơn vị tính";
        DataFlowNode output = NewNode(DataFlowNodeType.LabelOutput, "Sẵn sàng cho BarTender", 1190, 180);
        DataFlowDefinition flow = new()
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Flow tem mới" : name.Trim(),
            Nodes = [source1, select, validate, output]
        };

        if (!string.IsNullOrWhiteSpace(file2))
        {
            DataFlowNode source2 = NewNode(DataFlowNodeType.ExcelSource, "File bổ sung", 60, 370);
            source2.Settings["FilePath"] = file2;
            DataFlowNode join = NewNode(DataFlowNodeType.Join, "Tự nối hai file", 370, 260);
            join.Settings["LeftKey"] = leftKey ?? "";
            join.Settings["RightKey"] = rightKey ?? "";
            join.Settings["JoinType"] = "Left";
            flow.Nodes.Insert(1, source2);
            flow.Nodes.Insert(2, join);
            flow.Connections.AddRange(
            [
                new() { FromNodeId = source1.Id, ToNodeId = join.Id, ToInput = 0 },
                new() { FromNodeId = source2.Id, ToNodeId = join.Id, ToInput = 1 },
                new() { FromNodeId = join.Id, ToNodeId = select.Id }
            ]);
        }
        else
        {
            flow.Connections.Add(new()
            {
                FromNodeId = source1.Id,
                ToNodeId = select.Id
            });
        }
        flow.Connections.AddRange(
        [
            new() { FromNodeId = select.Id, ToNodeId = validate.Id },
            new() { FromNodeId = validate.Id, ToNodeId = output.Id }
        ]);
        return flow;
    }

    public static DataFlowDefinition CreateStarterFlow()
    {
        DataFlowNode sourceA = NewNode(DataFlowNodeType.ExcelSource, "File dữ liệu chính", 50, 100);
        DataFlowNode sourceB = NewNode(DataFlowNodeType.ExcelSource, "File bổ sung", 50, 330);
        DataFlowNode join = NewNode(DataFlowNodeType.Join, "Nối theo mã hàng", 330, 210);
        join.Settings["LeftKey"] = "Mã hàng";
        join.Settings["RightKey"] = "Mã hàng";
        join.Settings["JoinType"] = "Left";
        DataFlowNode select = NewNode(DataFlowNodeType.SelectColumns, "Chọn cột cho tem", 610, 210);
        select.Settings["Mappings"] = "Tên hàng=Tên hàng;Giá bán=Giá bán;Đơn vị tính=Đơn vị tính";
        DataFlowNode validate = NewNode(DataFlowNodeType.Validate, "Kiểm tra dữ liệu", 890, 210);
        validate.Settings["Required"] = "Tên hàng;Giá bán;Đơn vị tính";
        DataFlowNode output = NewNode(DataFlowNodeType.LabelOutput, "Mẫu tem BarTender", 1170, 210);

        return new DataFlowDefinition
        {
            Name = "Ghép dữ liệu cho tem",
            Description = "Đọc hai file, nối theo mã hàng và tạo dữ liệu BarTender.",
            Nodes = [sourceA, sourceB, join, select, validate, output],
            Connections =
            [
                new() { FromNodeId = sourceA.Id, ToNodeId = join.Id, ToInput = 0 },
                new() { FromNodeId = sourceB.Id, ToNodeId = join.Id, ToInput = 1 },
                new() { FromNodeId = join.Id, ToNodeId = select.Id },
                new() { FromNodeId = select.Id, ToNodeId = validate.Id },
                new() { FromNodeId = validate.Id, ToNodeId = output.Id }
            ]
        };
    }

    private static DataFlowNode NewNode(DataFlowNodeType type, string title, int x, int y) =>
        new() { Type = type, Title = title, X = x, Y = y };

    private static FlowTable ExecuteNode(DataFlowNode node, List<FlowTable> inputs)
    {
        return node.Type switch
        {
            DataFlowNodeType.ExcelSource => ReadExcel(node.Settings.GetValueOrDefault("FilePath", "")),
            DataFlowNodeType.Join => Join(inputs, node),
            DataFlowNodeType.Filter => Filter(RequireInput(inputs, node), node),
            DataFlowNodeType.SelectColumns => SelectColumns(RequireInput(inputs, node), node),
            DataFlowNodeType.Validate => ValidateRows(RequireInput(inputs, node), node),
            DataFlowNodeType.LabelOutput => RequireInput(inputs, node),
            _ => throw new InvalidOperationException($"Chưa hỗ trợ node {node.Type}.")
        };
    }

    private static FlowTable ReadExcel(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            throw new InvalidOperationException("Node nguồn chưa chọn file Excel hợp lệ.");
        if (Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            return ReadCsv(path);

        using FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using IWorkbook workbook = WorkbookFactory.Create(stream);
        ISheet sheet = workbook.GetSheetAt(0);
        IRow header = sheet.GetRow(sheet.FirstRowNum)
            ?? throw new InvalidOperationException("File Excel không có hàng tiêu đề.");
        List<string> columns = new();
        for (int index = Math.Max(0, (int)header.FirstCellNum); index < header.LastCellNum; index++)
            columns.Add(header.GetCell(index)?.ToString()?.Trim() ?? $"Cột {index + 1}");

        FlowTable table = new() { Columns = columns };
        for (int rowIndex = sheet.FirstRowNum + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            IRow? row = sheet.GetRow(rowIndex);
            if (row == null) continue;
            Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase);
            bool hasValue = false;
            for (int column = 0; column < columns.Count; column++)
            {
                string value = row.GetCell(column)?.ToString()?.Trim() ?? "";
                values[columns[column]] = value;
                hasValue |= !string.IsNullOrWhiteSpace(value);
            }
            if (hasValue) table.Rows.Add(values);
        }
        return table;
    }

    private static FlowTable ReadCsv(string path)
    {
        string[] lines = File.ReadAllLines(path);
        if (lines.Length == 0)
            throw new InvalidOperationException("File CSV không có dữ liệu.");
        char separator = DetectSeparator(lines[0]);
        List<string> columns = ParseCsvLine(lines[0], separator)
            .Select((value, index) => string.IsNullOrWhiteSpace(value) ? $"Cột {index + 1}" : value.Trim())
            .ToList();
        FlowTable table = new() { Columns = columns };
        foreach (string line in lines.Skip(1))
        {
            List<string> cells = ParseCsvLine(line, separator);
            Dictionary<string, string> row = new(StringComparer.OrdinalIgnoreCase);
            bool hasValue = false;
            for (int index = 0; index < columns.Count; index++)
            {
                string value = index < cells.Count ? cells[index].Trim() : "";
                row[columns[index]] = value;
                hasValue |= !string.IsNullOrWhiteSpace(value);
            }
            if (hasValue) table.Rows.Add(row);
        }
        return table;
    }

    private static char DetectSeparator(string header)
    {
        char[] candidates = [',', ';', '\t'];
        return candidates.OrderByDescending(separator => ParseCsvLine(header, separator).Count).First();
    }

    private static List<string> ParseCsvLine(string line, char separator)
    {
        List<string> result = new();
        StringBuilder cell = new();
        bool quoted = false;
        for (int index = 0; index < line.Length; index++)
        {
            char character = line[index];
            if (character == '"')
            {
                if (quoted && index + 1 < line.Length && line[index + 1] == '"')
                {
                    cell.Append('"');
                    index++;
                }
                else quoted = !quoted;
            }
            else if (character == separator && !quoted)
            {
                result.Add(cell.ToString());
                cell.Clear();
            }
            else cell.Append(character);
        }
        result.Add(cell.ToString());
        return result;
    }

    private static FlowTable Join(List<FlowTable> inputs, DataFlowNode node)
    {
        if (inputs.Count != 2)
            throw new InvalidOperationException("Node Nối dữ liệu cần đúng 2 dây đầu vào.");
        string leftKey = node.Settings.GetValueOrDefault("LeftKey", "");
        string rightKey = node.Settings.GetValueOrDefault("RightKey", "");
        if (!inputs[0].Columns.Contains(leftKey, StringComparer.OrdinalIgnoreCase) ||
            !inputs[1].Columns.Contains(rightKey, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Không tìm thấy cột nối “{leftKey}” hoặc “{rightKey}”.");

        FlowTable result = new();
        result.Columns.AddRange(inputs[0].Columns);
        foreach (string column in inputs[1].Columns)
            result.Columns.Add(UniqueColumn(result.Columns, column));

        Dictionary<string, List<Dictionary<string, string>>> rightIndex = inputs[1].Rows
            .GroupBy(row => Normalize(row.GetValueOrDefault(rightKey, "")))
            .ToDictionary(group => group.Key, group => group.ToList());
        bool inner = node.Settings.GetValueOrDefault("JoinType", "Left")
            .Equals("Inner", StringComparison.OrdinalIgnoreCase);

        foreach (Dictionary<string, string> leftRow in inputs[0].Rows)
        {
            string key = Normalize(leftRow.GetValueOrDefault(leftKey, ""));
            if (!rightIndex.TryGetValue(key, out List<Dictionary<string, string>>? matches))
            {
                if (!inner) result.Rows.Add(MergeRows(leftRow, null, inputs[1].Columns, result.Columns));
                continue;
            }
            foreach (Dictionary<string, string> rightRow in matches)
                result.Rows.Add(MergeRows(leftRow, rightRow, inputs[1].Columns, result.Columns));
        }
        return result;
    }

    private static FlowTable Filter(FlowTable input, DataFlowNode node)
    {
        string column = node.Settings.GetValueOrDefault("Column", "");
        if (!input.Columns.Contains(column, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Node Lọc không tìm thấy cột “{column}”.");
        string operation = node.Settings.GetValueOrDefault("Operation", "Có dữ liệu");
        string value = node.Settings.GetValueOrDefault("Value", "");
        FlowTable result = new() { Columns = [.. input.Columns] };
        result.Rows = input.Rows.Where(row =>
        {
            string cell = row.GetValueOrDefault(column, "");
            return operation switch
            {
                "Không có dữ liệu" => string.IsNullOrWhiteSpace(cell),
                "Bằng" => cell.Equals(value, StringComparison.OrdinalIgnoreCase),
                "Chứa" => cell.Contains(value, StringComparison.OrdinalIgnoreCase),
                _ => !string.IsNullOrWhiteSpace(cell)
            };
        }).Select(row => new Dictionary<string, string>(row, StringComparer.OrdinalIgnoreCase)).ToList();
        return result;
    }

    private static FlowTable SelectColumns(FlowTable input, DataFlowNode node)
    {
        List<(string Target, string Source)> mappings = node.Settings
            .GetValueOrDefault("Mappings", "")
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(item => item.Split('=', 2, StringSplitOptions.TrimEntries))
            .Where(parts => parts.Length == 2)
            .Select(parts => (parts[0], parts[1]))
            .ToList();
        if (mappings.Count == 0)
            return input;
        string[] missing = mappings.Select(mapping => mapping.Source)
            .Where(source => !input.Columns.Contains(source, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (missing.Length > 0)
            throw new InvalidOperationException($"Node Chọn cột không tìm thấy: {string.Join(", ", missing)}.");
        FlowTable result = new() { Columns = mappings.Select(mapping => mapping.Target).ToList() };
        result.Rows = input.Rows.Select(row => mappings.ToDictionary(
            mapping => mapping.Target,
            mapping => row.GetValueOrDefault(mapping.Source, ""),
            StringComparer.OrdinalIgnoreCase)).ToList();
        return result;
    }

    private static FlowTable ValidateRows(FlowTable input, DataFlowNode node)
    {
        string[] required = node.Settings.GetValueOrDefault("Required", "")
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        FlowTable result = new() { Columns = [.. input.Columns, "__Lỗi"] };
        foreach (Dictionary<string, string> row in input.Rows)
        {
            Dictionary<string, string> copy = new(row, StringComparer.OrdinalIgnoreCase);
            copy["__Lỗi"] = string.Join(", ", required
                .Where(column => string.IsNullOrWhiteSpace(row.GetValueOrDefault(column, "")))
                .Select(column => $"Thiếu {column}"));
            result.Rows.Add(copy);
        }
        return result;
    }

    private static FlowTable RequireInput(List<FlowTable> inputs, DataFlowNode node) =>
        inputs.FirstOrDefault() ?? throw new InvalidOperationException($"Node “{node.Title}” chưa có dây đầu vào.");

    private static Dictionary<string, string> MergeRows(
        Dictionary<string, string> left,
        Dictionary<string, string>? right,
        List<string> rightColumns,
        List<string> outputColumns)
    {
        Dictionary<string, string> result = new(left, StringComparer.OrdinalIgnoreCase);
        int start = left.Count;
        for (int index = 0; index < rightColumns.Count; index++)
            result[outputColumns[start + index]] = right?.GetValueOrDefault(rightColumns[index], "") ?? "";
        return result;
    }

    private static string UniqueColumn(List<string> existing, string requested)
    {
        if (!existing.Contains(requested, StringComparer.OrdinalIgnoreCase)) return requested;
        string candidate = $"{requested} (File 2)";
        int suffix = 2;
        while (existing.Contains(candidate, StringComparer.OrdinalIgnoreCase))
            candidate = $"{requested} (File 2-{suffix++})";
        return candidate;
    }

    private static IEnumerable<DataFlowNode> TopologicalOrder(DataFlowDefinition flow)
    {
        Dictionary<string, int> indegree = flow.Nodes.ToDictionary(node => node.Id, _ => 0);
        foreach (DataFlowConnection edge in flow.Connections)
            indegree[edge.ToNodeId]++;
        Queue<DataFlowNode> queue = new(flow.Nodes.Where(node => indegree[node.Id] == 0));
        List<DataFlowNode> ordered = new();
        while (queue.Count > 0)
        {
            DataFlowNode node = queue.Dequeue();
            ordered.Add(node);
            foreach (DataFlowConnection edge in flow.Connections.Where(edge => edge.FromNodeId == node.Id))
                if (--indegree[edge.ToNodeId] == 0)
                    queue.Enqueue(flow.Nodes.Single(item => item.Id == edge.ToNodeId));
        }
        if (ordered.Count != flow.Nodes.Count)
            throw new InvalidOperationException("Flow có dây nối vòng. Hãy nối dữ liệu từ trái sang phải.");
        return ordered;
    }

    private static void ValidateGraph(DataFlowDefinition flow)
    {
        if (flow.Nodes.Count == 0)
            throw new InvalidOperationException("Flow chưa có node.");
        HashSet<string> ids = flow.Nodes.Select(node => node.Id).ToHashSet();
        if (flow.Connections.Any(edge => !ids.Contains(edge.FromNodeId) || !ids.Contains(edge.ToNodeId)))
            throw new InvalidOperationException("Flow có dây nối không hợp lệ.");
        _ = TopologicalOrder(flow).ToList();
    }

    private static string Normalize(string value)
    {
        string decomposed = value.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
        StringBuilder result = new();
        foreach (char character in decomposed)
            if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
                result.Append(character == 'đ' ? 'd' : character);
        return result.ToString().Normalize(NormalizationForm.FormC);
    }

    private static string SafeName(string value) =>
        string.Concat(value.Select(character => Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
}
