using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using KiotVietLabelPrinter.Ui;

namespace KiotVietLabelPrinter.Forms;

public class FormFlowDesigner : Form
{
    private readonly DataFlowService service = new();
    private readonly ListBox flowList = new();
    private readonly FlowCanvas canvas = new();
    private readonly PropertyGrid properties = new();
    private readonly TreeView logicTree = new();
    private readonly TextBox txtName = new();
    private readonly Label lblState = new();
    private DataFlowDefinition current = DataFlowService.CreateStarterFlow();

    public FormFlowDesigner()
    {
        Text = "Thiết kế Flow dữ liệu";
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(1180, 720);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = AppTheme.Canvas;
        Font = AppTheme.Body();
        BuildUi();
        ReloadFlows();
        LoadFlow(current);
    }

    private void BuildUi()
    {
        Controls.Add(BuildBody());
        Controls.Add(BuildFooter());
        Controls.Add(BuildHeader());
    }

    private Control BuildHeader()
    {
        Panel header = new() { Dock = DockStyle.Top, Height = 88, BackColor = AppTheme.Ink };
        header.Controls.Add(new Label
        {
            Text = "Flow dữ liệu cho tem",
            Left = 28, Top = 14, Width = 440, Height = 38,
            Font = AppTheme.Display(22), ForeColor = Color.White
        });
        header.Controls.Add(new Label
        {
            Text = "Thêm khối → nối dây → chạy thử → lưu quy trình.",
            Left = 30, Top = 53, Width = 540, Height = 22,
            ForeColor = Color.FromArgb(195, 211, 206)
        });
        return header;
    }

    private Control BuildBody()
    {
        TableLayoutPanel body = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            Padding = new Padding(16),
            BackColor = AppTheme.Canvas
        };
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 245));
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300));
        body.Controls.Add(BuildLeft(), 0, 0);
        body.Controls.Add(BuildCanvasHost(), 1, 0);
        body.Controls.Add(BuildRight(), 2, 0);
        return body;
    }

    private Control BuildLeft()
    {
        Panel panel = Card();
        panel.Padding = new Padding(14);
        Label title = new()
        {
            Text = "QUY TRÌNH ĐÃ LƯU", Dock = DockStyle.Top, Height = 30,
            Font = AppTheme.Body(9, FontStyle.Bold), ForeColor = AppTheme.Muted
        };
        FlowLayoutPanel actions = new()
        {
            Dock = DockStyle.Bottom, Height = 88, FlowDirection = FlowDirection.LeftToRight
        };
        Button add = SmallButton("+ Flow mới");
        Button duplicate = SmallButton("Nhân bản");
        Button delete = SmallButton("Xóa");
        delete.ForeColor = AppTheme.Danger;
        add.Click += (_, _) => LoadFlow(DataFlowService.CreateStarterFlow());
        duplicate.Click += (_, _) =>
        {
            DataFlowDefinition copy = Newtonsoft.Json.JsonConvert.DeserializeObject<DataFlowDefinition>(
                Newtonsoft.Json.JsonConvert.SerializeObject(current))!;
            copy.Id = Guid.NewGuid().ToString("N");
            copy.Name += " - bản sao";
            LoadFlow(copy);
        };
        delete.Click += (_, _) => DeleteCurrent();
        actions.Controls.AddRange([add, duplicate, delete]);
        flowList.Dock = DockStyle.Fill;
        flowList.BorderStyle = BorderStyle.None;
        flowList.DisplayMember = nameof(DataFlowDefinition.Name);
        flowList.SelectedIndexChanged += (_, _) =>
        {
            if (flowList.SelectedItem is DataFlowDefinition flow)
                LoadFlow(flow);
        };
        panel.Controls.Add(flowList);
        panel.Controls.Add(actions);
        panel.Controls.Add(title);
        return panel;
    }

    private Control BuildCanvasHost()
    {
        Panel host = Card();
        host.Margin = new Padding(12, 0, 12, 0);
        ToolStrip strip = new() { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, Height = 42 };
        strip.Items.Add(new ToolStripLabel("Thêm khối:"));
        AddNodeButton(strip, "Excel", DataFlowNodeType.ExcelSource);
        AddNodeButton(strip, "Nối", DataFlowNodeType.Join);
        AddNodeButton(strip, "Lọc", DataFlowNodeType.Filter);
        AddNodeButton(strip, "Chọn cột", DataFlowNodeType.SelectColumns);
        AddNodeButton(strip, "Kiểm tra", DataFlowNodeType.Validate);
        AddNodeButton(strip, "Mẫu tem", DataFlowNodeType.LabelOutput);
        strip.Items.Add(new ToolStripSeparator());
        ToolStripButton connect = new("Nối 2 khối");
        connect.Click += (_, _) => canvas.ConnectSelected();
        strip.Items.Add(connect);
        ToolStripButton remove = new("Xóa khối/dây");
        remove.Click += (_, _) => canvas.DeleteSelection();
        strip.Items.Add(remove);
        canvas.Dock = DockStyle.Fill;
        canvas.BackColor = Color.FromArgb(242, 245, 244);
        canvas.NodeSelected += node =>
        {
            properties.SelectedObject = node == null ? null : new FlowNodePropertyAdapter(node);
            properties.Refresh();
        };
        canvas.GraphChanged += () => RefreshTree();
        host.Controls.Add(canvas);
        host.Controls.Add(strip);
        return host;
    }

    private Control BuildRight()
    {
        TabControl tabs = new() { Dock = DockStyle.Fill, Margin = Padding.Empty };
        TabPage settings = new("Cài đặt khối") { BackColor = AppTheme.Surface };
        properties.Dock = DockStyle.Fill;
        properties.ToolbarVisible = false;
        properties.PropertySort = PropertySort.Categorized;
        properties.PropertyValueChanged += (_, _) =>
        {
            canvas.RefreshNodes();
            RefreshTree();
        };
        settings.Controls.Add(properties);
        TabPage tree = new("Cây logic") { BackColor = AppTheme.Surface };
        logicTree.Dock = DockStyle.Fill;
        logicTree.BorderStyle = BorderStyle.None;
        logicTree.Font = AppTheme.Body(9.5F);
        tree.Controls.Add(logicTree);
        tabs.TabPages.Add(settings);
        tabs.TabPages.Add(tree);
        return tabs;
    }

    private Control BuildFooter()
    {
        Panel footer = new() { Dock = DockStyle.Bottom, Height = 76, BackColor = AppTheme.Surface };
        Label nameLabel = new()
        {
            Text = "Tên Flow", Left = 22, Top = 14, Width = 80, Height = 20,
            ForeColor = AppTheme.Muted
        };
        txtName.SetBounds(22, 34, 330, 30);
        txtName.TextChanged += (_, _) => current.Name = txtName.Text.Trim();
        lblState.SetBounds(380, 28, 430, 32);
        lblState.ForeColor = AppTheme.Muted;
        Button test = new() { Text = "▶ Chạy thử", Width = 130, Height = 40, Top = 18 };
        Button save = new() { Text = "Lưu Flow", Width = 130, Height = 40, Top = 18 };
        Button close = new() { Text = "Đóng", Width = 90, Height = 40, Top = 18 };
        close.Left = footer.Width - 112;
        save.Left = close.Left - 144;
        test.Left = save.Left - 144;
        close.Anchor = save.Anchor = test.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(close);
        AppTheme.StylePrimary(save);
        AppTheme.StyleSecondary(test);
        close.Click += (_, _) => Close();
        save.Click += (_, _) => SaveCurrent();
        test.Click += (_, _) => RunTest();
        footer.Controls.AddRange([nameLabel, txtName, lblState, test, save, close]);
        return footer;
    }

    private static Panel Card() => new()
    {
        Dock = DockStyle.Fill,
        BackColor = AppTheme.Surface,
        Margin = Padding.Empty
    };

    private static Button SmallButton(string text)
    {
        Button button = new() { Text = text, Width = 96, Height = 32 };
        AppTheme.StyleSecondary(button);
        return button;
    }

    private void AddNodeButton(ToolStrip strip, string text, DataFlowNodeType type)
    {
        ToolStripButton button = new(text);
        button.Click += (_, _) => canvas.AddNode(type);
        strip.Items.Add(button);
    }

    private void ReloadFlows()
    {
        flowList.DataSource = service.LoadAll().ToList();
    }

    private void LoadFlow(DataFlowDefinition flow)
    {
        current = flow;
        txtName.Text = flow.Name;
        canvas.LoadFlow(flow);
        RefreshTree();
        lblState.Text = "Chọn một khối để sửa cài đặt.";
    }

    private void SaveCurrent()
    {
        if (string.IsNullOrWhiteSpace(current.Name))
        {
            MessageBox.Show("Hãy đặt tên cho Flow.", "Chưa thể lưu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        try
        {
            service.Save(current);
            ReloadFlows();
            lblState.Text = $"✓ Đã lưu lúc {DateTime.Now:HH:mm:ss}";
            lblState.ForeColor = AppTheme.AccentDark;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Không lưu được Flow", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void DeleteCurrent()
    {
        if (MessageBox.Show($"Xóa Flow “{current.Name}”?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        service.Delete(current);
        ReloadFlows();
        LoadFlow(DataFlowService.CreateStarterFlow());
    }

    private void RunTest()
    {
        try
        {
            FlowTable result = service.Execute(current);
            using FormFlowPreview preview = new(current.Name, result);
            preview.ShowDialog(this);
            lblState.Text = $"✓ Flow chạy được: {result.Rows.Count:N0} dòng";
            lblState.ForeColor = AppTheme.AccentDark;
        }
        catch (Exception ex)
        {
            lblState.Text = "Flow cần sửa trước khi chạy.";
            lblState.ForeColor = AppTheme.Danger;
            MessageBox.Show(ex.Message, "Flow chưa chạy được", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void RefreshTree()
    {
        logicTree.Nodes.Clear();
        TreeNode root = logicTree.Nodes.Add(string.IsNullOrWhiteSpace(current.Name) ? "Flow chưa đặt tên" : current.Name);
        foreach (DataFlowNode node in OrderedForDisplay(current))
        {
            TreeNode item = root.Nodes.Add($"{NodeIcon(node.Type)} {node.Title}");
            foreach (KeyValuePair<string, string> setting in node.Settings.Where(pair => !string.IsNullOrWhiteSpace(pair.Value)))
                item.Nodes.Add($"{SettingName(setting.Key)}: {Path.GetFileName(setting.Value)}");
        }
        root.Expand();
    }

    private static IEnumerable<DataFlowNode> OrderedForDisplay(DataFlowDefinition flow)
    {
        HashSet<string> yielded = new();
        Queue<DataFlowNode> queue = new(flow.Nodes.Where(node =>
            flow.Connections.All(edge => edge.ToNodeId != node.Id)).OrderBy(node => node.Y));
        while (queue.Count > 0)
        {
            DataFlowNode node = queue.Dequeue();
            if (!yielded.Add(node.Id)) continue;
            yield return node;
            foreach (DataFlowConnection edge in flow.Connections.Where(edge => edge.FromNodeId == node.Id))
            {
                DataFlowNode? next = flow.Nodes.FirstOrDefault(item => item.Id == edge.ToNodeId);
                if (next != null) queue.Enqueue(next);
            }
        }
        foreach (DataFlowNode node in flow.Nodes.Where(node => !yielded.Contains(node.Id)))
            yield return node;
    }

    private static string NodeIcon(DataFlowNodeType type) => type switch
    {
        DataFlowNodeType.ExcelSource => "▦",
        DataFlowNodeType.Join => "⌘",
        DataFlowNodeType.Filter => "▽",
        DataFlowNodeType.SelectColumns => "≡",
        DataFlowNodeType.Validate => "✓",
        _ => "▣"
    };

    private static string SettingName(string key) => key switch
    {
        "FilePath" => "File",
        "LeftKey" => "Cột file 1",
        "RightKey" => "Cột file 2",
        "Mappings" => "Ánh xạ",
        "Required" => "Bắt buộc",
        _ => key
    };
}

