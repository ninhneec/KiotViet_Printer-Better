using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Ui;
using System.ComponentModel;

namespace KiotVietLabelPrinter.Forms;

public class FlowCanvas : Panel
{
    private const int NodeWidth = 190;
    private const int NodeHeight = 86;
    private readonly Dictionary<string, FlowNodeControl> controlsById = new();
    private DataFlowDefinition? flow;
    private Point dragOrigin;
    private Point nodeOrigin;
    private FlowNodeControl? dragging;
    private readonly List<FlowNodeControl> selected = new();

    public event Action<DataFlowNode?>? NodeSelected;
    public event Action? GraphChanged;

    public FlowCanvas()
    {
        DoubleBuffered = true;
        AutoScroll = true;
        Paint += DrawConnections;
        MouseDown += (_, _) => SelectOnly(null);
    }

    public void LoadFlow(DataFlowDefinition value)
    {
        flow = value;
        Controls.Clear();
        controlsById.Clear();
        selected.Clear();
        foreach (DataFlowNode node in flow.Nodes)
            AddNodeControl(node);
        AutoScrollMinSize = new Size(
            Math.Max(1500, flow.Nodes.Select(node => node.X).DefaultIfEmpty().Max() + 300),
            Math.Max(700, flow.Nodes.Select(node => node.Y).DefaultIfEmpty().Max() + 200));
        Invalidate();
    }

    public void AddNode(DataFlowNodeType type)
    {
        if (flow == null) return;
        DataFlowNode node = new()
        {
            Type = type,
            Title = DefaultTitle(type),
            X = Math.Max(30, -AutoScrollPosition.X + 70 + (flow.Nodes.Count % 4) * 35),
            Y = Math.Max(30, -AutoScrollPosition.Y + 80 + (flow.Nodes.Count % 5) * 105)
        };
        ApplyDefaults(node);
        flow.Nodes.Add(node);
        AddNodeControl(node);
        SelectOnly(controlsById[node.Id]);
        GraphChanged?.Invoke();
    }

    public void ConnectSelected()
    {
        if (flow == null || selected.Count != 2)
        {
            MessageBox.Show("Giữ Ctrl và chọn đúng 2 khối: khối nguồn trước, khối nhận sau.",
                "Nối dây", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        DataFlowNode from = selected[0].Node;
        DataFlowNode to = selected[1].Node;
        if (from.Id == to.Id) return;
        if (flow.Connections.Any(edge => edge.FromNodeId == from.Id && edge.ToNodeId == to.Id))
            return;
        int input = flow.Connections.Count(edge => edge.ToNodeId == to.Id);
        flow.Connections.Add(new DataFlowConnection { FromNodeId = from.Id, ToNodeId = to.Id, ToInput = input });
        selected.Clear();
        RefreshSelection();
        Invalidate();
        GraphChanged?.Invoke();
    }

    public void DeleteSelection()
    {
        if (flow == null || selected.Count == 0) return;
        HashSet<string> ids = selected.Select(control => control.Node.Id).ToHashSet();
        flow.Connections.RemoveAll(edge => ids.Contains(edge.FromNodeId) || ids.Contains(edge.ToNodeId));
        flow.Nodes.RemoveAll(node => ids.Contains(node.Id));
        foreach (string id in ids)
        {
            Controls.Remove(controlsById[id]);
            controlsById.Remove(id);
        }
        selected.Clear();
        NodeSelected?.Invoke(null);
        Invalidate();
        GraphChanged?.Invoke();
    }

    public void RefreshNodes()
    {
        foreach (FlowNodeControl control in controlsById.Values)
            control.UpdateText();
        Invalidate();
    }

    private void AddNodeControl(DataFlowNode node)
    {
        FlowNodeControl control = new(node)
        {
            Size = new Size(NodeWidth, NodeHeight),
            Location = new Point(node.X, node.Y)
        };
        control.MouseDown += NodeMouseDown;
        control.MouseMove += NodeMouseMove;
        control.MouseUp += NodeMouseUp;
        control.DoubleClick += NodeDoubleClick;
        foreach (Control child in control.Controls)
        {
            child.MouseDown += NodeMouseDown;
            child.MouseMove += NodeMouseMove;
            child.MouseUp += NodeMouseUp;
            child.DoubleClick += NodeDoubleClick;
        }
        controlsById[node.Id] = control;
        Controls.Add(control);
        control.BringToFront();
    }

    private void NodeDoubleClick(object? sender, EventArgs e)
    {
        FlowNodeControl control = sender as FlowNodeControl ?? (sender as Control)?.Parent as FlowNodeControl
            ?? throw new InvalidOperationException();
        if (control.Node.Type != DataFlowNodeType.ExcelSource)
            return;
        using OpenFileDialog dialog = new()
        {
            Filter = "Excel hoặc CSV|*.xls;*.xlsx;*.csv|Tất cả file|*.*",
            Title = $"Chọn file cho {control.Node.Title}"
        };
        string currentPath = control.Node.Settings.GetValueOrDefault("FilePath", "");
        if (File.Exists(currentPath))
        {
            dialog.InitialDirectory = Path.GetDirectoryName(currentPath);
            dialog.FileName = Path.GetFileName(currentPath);
        }
        if (dialog.ShowDialog(FindForm()) != DialogResult.OK)
            return;
        control.Node.Settings["FilePath"] = dialog.FileName;
        control.UpdateText();
        SelectOnly(control);
        GraphChanged?.Invoke();
    }

    private void NodeMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left) return;
        FlowNodeControl control = sender as FlowNodeControl ?? (sender as Control)?.Parent as FlowNodeControl
            ?? throw new InvalidOperationException();
        bool additive = (ModifierKeys & Keys.Control) == Keys.Control;
        if (!additive) SelectOnly(control);
        else ToggleSelection(control);
        dragging = control;
        dragOrigin = Cursor.Position;
        nodeOrigin = control.Location;
        control.Capture = true;
    }

    private void NodeMouseMove(object? sender, MouseEventArgs e)
    {
        if (dragging == null || e.Button != MouseButtons.Left) return;
        Point delta = new(Cursor.Position.X - dragOrigin.X, Cursor.Position.Y - dragOrigin.Y);
        dragging.Location = new Point(
            Math.Max(8, nodeOrigin.X + delta.X),
            Math.Max(8, nodeOrigin.Y + delta.Y));
        dragging.Node.X = dragging.Left;
        dragging.Node.Y = dragging.Top;
        Invalidate();
    }

    private void NodeMouseUp(object? sender, MouseEventArgs e)
    {
        if (dragging == null) return;
        dragging.Capture = false;
        dragging = null;
        GraphChanged?.Invoke();
    }

    private void SelectOnly(FlowNodeControl? control)
    {
        selected.Clear();
        if (control != null) selected.Add(control);
        RefreshSelection();
        NodeSelected?.Invoke(control?.Node);
    }

    private void ToggleSelection(FlowNodeControl control)
    {
        if (!selected.Remove(control)) selected.Add(control);
        RefreshSelection();
        NodeSelected?.Invoke(control.Node);
    }

    private void RefreshSelection()
    {
        foreach (FlowNodeControl control in controlsById.Values)
            control.Selected = selected.Contains(control);
    }

    private void DrawConnections(object? sender, PaintEventArgs e)
    {
        if (flow == null) return;
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using Pen pen = new(AppTheme.AccentDark, 3);
        pen.EndCap = System.Drawing.Drawing2D.LineCap.RoundAnchor;
        foreach (DataFlowConnection edge in flow.Connections)
        {
            if (!controlsById.TryGetValue(edge.FromNodeId, out FlowNodeControl? from) ||
                !controlsById.TryGetValue(edge.ToNodeId, out FlowNodeControl? to)) continue;
            Point start = new(from.Right, from.Top + from.Height / 2);
            Point end = new(to.Left, to.Top + to.Height / 2);
            int bend = Math.Max(55, Math.Abs(end.X - start.X) / 2);
            using System.Drawing.Drawing2D.GraphicsPath path = new();
            path.AddBezier(start, new Point(start.X + bend, start.Y), new Point(end.X - bend, end.Y), end);
            e.Graphics.DrawPath(pen, path);
        }
    }

    private static string DefaultTitle(DataFlowNodeType type) => type switch
    {
        DataFlowNodeType.ExcelSource => "Nguồn Excel",
        DataFlowNodeType.Join => "Nối dữ liệu",
        DataFlowNodeType.Filter => "Lọc dữ liệu",
        DataFlowNodeType.SelectColumns => "Chọn và đổi tên cột",
        DataFlowNodeType.Validate => "Kiểm tra lỗi",
        _ => "Mẫu tem BarTender"
    };

    private static void ApplyDefaults(DataFlowNode node)
    {
        if (node.Type == DataFlowNodeType.Join)
        {
            node.Settings["LeftKey"] = "Mã hàng";
            node.Settings["RightKey"] = "Mã hàng";
            node.Settings["JoinType"] = "Left";
        }
        else if (node.Type == DataFlowNodeType.Filter)
        {
            node.Settings["Column"] = "Đơn vị tính";
            node.Settings["Operation"] = "Có dữ liệu";
        }
        else if (node.Type == DataFlowNodeType.SelectColumns)
            node.Settings["Mappings"] = "Tên hàng=Tên hàng;Giá bán=Giá bán;Đơn vị tính=Đơn vị tính";
        else if (node.Type == DataFlowNodeType.Validate)
            node.Settings["Required"] = "Tên hàng;Giá bán;Đơn vị tính";
    }
}

internal class FlowNodeControl : Panel
{
    private readonly Label title = new();
    private readonly Label subtitle = new();
    [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public DataFlowNode Node { get; }
    private bool selected;

    [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool Selected
    {
        get => selected;
        set
        {
            selected = value;
            BackColor = value ? Color.FromArgb(220, 241, 233) : Color.White;
            Invalidate();
        }
    }

    public FlowNodeControl(DataFlowNode node)
    {
        Node = node;
        BackColor = Color.White;
        Padding = new Padding(13);
        Cursor = Cursors.SizeAll;
        title.Dock = DockStyle.Top;
        title.Height = 28;
        title.Font = AppTheme.Body(10, FontStyle.Bold);
        title.ForeColor = AppTheme.Ink;
        subtitle.Dock = DockStyle.Fill;
        subtitle.ForeColor = AppTheme.Muted;
        subtitle.Font = AppTheme.Body(8.5F);
        Controls.Add(subtitle);
        Controls.Add(title);
        UpdateText();
        Paint += (_, e) =>
        {
            using Pen border = new(Selected ? AppTheme.AccentDark : AppTheme.Border, Selected ? 3 : 1);
            e.Graphics.DrawRectangle(border, 0, 0, Width - 1, Height - 1);
            using Brush input = new SolidBrush(AppTheme.Ink);
            using Brush output = new SolidBrush(AppTheme.AccentDark);
            e.Graphics.FillEllipse(input, -5, Height / 2 - 6, 12, 12);
            e.Graphics.FillEllipse(output, Width - 7, Height / 2 - 6, 12, 12);
        };
    }

    public void UpdateText()
    {
        title.Text = Node.Title;
        subtitle.Text = Node.Type switch
        {
            DataFlowNodeType.ExcelSource => Path.GetFileName(Node.Settings.GetValueOrDefault("FilePath", "Chưa chọn file")),
            DataFlowNodeType.Join => $"{Node.Settings.GetValueOrDefault("LeftKey", "?")} ↔ {Node.Settings.GetValueOrDefault("RightKey", "?")}",
            DataFlowNodeType.Filter => $"{Node.Settings.GetValueOrDefault("Column", "?")} · {Node.Settings.GetValueOrDefault("Operation", "")}",
            DataFlowNodeType.SelectColumns => "Ánh xạ cột đầu ra",
            DataFlowNodeType.Validate => "Báo màu dữ liệu lỗi",
            _ => "Đầu ra BarTender"
        };
    }
}
