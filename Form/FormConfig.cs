using System.ComponentModel;
using System.Diagnostics;
using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using KiotVietLabelPrinter.Ui;

namespace KiotVietLabelPrinter.Forms;

public class FormConfig : Form
{
    private readonly TextBox txtBarTender = new();
    private readonly TextBox txtDefaultEmployee = new();
    private readonly CheckBox chkRememberEmployee = new();
    private readonly ListBox lstLabels = new();
    private readonly TextBox txtCode = new();
    private readonly TextBox txtName = new();
    private readonly TextBox txtDescription = new();
    private readonly TextBox txtTemplate = new();
    private readonly TextBox txtData = new();
    private readonly ComboBox cboHandler = new();
    private readonly CheckBox chkEnabled = new();
    private readonly CheckBox chkEmployee = new();
    private readonly CheckBox chkParser = new();
    private readonly CheckBox chkAppendEmployee = new();
    private readonly Label lblStatus = new();
    private readonly Button btnSave = new();
    private readonly BindingList<LabelDefinition> labels = new();
    private bool loading;

    public FormConfig()
    {
        Text = "Quản lý mẫu tem";
        MinimumSize = new Size(1080, 720);
        Size = new Size(1180, 780);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = AppTheme.Canvas;
        Font = AppTheme.Body();

        BuildUi();
        LoadConfig();
    }

    private void BuildUi()
    {
        var header = new Panel { Dock = DockStyle.Top, Height = 96, BackColor = AppTheme.Ink };
        header.Controls.Add(new Label
        {
            Text = "Mẫu tem & dữ liệu",
            Left = 32, Top = 18, Width = 540, Height = 38,
            Font = AppTheme.Display(22), ForeColor = Color.White
        });
        header.Controls.Add(new Label
        {
            Text = "Chọn một mẫu bên trái, sau đó thay file BarTender hoặc file dữ liệu bên phải.",
            Left = 34, Top = 57, Width = 720, Height = 24,
            Font = AppTheme.Body(10), ForeColor = Color.FromArgb(196, 211, 206)
        });
        Controls.Add(header);

        var footer = new Panel { Dock = DockStyle.Bottom, Height = 76, BackColor = AppTheme.Surface };
        btnSave.Text = "Lưu thay đổi";
        btnSave.SetBounds(Width - 210, 17, 160, 42);
        btnSave.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StylePrimary(btnSave);
        btnSave.Click += Save_Click;
        footer.Controls.Add(btnSave);

        var btnCancel = new Button { Text = "Đóng", Width = 100, Height = 42, Top = 17 };
        btnCancel.Left = btnSave.Left - 116;
        btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(btnCancel);
        btnCancel.Click += (_, _) => Close();
        footer.Controls.Add(btnCancel);
        Controls.Add(footer);

        var body = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24),
            ColumnCount = 2,
            RowCount = 1,
            BackColor = AppTheme.Canvas
        };
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 310));
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        Controls.Add(body);
        body.BringToFront();

        body.Controls.Add(BuildLabelList(), 0, 0);
        body.Controls.Add(BuildEditor(), 1, 0);
    }

    private Control BuildLabelList()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Padding = new Padding(18) };
        panel.Controls.Add(new Label
        {
            Text = "Các loại tem",
            Dock = DockStyle.Top,
            Height = 32,
            Font = AppTheme.Display(15),
            ForeColor = AppTheme.Ink
        });

        var hint = new Label
        {
            Text = "Mẫu bị tắt vẫn được giữ lại và có thể bật lại bất cứ lúc nào.",
            Dock = DockStyle.Top,
            Height = 45,
            Font = AppTheme.Body(9),
            ForeColor = AppTheme.Muted
        };
        panel.Controls.Add(hint);
        hint.BringToFront();

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom, Height = 48,
            FlowDirection = FlowDirection.LeftToRight
        };
        var add = new Button { Text = "+ Thêm mẫu", Width = 122, Height = 36 };
        var remove = new Button { Text = "Bỏ khỏi app", Width = 122, Height = 36 };
        AppTheme.StylePrimary(add);
        AppTheme.StyleSecondary(remove);
        remove.ForeColor = AppTheme.Danger;
        add.Click += AddLabel_Click;
        remove.Click += RemoveLabel_Click;
        actions.Controls.Add(add);
        actions.Controls.Add(remove);
        panel.Controls.Add(actions);

        lstLabels.Dock = DockStyle.Fill;
        lstLabels.BorderStyle = BorderStyle.None;
        lstLabels.Font = AppTheme.Body(11);
        lstLabels.ItemHeight = 34;
        lstLabels.IntegralHeight = false;
        lstLabels.SelectedIndexChanged += LabelSelectionChanged;
        panel.Controls.Add(lstLabels);
        lstLabels.BringToFront();
        return panel;
    }

    private Control BuildEditor()
    {
        var host = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Canvas, Padding = new Padding(18, 0, 0, 0) };
        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = AppTheme.Surface, Padding = new Padding(28) };
        host.Controls.Add(scroll);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 3,
            RowCount = 12
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 118));
        scroll.Controls.Add(layout);

        AddTitle(layout, "Thông tin mẫu", "Tên và mô tả sẽ xuất hiện ở màn hình chọn tem.", 0);
        AddField(layout, "Tên loại tem", txtName, 1);
        AddField(layout, "Mã nhận diện", txtCode, 2);
        AddField(layout, "Mô tả ngắn", txtDescription, 3);
        AddFileField(layout, "File BarTender", txtTemplate, "Chọn .btw", "*.btw", 4);
        AddFileField(layout, "File dữ liệu", txtData, "Chọn data", "*.xls;*.xlsx;*.csv", 5);

        cboHandler.DropDownStyle = ComboBoxStyle.DropDownList;
        cboHandler.Items.AddRange(["GENERIC", "FULL", "BARCODE", "GLASSES"]);
        AddField(layout, "Cách xử lý", cboHandler, 6);

        var options = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
        chkEnabled.Text = "Hiển thị mẫu này trong app";
        chkEmployee.Text = "Yêu cầu mã nhân viên";
        chkParser.Text = "Tách mã barcode";
        chkAppendEmployee.Text = "Nối mã nhân viên";
        foreach (var check in new[] { chkEnabled, chkEmployee, chkParser, chkAppendEmployee })
        {
            check.AutoSize = true;
            check.Margin = new Padding(0, 6, 24, 6);
            options.Controls.Add(check);
            check.CheckedChanged += (_, _) => UpdateSelectedFromEditor();
        }
        layout.Controls.Add(AppTheme.Caption("Tùy chọn"), 0, 7);
        layout.Controls.Add(options, 1, 7);
        layout.SetColumnSpan(options, 2);

        AddTitle(layout, "Cài đặt chung", "Chỉ cần thiết lập một lần trên máy dùng để in.", 8);
        AddFileField(layout, "BarTender.exe", txtBarTender, "Tìm ứng dụng", "*.exe", 9);
        AddField(layout, "Mã NV mặc định", txtDefaultEmployee, 10);
        chkRememberEmployee.Text = "Ghi nhớ mã nhân viên cho lần in sau";
        chkRememberEmployee.AutoSize = true;
        layout.Controls.Add(chkRememberEmployee, 1, 11);

        lblStatus.Dock = DockStyle.Bottom;
        lblStatus.Height = 52;
        lblStatus.Padding = new Padding(14);
        lblStatus.Font = AppTheme.Body(9.5F, FontStyle.Bold);
        scroll.Controls.Add(lblStatus);

        foreach (var text in new[] { txtName, txtCode, txtDescription, txtTemplate, txtData })
            text.TextChanged += (_, _) => UpdateSelectedFromEditor();
        cboHandler.SelectedIndexChanged += (_, _) => UpdateSelectedFromEditor();

        return host;
    }

    private static void AddTitle(TableLayoutPanel layout, string title, string subtitle, int row)
    {
        var block = new Panel { Dock = DockStyle.Fill, Height = 70, Margin = new Padding(0, 4, 0, 8) };
        block.Controls.Add(new Label
        {
            Text = title, Left = 0, Top = 4, Width = 500, Height = 30,
            Font = AppTheme.Display(15), ForeColor = AppTheme.Ink
        });
        block.Controls.Add(new Label
        {
            Text = subtitle, Left = 1, Top = 35, Width = 580, Height = 25,
            Font = AppTheme.Body(9), ForeColor = AppTheme.Muted
        });
        layout.Controls.Add(block, 0, row);
        layout.SetColumnSpan(block, 3);
    }

    private static void AddField(TableLayoutPanel layout, string label, Control input, int row)
    {
        input.Dock = DockStyle.Fill;
        input.Margin = new Padding(0, 3, 8, 10);
        input.Height = 32;
        layout.Controls.Add(AppTheme.Caption(label), 0, row);
        layout.Controls.Add(input, 1, row);
        layout.SetColumnSpan(input, 2);
    }

    private void AddFileField(TableLayoutPanel layout, string label, TextBox input, string buttonText, string pattern, int row)
    {
        input.Dock = DockStyle.Fill;
        input.Margin = new Padding(0, 3, 8, 10);
        var browse = new Button { Text = buttonText, Dock = DockStyle.Top, Height = 32, Margin = new Padding(0, 3, 0, 10) };
        AppTheme.StyleSecondary(browse);
        browse.Click += (_, _) => Browse(input, pattern);
        var menu = new ContextMenuStrip();
        menu.Items.Add("Mở file", null, (_, _) => OpenPath(input.Text));
        menu.Items.Add("Mở thư mục chứa file", null, (_, _) => OpenFolder(input.Text));
        input.ContextMenuStrip = menu;
        layout.Controls.Add(AppTheme.Caption(label), 0, row);
        layout.Controls.Add(input, 1, row);
        layout.Controls.Add(browse, 2, row);
    }

    private void LoadConfig()
    {
        loading = true;
        var config = ConfigService.Instance.Config;
        txtBarTender.Text = config.BarTenderExe;
        txtDefaultEmployee.Text = config.DefaultEmployee;
        chkRememberEmployee.Checked = config.RememberEmployee;
        labels.Clear();
        foreach (var item in config.Labels)
            labels.Add(Clone(item));
        RefreshList();
        if (labels.Count > 0) lstLabels.SelectedIndex = 0;
        loading = false;
        ShowSelected();
    }

    private void RefreshList()
    {
        var selected = lstLabels.SelectedItem as LabelDefinition;
        lstLabels.DataSource = null;
        lstLabels.DisplayMember = nameof(LabelDefinition.DisplayName);
        lstLabels.DataSource = labels;
        if (selected != null) lstLabels.SelectedItem = selected;
    }

    private void LabelSelectionChanged(object? sender, EventArgs e)
    {
        if (!loading) ShowSelected();
    }

    private void ShowSelected()
    {
        loading = true;
        var item = lstLabels.SelectedItem as LabelDefinition;
        var enabled = item != null;
        foreach (var control in new Control[] { txtCode, txtName, txtDescription, txtTemplate, txtData, cboHandler, chkEnabled, chkEmployee, chkParser, chkAppendEmployee })
            control.Enabled = enabled;
        if (item != null)
        {
            txtCode.Text = item.Code;
            txtName.Text = item.Name;
            txtDescription.Text = item.Description;
            txtTemplate.Text = item.TemplatePath;
            txtData.Text = item.DataFilePath;
            cboHandler.SelectedItem = item.HandlerType;
            if (cboHandler.SelectedIndex < 0) cboHandler.SelectedItem = "GENERIC";
            chkEnabled.Checked = item.IsEnabled;
            chkEmployee.Checked = item.RequiresEmployeeCode;
            chkParser.Checked = item.UseBarcodeParser;
            chkAppendEmployee.Checked = item.AppendEmployeeCode;
        }
        loading = false;
        UpdateStatus();
    }

    private void UpdateSelectedFromEditor()
    {
        if (loading || lstLabels.SelectedItem is not LabelDefinition item) return;
        item.Code = txtCode.Text.Trim();
        item.Name = txtName.Text.Trim();
        item.Description = txtDescription.Text.Trim();
        item.TemplatePath = txtTemplate.Text.Trim();
        item.DataFilePath = txtData.Text.Trim();
        item.HandlerType = cboHandler.SelectedItem?.ToString() ?? "GENERIC";
        item.IsEnabled = chkEnabled.Checked;
        item.RequiresEmployeeCode = chkEmployee.Checked;
        item.UseBarcodeParser = chkParser.Checked;
        item.AppendEmployeeCode = chkAppendEmployee.Checked;
        RefreshList();
        UpdateStatus();
    }

    private void UpdateStatus()
    {
        if (lstLabels.SelectedItem is not LabelDefinition item)
        {
            lblStatus.Text = "Chưa có mẫu tem. Bấm “Thêm mẫu” để bắt đầu.";
            lblStatus.BackColor = AppTheme.SurfaceMuted;
            lblStatus.ForeColor = AppTheme.Muted;
            return;
        }
        var issues = item.GetReadinessIssues();
        lblStatus.Text = issues.Count == 0
            ? "✓ Mẫu đã sẵn sàng để in."
            : "Cần bổ sung: " + string.Join(" • ", issues);
        lblStatus.BackColor = issues.Count == 0 ? Color.FromArgb(224, 243, 235) : Color.FromArgb(255, 243, 220);
        lblStatus.ForeColor = issues.Count == 0 ? AppTheme.AccentDark : Color.FromArgb(136, 91, 22);
    }

    private void AddLabel_Click(object? sender, EventArgs e)
    {
        var item = new LabelDefinition
        {
            Code = $"LABEL_{labels.Count + 1}",
            Name = "Mẫu tem mới",
            Description = "Chưa có mô tả",
            HandlerType = "GENERIC",
            IsEnabled = true
        };
        labels.Add(item);
        RefreshList();
        lstLabels.SelectedItem = item;
        txtName.Focus();
        txtName.SelectAll();
    }

    private void RemoveLabel_Click(object? sender, EventArgs e)
    {
        if (lstLabels.SelectedItem is not LabelDefinition item) return;
        var answer = MessageBox.Show(
            $"Bỏ “{item.Name}” khỏi danh sách?\n\nFile .btw và file dữ liệu trên máy sẽ không bị xóa.",
            "Bỏ mẫu khỏi app", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (answer != DialogResult.Yes) return;
        labels.Remove(item);
        RefreshList();
        if (labels.Count > 0) lstLabels.SelectedIndex = 0;
        else ShowSelected();
    }

    private void Save_Click(object? sender, EventArgs e)
    {
        UpdateSelectedFromEditor();
        if (string.IsNullOrWhiteSpace(txtBarTender.Text))
        {
            MessageBox.Show("Hãy chọn file BarTender.exe trước khi lưu.", "Thiếu BarTender");
            return;
        }
        var duplicate = labels.GroupBy(x => x.Code, StringComparer.OrdinalIgnoreCase).FirstOrDefault(x => x.Count() > 1);
        if (duplicate != null)
        {
            MessageBox.Show($"Mã mẫu bị trùng: {duplicate.Key}", "Không thể lưu");
            return;
        }
        ConfigService.Instance.Config.BarTenderExe = txtBarTender.Text.Trim();
        ConfigService.Instance.Config.DefaultEmployee = txtDefaultEmployee.Text.Trim();
        ConfigService.Instance.Config.RememberEmployee = chkRememberEmployee.Checked;
        ConfigService.Instance.Config.Labels = labels.Select(Clone).ToList();
        ConfigService.Instance.Save();
        DialogResult = DialogResult.OK;
        Close();
    }

    private static LabelDefinition Clone(LabelDefinition x) => new()
    {
        Code = x.Code, Name = x.Name, Description = x.Description, IconText = x.IconText,
        IsEnabled = x.IsEnabled, TemplatePath = x.TemplatePath, DataFilePath = x.DataFilePath,
        HandlerType = x.HandlerType, RequiresEmployeeCode = x.RequiresEmployeeCode,
        UseBarcodeParser = x.UseBarcodeParser, AppendEmployeeCode = x.AppendEmployeeCode,
        TargetNameColumnIndex = x.TargetNameColumnIndex
    };

    private static void Browse(TextBox target, string pattern)
    {
        using var dialog = new OpenFileDialog { Filter = $"File phù hợp|{pattern}|Tất cả file|*.*" };
        if (File.Exists(target.Text)) dialog.InitialDirectory = Path.GetDirectoryName(target.Text);
        if (dialog.ShowDialog() == DialogResult.OK) target.Text = dialog.FileName;
    }

    private static void OpenPath(string path)
    {
        if (!File.Exists(path)) return;
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private static void OpenFolder(string path)
    {
        var folder = File.Exists(path) ? Path.GetDirectoryName(path) : path;
        if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{folder}\"") { UseShellExecute = true });
    }
}
