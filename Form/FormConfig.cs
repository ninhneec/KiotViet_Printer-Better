using System.ComponentModel;
using System.Diagnostics;
using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using KiotVietLabelPrinter.Ui;

namespace KiotVietLabelPrinter.Forms;

public class FormConfig : Form
{
    private readonly TextBox txtBarTender = new();
    private readonly ListBox lstLabels = new();
    private readonly TextBox txtCode = new();
    private readonly TextBox txtName = new();
    private readonly TextBox txtDescription = new();
    private readonly TextBox txtTemplate = new();
    private readonly TextBox txtDataFile = new();
    private readonly CheckBox chkEnabled = new();
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
            Text = "Quản lý mẫu tem",
            Left = 32, Top = 18, Width = 540, Height = 38,
            Font = AppTheme.Display(22), ForeColor = Color.White
        });
        header.Controls.Add(new Label
        {
            Text = "Mỗi mẫu dùng một file data cố định gồm: Tên hàng, Giá bán, Đơn vị tính.",
            Left = 34, Top = 57, Width = 720, Height = 24,
            Font = AppTheme.Body(10), ForeColor = Color.FromArgb(196, 211, 206)
        });
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
        body.Controls.Add(BuildLabelList(), 0, 0);
        body.Controls.Add(BuildEditor(), 1, 0);

        // Thứ tự thêm control rất quan trọng với WinForms Dock.
        // Fill phải được thêm trước, sau đó footer/header để chúng không bị che.
        Controls.Add(body);
        Controls.Add(footer);
        Controls.Add(header);
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
        var add = new Button { Text = "+ Nhập mẫu", Width = 122, Height = 36 };
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
            RowCount = 10
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 118));
        scroll.Controls.Add(layout);

        AddTitle(layout, "Thông tin mẫu", "Tên và mô tả sẽ xuất hiện ở màn hình chọn tem.", 0);
        AddField(layout, "Tên loại tem", txtName, 1);
        txtCode.ReadOnly = true;
        txtCode.BackColor = AppTheme.SurfaceMuted;
        AddField(layout, "Mã mẫu (tự động)", txtCode, 2);
        AddField(layout, "Mô tả ngắn", txtDescription, 3);
        AddFileField(layout, "File BarTender", txtTemplate, "Chọn .btw", "*.btw", 4);
        AddManagedDataField(layout, 5);
        chkEnabled.Text = "Hiển thị mẫu này trong app";
        chkEnabled.AutoSize = true;
        chkEnabled.CheckedChanged += (_, _) => UpdateSelectedFromEditor();
        layout.Controls.Add(AppTheme.Caption("Hiển thị"), 0, 6);
        layout.Controls.Add(chkEnabled, 1, 6);
        var saveCurrent = new Button
        {
            Text = "Lưu mẫu này",
            Dock = DockStyle.Top,
            Height = 34,
            Margin = new Padding(0, 0, 0, 8)
        };
        AppTheme.StylePrimary(saveCurrent);
        saveCurrent.Click += (_, _) => PersistConfig(closeAfterSave: false);
        layout.Controls.Add(saveCurrent, 2, 6);

        AddTitle(layout, "BarTender", "Chọn chương trình BarTender đã cài trên máy.", 7);
        AddFileField(layout, "BarTender.exe", txtBarTender, "Tìm ứng dụng", "*.exe", 8);
        AddFlowSettings(layout, 9);

        lblStatus.Dock = DockStyle.Bottom;
        lblStatus.Height = 68;
        lblStatus.Padding = new Padding(14);
        lblStatus.Font = AppTheme.Body(9.5F, FontStyle.Bold);
        scroll.Controls.Add(lblStatus);

        foreach (var text in new[] { txtName, txtCode, txtDescription, txtTemplate })
            text.TextChanged += (_, _) => UpdateSelectedFromEditor();
        txtBarTender.TextChanged += (_, _) =>
        {
            if (!loading)
                UpdateStatus();
        };

        return host;
    }

    private void AddFlowSettings(TableLayoutPanel layout, int row)
    {
        var block = new Panel { Dock = DockStyle.Fill, Height = 82, Margin = new Padding(0, 12, 0, 8) };
        block.Controls.Add(new Label
        {
            Text = "Flow dữ liệu", Left = 0, Top = 2, Width = 260, Height = 28,
            Font = AppTheme.Display(15), ForeColor = AppTheme.Ink
        });
        block.Controls.Add(new Label
        {
            Text = "Ghép nhiều file, lọc và đổi cột trước khi đưa sang mẫu tem.",
            Left = 1, Top = 34, Width = 500, Height = 24,
            Font = AppTheme.Body(9), ForeColor = AppTheme.Muted
        });
        var openFlow = new Button
        {
            Text = "Mở ghép dữ liệu",
            Width = 150, Height = 36, Top = 18, Left = 520,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        AppTheme.StylePrimary(openFlow);
        openFlow.Click += (_, _) =>
        {
            using var form = new FormFlowWizard();
            form.ShowDialog(this);
        };
        block.Controls.Add(openFlow);
        layout.Controls.Add(block, 0, row);
        layout.SetColumnSpan(block, 3);
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
        browse.Click += (_, _) =>
        {
            if (label.Equals("BarTender.exe", StringComparison.OrdinalIgnoreCase))
                BrowseBarTender();
            else
                Browse(input, pattern);
        };
        var menu = new ContextMenuStrip();
        menu.Items.Add("Mở file", null, (_, _) => OpenPath(input.Text));
        menu.Items.Add("Mở thư mục chứa file", null, (_, _) => OpenFolder(input.Text));
        input.ContextMenuStrip = menu;
        layout.Controls.Add(AppTheme.Caption(label), 0, row);
        layout.Controls.Add(input, 1, row);
        layout.Controls.Add(browse, 2, row);
    }

    private void AddManagedDataField(TableLayoutPanel layout, int row)
    {
        txtDataFile.ReadOnly = true;
        txtDataFile.BackColor = AppTheme.SurfaceMuted;
        txtDataFile.Dock = DockStyle.Fill;
        txtDataFile.Margin = new Padding(0, 3, 8, 10);
        var choose = new Button
        {
            Text = "Chọn data",
            Dock = DockStyle.Top,
            Height = 32,
            Margin = new Padding(0, 3, 0, 10)
        };
        AppTheme.StyleSecondary(choose);
        choose.Click += (_, _) => BrowseDataFile();
        var menu = new ContextMenuStrip();
        menu.Items.Add("Mở file", null, (_, _) => OpenPath(txtDataFile.Text));
        menu.Items.Add("Mở thư mục chứa file", null, (_, _) => OpenFolder(txtDataFile.Text));
        menu.Items.Add("Dùng file app tự tạo", null, (_, _) =>
        {
            txtDataFile.Text = ConfigService.Instance.GetManagedDataFilePath(txtCode.Text.Trim());
            UpdateSelectedFromEditor();
        });
        txtDataFile.ContextMenuStrip = menu;
        layout.Controls.Add(AppTheme.Caption("Data cố định"), 0, row);
        layout.Controls.Add(txtDataFile, 1, row);
        layout.Controls.Add(choose, 2, row);
    }

    private void LoadConfig()
    {
        loading = true;
        var config = ConfigService.Instance.Config;
        txtBarTender.Text = config.BarTenderExe;
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
        foreach (var control in new Control[] { txtCode, txtName, txtDescription, txtTemplate, txtDataFile, chkEnabled })
            control.Enabled = enabled;
        if (item != null)
        {
            txtCode.Text = item.Code;
            txtName.Text = item.Name;
            txtDescription.Text = item.Description;
            txtTemplate.Text = item.TemplatePath;
            txtDataFile.Text = item.DataFilePath;
            chkEnabled.Checked = item.IsEnabled;
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
        item.DataFilePath = string.IsNullOrWhiteSpace(txtDataFile.Text)
            ? ConfigService.Instance.GetManagedDataFilePath(item.Code)
            : txtDataFile.Text.Trim();
        txtDataFile.Text = item.DataFilePath;
        item.HandlerType = "DIRECT_PRICE";
        item.IsEnabled = chkEnabled.Checked;
        item.RequiresEmployeeCode = false;
        item.UseBarcodeParser = false;
        item.AppendEmployeeCode = false;
        // Không bind lại DataSource trong lúc đang gõ. Bind lại sẽ làm
        // ListBox đổi selection và đổ dữ liệu của mẫu khác vào form.
        lstLabels.Refresh();
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
        if (!File.Exists(txtBarTender.Text) ||
            !Path.GetFileName(txtBarTender.Text).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase))
        {
            issues.Add("chưa chọn đúng bartend.exe");
        }
        lblStatus.Text = issues.Count == 0
            ? $"✓ Mẫu đã sẵn sàng để in.\nCài đặt: {ConfigService.Instance.ConfigFilePath}"
            : "Cần bổ sung: " + string.Join(" • ", issues) +
              $"\nCài đặt: {ConfigService.Instance.ConfigFilePath}";
        lblStatus.BackColor = issues.Count == 0 ? Color.FromArgb(224, 243, 235) : Color.FromArgb(255, 243, 220);
        lblStatus.ForeColor = issues.Count == 0 ? AppTheme.AccentDark : Color.FromArgb(136, 91, 22);
    }

    private void AddLabel_Click(object? sender, EventArgs e)
    {
        using OpenFileDialog dialog = new()
        {
            Filter = "Mẫu BarTender|*.btw",
            Title = "Chọn file mẫu BarTender"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        string baseCode = Path.GetFileNameWithoutExtension(dialog.FileName)
            .ToUpperInvariant()
            .Replace(" ", "_");
        string code = baseCode;
        int suffix = 2;
        while (labels.Any(item => item.Code.Equals(code, StringComparison.OrdinalIgnoreCase)))
            code = $"{baseCode}_{suffix++}";

        var item = new LabelDefinition
        {
            Code = code,
            Name = Path.GetFileNameWithoutExtension(dialog.FileName),
            Description = "Tên hàng, giá bán và đơn vị tính",
            TemplatePath = dialog.FileName,
            DataFilePath = ConfigService.Instance.GetManagedDataFilePath(code),
            HandlerType = "DIRECT_PRICE",
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
        PersistConfig(closeAfterSave: true);
    }

    private void PersistConfig(bool closeAfterSave)
    {
        UpdateSelectedFromEditor();
        string? selectedCode = (lstLabels.SelectedItem as LabelDefinition)?.Code;
        var duplicate = labels.GroupBy(x => x.Code, StringComparer.OrdinalIgnoreCase).FirstOrDefault(x => x.Count() > 1);
        if (duplicate != null)
        {
            MessageBox.Show($"Mã mẫu bị trùng: {duplicate.Key}", "Không thể lưu");
            return;
        }
        ConfigService.Instance.Config.BarTenderExe = txtBarTender.Text.Trim();
        ConfigService.Instance.Config.Labels = labels
            .Select(Clone)
            .Select(item =>
            {
                // Lưu một bản dự phòng trong LocalAppData nhưng vẫn giữ nguyên
                // đường dẫn mà người dùng đã chọn trong màn hình setting.
                ConfigService.Instance.StoreTemplate(item.TemplatePath, item.Code);
                if (string.IsNullOrWhiteSpace(item.DataFilePath))
                    item.DataFilePath = ConfigService.Instance.GetManagedDataFilePath(item.Code);
                new ExcelService().EnsureDirectPriceDataFile(item.DataFilePath);
                return item;
            })
            .ToList();
        ConfigService.Instance.Save();

        if (closeAfterSave)
        {
            DialogResult = DialogResult.OK;
            Close();
            return;
        }

        loading = true;
        labels.Clear();
        foreach (LabelDefinition item in ConfigService.Instance.Config.Labels)
            labels.Add(Clone(item));
        RefreshList();
        LabelDefinition? savedSelection = labels.FirstOrDefault(item =>
            item.Code.Equals(selectedCode, StringComparison.OrdinalIgnoreCase));
        if (savedSelection != null)
            lstLabels.SelectedItem = savedSelection;
        else if (labels.Count > 0)
            lstLabels.SelectedIndex = 0;
        loading = false;
        ShowSelected();
        lblStatus.Text = "✓ Đã lưu mẫu vào máy. Có thể đóng cửa sổ và sử dụng ngay.";
        lblStatus.BackColor = Color.FromArgb(224, 243, 235);
        lblStatus.ForeColor = AppTheme.AccentDark;
    }

    private static LabelDefinition Clone(LabelDefinition x) => new()
    {
        Code = x.Code, Name = x.Name, Description = x.Description, IconText = x.IconText,
        IsEnabled = x.IsEnabled, TemplatePath = x.TemplatePath, DataFilePath = x.DataFilePath,
        HandlerType = "DIRECT_PRICE", RequiresEmployeeCode = false,
        UseBarcodeParser = false, AppendEmployeeCode = false,
        TargetNameColumnIndex = x.TargetNameColumnIndex
    };

    private static void Browse(TextBox target, string pattern)
    {
        using var dialog = new OpenFileDialog { Filter = $"File phù hợp|{pattern}|Tất cả file|*.*" };
        if (File.Exists(target.Text)) dialog.InitialDirectory = Path.GetDirectoryName(target.Text);
        if (dialog.ShowDialog() == DialogResult.OK) target.Text = dialog.FileName;
    }

    private void BrowseDataFile()
    {
        using OpenFileDialog dialog = new()
        {
            Filter = "File data BarTender|*.xls;*.xlsx|Tất cả file|*.*",
            Title = "Chọn đúng file data mà mẫu BarTender đang kết nối"
        };
        if (File.Exists(txtDataFile.Text))
            dialog.InitialDirectory = Path.GetDirectoryName(txtDataFile.Text);
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;
        txtDataFile.Text = dialog.FileName;
        UpdateSelectedFromEditor();
    }

    private void BrowseBarTender()
    {
        IReadOnlyList<string> detected = ConfigService.Instance.FindBarTenderExecutables();
        using OpenFileDialog dialog = new()
        {
            Filter = "BarTender (bartend.exe)|bartend.exe",
            Title = "Chọn đúng file bartend.exe của BarTender",
            CheckFileExists = true,
            FileName = "bartend.exe"
        };
        if (File.Exists(txtBarTender.Text))
            dialog.InitialDirectory = Path.GetDirectoryName(txtBarTender.Text);
        else if (detected.Count > 0)
            dialog.InitialDirectory = Path.GetDirectoryName(detected[0]);

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;
        if (!Path.GetFileName(dialog.FileName).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(
                "File này không phải bartend.exe.\n\nHãy mở thư mục cài BarTender và chọn đúng bartend.exe.",
                "Chọn sai ứng dụng",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }
        txtBarTender.Text = dialog.FileName;
        UpdateStatus();
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
