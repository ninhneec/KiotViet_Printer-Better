using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using KiotVietLabelPrinter.Ui;

namespace KiotVietLabelPrinter.Forms;

public class FormMain : Form
{
    private readonly LabelService labelService = new();
    private readonly LabelCatalogService catalogService = new();
    private readonly TextBox txtExcel = new();
    private readonly TextBox txtEmployee = new();
    private readonly DataGridView grid = new();
    private readonly FlowLayoutPanel templateList = new();
    private readonly Label lblFileState = new();
    private readonly Label lblTemplateState = new();
    private readonly Label lblSummary = new();
    private readonly Label lblEmployee = new();
    private readonly Button btnPreview = new();
    private readonly Button btnPrint = new();
    private LabelDefinition? selectedLabel;
    private List<ProductRow> products = new();

    public FormMain()
    {
        Text = "In tem KiotViet";
        MinimumSize = new Size(1120, 720);
        Size = new Size(1320, 820);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = AppTheme.Canvas;
        Font = AppTheme.Body();
        AllowDrop = true;
        DragEnter += OnDragEnter;
        DragDrop += OnDragDrop;

        BuildUi();
        RestoreLastFile();
        ReloadTemplates();
    }

    private void BuildUi()
    {
        Controls.Add(BuildFooter());
        Controls.Add(BuildBody());
        Controls.Add(BuildHeader());
    }

    private Control BuildHeader()
    {
        var header = new Panel { Dock = DockStyle.Top, Height = 92, BackColor = AppTheme.Ink };
        header.Controls.Add(new Label
        {
            Text = "IN TEM",
            Left = 28, Top = 14, Width = 240, Height = 40,
            Font = AppTheme.Display(24), ForeColor = Color.White
        });
        header.Controls.Add(new Label
        {
            Text = "Chọn file → kiểm tra dữ liệu → chọn mẫu → in",
            Left = 30, Top = 55, Width = 520, Height = 24,
            Font = AppTheme.Body(10), ForeColor = Color.FromArgb(195, 211, 206)
        });

        var config = new Button { Text = "Mẫu tem & cài đặt", Width = 170, Height = 40, Top = 25 };
        config.Left = header.Width - 322;
        config.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(config);
        config.Click += (_, _) =>
        {
            using var form = new FormConfig();
            if (form.ShowDialog(this) == DialogResult.OK) ReloadTemplates();
        };
        header.Controls.Add(config);

        var history = new Button { Text = "Lịch sử", Width = 110, Height = 40, Top = 25 };
        history.Left = header.Width - 136;
        history.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(history);
        history.Click += (_, _) => { using var form = new FormHistory(); form.ShowDialog(this); };
        header.Controls.Add(history);
        return header;
    }

    private Control BuildBody()
    {
        var body = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(22, 18, 22, 12),
            RowCount = 2,
            ColumnCount = 2,
            BackColor = AppTheme.Canvas
        };
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
        body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));

        body.Controls.Add(BuildFileStep(), 0, 0);
        body.SetColumnSpan(body.GetControlFromPosition(0, 0)!, 2);
        body.Controls.Add(BuildDataPanel(), 0, 1);
        body.Controls.Add(BuildTemplatePanel(), 1, 1);
        return body;
    }

    private Control BuildFileStep()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Padding = new Padding(20) };
        panel.Controls.Add(new Label
        {
            Text = "1  Chọn file Excel từ KiotViet",
            Left = 18, Top = 13, Width = 330, Height = 28,
            Font = AppTheme.Display(14), ForeColor = AppTheme.Ink
        });

        txtExcel.SetBounds(20, 52, 720, 34);
        txtExcel.ReadOnly = true;
        txtExcel.PlaceholderText = "Chưa chọn file .xls hoặc .xlsx";
        panel.Controls.Add(txtExcel);

        var choose = new Button { Text = "Chọn file", Width = 120, Height = 34, Left = 756, Top = 50 };
        AppTheme.StylePrimary(choose);
        choose.Click += (_, _) => ChooseExcel();
        panel.Controls.Add(choose);

        lblFileState.SetBounds(894, 53, 330, 36);
        lblFileState.ForeColor = AppTheme.Muted;
        lblFileState.Font = AppTheme.Body(9.5F, FontStyle.Bold);
        panel.Controls.Add(lblFileState);
        return panel;
    }

    private Control BuildDataPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Margin = new Padding(0, 12, 10, 0), Padding = new Padding(18) };
        panel.Controls.Add(new Label
        {
            Text = "2  Kiểm tra dữ liệu",
            Dock = DockStyle.Top, Height = 34,
            Font = AppTheme.Display(14), ForeColor = AppTheme.Ink
        });
        lblSummary.Dock = DockStyle.Top;
        lblSummary.Height = 28;
        lblSummary.ForeColor = AppTheme.Muted;
        lblSummary.Text = "Dữ liệu sản phẩm sẽ hiện ở đây.";
        panel.Controls.Add(lblSummary);
        lblSummary.BringToFront();

        grid.Dock = DockStyle.Fill;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect = false;
        grid.BackgroundColor = AppTheme.Surface;
        grid.BorderStyle = BorderStyle.None;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.SurfaceMuted;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.Ink;
        grid.ColumnHeadersDefaultCellStyle.Font = AppTheme.Body(9F, FontStyle.Bold);
        grid.EnableHeadersVisualStyles = false;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(218, 238, 231);
        grid.DefaultCellStyle.SelectionForeColor = AppTheme.Ink;
        panel.Controls.Add(grid);
        grid.BringToFront();
        return panel;
    }

    private Control BuildTemplatePanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Margin = new Padding(10, 12, 0, 0), Padding = new Padding(18) };
        panel.Controls.Add(new Label
        {
            Text = "3  Chọn loại tem",
            Dock = DockStyle.Top, Height = 34,
            Font = AppTheme.Display(14), ForeColor = AppTheme.Ink
        });
        lblTemplateState.Dock = DockStyle.Bottom;
        lblTemplateState.Height = 48;
        lblTemplateState.Padding = new Padding(10);
        lblTemplateState.Font = AppTheme.Body(9F, FontStyle.Bold);
        panel.Controls.Add(lblTemplateState);

        var employeePanel = new Panel { Dock = DockStyle.Bottom, Height = 70 };
        lblEmployee.Text = "Mã nhân viên";
        lblEmployee.SetBounds(0, 7, 130, 24);
        txtEmployee.SetBounds(0, 34, 250, 30);
        employeePanel.Controls.Add(lblEmployee);
        employeePanel.Controls.Add(txtEmployee);
        panel.Controls.Add(employeePanel);

        templateList.Dock = DockStyle.Fill;
        templateList.AutoScroll = true;
        templateList.FlowDirection = FlowDirection.TopDown;
        templateList.WrapContents = false;
        templateList.Padding = new Padding(0, 6, 0, 6);
        panel.Controls.Add(templateList);
        templateList.BringToFront();
        return panel;
    }

    private Control BuildFooter()
    {
        var footer = new Panel { Dock = DockStyle.Bottom, Height = 82, BackColor = AppTheme.Surface, Padding = new Padding(22) };
        btnPrint.Text = "In tem";
        btnPrint.Width = 170;
        btnPrint.Height = 44;
        btnPrint.Left = footer.Width - 194;
        btnPrint.Top = 18;
        btnPrint.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StylePrimary(btnPrint);
        btnPrint.Enabled = false;
        btnPrint.Click += Print_Click;
        footer.Controls.Add(btnPrint);

        btnPreview.Text = "Xem trước bản in";
        btnPreview.Width = 170;
        btnPreview.Height = 44;
        btnPreview.Left = btnPrint.Left - 184;
        btnPreview.Top = 18;
        btnPreview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(btnPreview);
        btnPreview.Enabled = false;
        btnPreview.Click += Preview_Click;
        footer.Controls.Add(btnPreview);

        footer.Controls.Add(new Label
        {
            Text = "App không thay đổi file Excel gốc.",
            Left = 24, Top = 31, Width = 400, Height = 24,
            Font = AppTheme.Body(9.5F), ForeColor = AppTheme.Muted
        });
        return footer;
    }

    private void ChooseExcel()
    {
        using var dialog = new OpenFileDialog { Filter = "Excel KiotViet|*.xls;*.xlsx|Tất cả file|*.*" };
        var lastFolder = ConfigService.Instance.Config.LastFolder;
        if (Directory.Exists(lastFolder)) dialog.InitialDirectory = lastFolder;
        if (dialog.ShowDialog(this) == DialogResult.OK) LoadExcel(dialog.FileName);
    }

    private void LoadExcel(string path)
    {
        try
        {
            products = labelService.ReadProducts(path);
            txtExcel.Text = path;
            grid.DataSource = products;
            FormatGrid();
            lblFileState.Text = $"✓ Đã đọc {products.Count:N0} sản phẩm";
            lblFileState.ForeColor = AppTheme.AccentDark;
            lblSummary.Text = $"{products.Count:N0} sản phẩm • Tổng số lượng {products.Sum(x => x.Quantity):N0}";
            ConfigService.Instance.Config.LastExcelFile = path;
            ConfigService.Instance.Config.LastFolder = Path.GetDirectoryName(path) ?? "";
            ConfigService.Instance.Save();
            UpdateActions();
        }
        catch (Exception ex)
        {
            products.Clear();
            grid.DataSource = null;
            lblFileState.Text = "Không đọc được file";
            lblFileState.ForeColor = AppTheme.Danger;
            MessageBox.Show(ex.Message, "Không đọc được dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            UpdateActions();
        }
    }

    private void FormatGrid()
    {
        RenameColumn("ProductCode", "Mã hàng", 75);
        RenameColumn("Barcode", "Mã vạch", 90);
        RenameColumn("ProductName", "Tên hàng", 145);
        RenameColumn("ProductNameWithAttr", "Tên hàng (thuộc tính)", 170);
        RenameColumn("Unit", "Đơn vị tính", 65);
        RenameColumn("Quantity", "Số lượng", 65);
        RenameColumn("Price", "Giá bán", 75);
        RenameColumn("Description", "Mô tả", 100);
        if (grid.Columns["Price"] is { } price) price.DefaultCellStyle.Format = "N0";
    }

    private void RenameColumn(string name, string text, float weight)
    {
        if (grid.Columns[name] is not { } column) return;
        column.HeaderText = text;
        column.FillWeight = weight;
    }

    private void ReloadTemplates()
    {
        selectedLabel = null;
        templateList.Controls.Clear();
        foreach (var item in catalogService.GetAllEnabled())
            templateList.Controls.Add(CreateTemplateButton(item));
        if (templateList.Controls.Count == 0)
        {
            templateList.Controls.Add(new Label
            {
                Text = "Chưa có mẫu tem.\nMở “Mẫu tem & cài đặt” để thêm mẫu đầu tiên.",
                Width = 300, Height = 70, ForeColor = AppTheme.Muted
            });
        }
        UpdateActions();
    }

    private Control CreateTemplateButton(LabelDefinition item)
    {
        var issues = item.GetReadinessIssues();
        var ready = issues.Count == 0;
        var button = new Button
        {
            Text = $"{item.Name}\r\n{(ready ? "Sẵn sàng" : string.Join(", ", issues))}",
            TextAlign = ContentAlignment.MiddleLeft,
            Width = Math.Max(280, templateList.ClientSize.Width - 28),
            Height = 68,
            Margin = new Padding(0, 0, 0, 9),
            Tag = item,
            Enabled = ready,
            AutoEllipsis = true
        };
        AppTheme.StyleSecondary(button);
        button.ForeColor = ready ? AppTheme.Ink : AppTheme.Muted;
        button.Click += (_, _) => SelectTemplate(item, button);
        return button;
    }

    private void SelectTemplate(LabelDefinition item, Button selectedButton)
    {
        selectedLabel = item;
        foreach (Control control in templateList.Controls)
        {
            if (control is not Button button) continue;
            AppTheme.StyleSecondary(button);
        }
        selectedButton.BackColor = Color.FromArgb(218, 238, 231);
        selectedButton.FlatAppearance.BorderColor = AppTheme.Accent;
        var showEmployee = item.RequiresEmployeeCode || item.AppendEmployeeCode || item.HandlerType is "BARCODE" or "GLASSES";
        lblEmployee.Visible = showEmployee;
        txtEmployee.Visible = showEmployee;
        lblEmployee.Text = item.HandlerType == "GLASSES" ? "Mã màu" : "Mã nhân viên";
        if (showEmployee && string.IsNullOrWhiteSpace(txtEmployee.Text))
            txtEmployee.Text = ConfigService.Instance.Config.DefaultEmployee;
        UpdateActions();
    }

    private void UpdateActions()
    {
        var ready = products.Count > 0 && selectedLabel != null && selectedLabel.GetReadinessIssues().Count == 0;
        btnPreview.Enabled = ready;
        btnPrint.Enabled = ready;
        lblTemplateState.Text = selectedLabel == null
            ? "Chưa chọn loại tem."
            : ready ? $"✓ Sẽ in bằng: {selectedLabel.Name}" : "Hãy chọn file dữ liệu trước.";
        lblTemplateState.BackColor = ready ? Color.FromArgb(224, 243, 235) : AppTheme.SurfaceMuted;
        lblTemplateState.ForeColor = ready ? AppTheme.AccentDark : AppTheme.Muted;
    }

    private void Preview_Click(object? sender, EventArgs e)
    {
        if (!EnsureReady()) return;
        using var form = new FormPreview(txtExcel.Text, selectedLabel!.Code, txtEmployee.Text.Trim());
        form.ShowDialog(this);
    }

    private async void Print_Click(object? sender, EventArgs e)
    {
        if (!EnsureReady()) return;
        var answer = MessageBox.Show(
            $"In {products.Sum(x => x.Quantity):N0} tem bằng mẫu “{selectedLabel!.Name}”?",
            "Xác nhận in", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (answer != DialogResult.Yes) return;

        btnPrint.Enabled = false;
        btnPreview.Enabled = false;
        btnPrint.Text = "Đang in…";
        UseWaitCursor = true;
        try
        {
            var count = await Task.Run(() => labelService.Print(txtExcel.Text, selectedLabel.Code, txtEmployee.Text.Trim()));
            MessageBox.Show($"Đã gửi lệnh in cho {count:N0} sản phẩm.", "In thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Không thể in", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
            btnPrint.Text = "In tem";
            UpdateActions();
        }
    }

    private bool EnsureReady()
    {
        if (products.Count == 0 || selectedLabel == null)
        {
            MessageBox.Show("Hãy chọn file Excel và loại tem trước.", "Chưa đủ thông tin");
            return false;
        }
        if (selectedLabel.RequiresEmployeeCode && string.IsNullOrWhiteSpace(txtEmployee.Text))
        {
            MessageBox.Show("Hãy nhập mã nhân viên.", "Thiếu mã nhân viên");
            txtEmployee.Focus();
            return false;
        }
        return true;
    }

    private void RestoreLastFile()
    {
        var last = ConfigService.Instance.Config.LastExcelFile;
        if (File.Exists(last)) LoadExcel(last);
        else
        {
            lblFileState.Text = "Có thể kéo-thả file vào cửa sổ";
            lblEmployee.Visible = false;
            txtEmployee.Visible = false;
        }
    }

    private void OnDragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            e.Effect = DragDropEffects.Copy;
    }

    private void OnDragDrop(object? sender, DragEventArgs e)
    {
        var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
        var file = files?.FirstOrDefault(x => x.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
        if (file != null) LoadExcel(file);
    }
}
