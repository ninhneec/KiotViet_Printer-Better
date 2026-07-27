using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using KiotVietLabelPrinter.Ui;

namespace KiotVietLabelPrinter.Forms;

public class FormMain : Form
{
    private readonly LabelService labelService = new();
    private readonly ExcelService excelService = new();
    private readonly LabelCatalogService catalogService = new();
    private readonly TextBox txtExcel = new();
    private readonly TextBox txtEmployee = new();
    private readonly DataGridView grid = new();
    private readonly FlowLayoutPanel templateList = new();
    private readonly Label lblFileState = new();
    private readonly Label lblTemplateState = new();
    private readonly Label lblSummary = new();
    private readonly Label lblEmptyData = new();
    private readonly Label lblEmployee = new();
    private readonly Button btnPreview = new();
    private readonly Button btnPrint = new();
    private readonly Label lblZoom = new();
    private readonly TrackBar zoomSlider = new();
    private readonly ToolTip fileToolTip = new();
    private readonly Dictionary<string, int> baseColumnWidths = new();
    private readonly Stack<CellEdit> undoStack = new();
    private object? valueBeforeEdit;
    private bool applyingUndo;
    private LabelDefinition? selectedLabel;
    private List<ProductRow> products = new();
    private string selectedExcelPath = "";
    private int zoomPercent = 100;
    private sealed record CellEdit(int RowIndex, int ColumnIndex, object? PreviousValue);

    public FormMain()
    {
        Text = "In tem KiotViet";
        MinimumSize = new Size(1100, 700);
        Size = new Size(1360, 860);
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
        var header = new Panel { Dock = DockStyle.Top, Height = 104, BackColor = AppTheme.Ink };
        header.Controls.Add(new Label
        {
            Text = "In tem KiotViet",
            Left = 30, Top = 17, Width = 360, Height = 40,
            Font = AppTheme.Display(23), ForeColor = Color.White
        });
        header.Controls.Add(new Label
        {
            Text = "Đưa file vào, kiểm tra nhanh rồi in — app tự xử lý file data BarTender.",
            Left = 32, Top = 60, Width = 650, Height = 24,
            Font = AppTheme.Body(10), ForeColor = Color.FromArgb(195, 211, 206)
        });

        var config = new Button { Text = "Quản lý mẫu tem", Width = 166, Height = 40, Top = 31 };
        config.Left = header.Width - 322;
        config.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleHeaderButton(config);
        config.Click += (_, _) =>
        {
            using var form = new FormConfig();
            form.ShowDialog(this);
            ReloadTemplates();
        };
        header.Controls.Add(config);

        var history = new Button { Text = "Lịch sử", Width = 110, Height = 40, Top = 25 };
        history.Top = 31;
        history.Left = header.Width - 136;
        history.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleHeaderButton(history);
        history.Click += (_, _) => { using var form = new FormHistory(); form.ShowDialog(this); };
        header.Controls.Add(history);
        return header;
    }

    private Control BuildBody()
    {
        var body = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24, 20, 24, 14),
            RowCount = 2,
            ColumnCount = 2,
            BackColor = AppTheme.Canvas
        };
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, 126));
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
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Padding = new Padding(22) };
        panel.Controls.Add(new Label
        {
            Text = "01",
            Left = 20, Top = 17, Width = 36, Height = 24,
            Font = AppTheme.Body(9F, FontStyle.Bold), ForeColor = AppTheme.AccentDark,
            BackColor = AppTheme.AccentSoft, TextAlign = ContentAlignment.MiddleCenter
        });
        panel.Controls.Add(new Label
        {
            Text = "Chọn bảng giá KiotViet",
            Left = 68, Top = 14, Width = 360, Height = 30,
            Font = AppTheme.Display(14), ForeColor = AppTheme.Ink
        });

        txtExcel.SetBounds(22, 60, 720, 38);
        txtExcel.ReadOnly = true;
        txtExcel.BackColor = Color.White;
        txtExcel.PlaceholderText = "Kéo file vào đây hoặc bấm Chọn file Excel";
        panel.Controls.Add(txtExcel);

        var choose = new Button { Text = "Chọn file Excel", Width = 145, Height = 38, Left = 758, Top = 60 };
        AppTheme.StylePrimary(choose);
        choose.Click += (_, _) => ChooseExcel();
        panel.Controls.Add(choose);

        lblFileState.SetBounds(922, 63, 330, 34);
        lblFileState.ForeColor = AppTheme.Muted;
        lblFileState.Font = AppTheme.Body(9.5F, FontStyle.Bold);
        panel.Controls.Add(lblFileState);
        return panel;
    }

    private Control BuildDataPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Margin = new Padding(0, 14, 10, 0), Padding = new Padding(20) };
        panel.Controls.Add(new Label
        {
            Text = "02  Kiểm tra và sửa dữ liệu",
            Dock = DockStyle.Top, Height = 34,
            Font = AppTheme.Display(14), ForeColor = AppTheme.Ink
        });
        lblSummary.Dock = DockStyle.Top;
        lblSummary.Height = 28;
        lblSummary.ForeColor = AppTheme.Muted;
        lblSummary.Text = "Dữ liệu sản phẩm sẽ hiện ở đây.";
        panel.Controls.Add(lblSummary);
        lblSummary.BringToFront();

        var commandBar = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 42,
            ColumnCount = 2,
            Padding = new Padding(0, 3, 0, 3)
        };
        commandBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        commandBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        var editActions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
        var zoomBar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var zoomIn = new Button { Text = "+", Width = 36, Height = 30 };
        var zoomOut = new Button { Text = "−", Width = 36, Height = 30 };
        zoomSlider.Minimum = 50;
        zoomSlider.Maximum = 200;
        zoomSlider.TickFrequency = 25;
        zoomSlider.SmallChange = 10;
        zoomSlider.LargeChange = 25;
        zoomSlider.Value = 100;
        zoomSlider.Width = 150;
        zoomSlider.Height = 34;
        zoomSlider.TickStyle = TickStyle.None;
        zoomSlider.ValueChanged += (_, _) => SetZoom(zoomSlider.Value);
        AppTheme.StyleSecondary(zoomIn);
        AppTheme.StyleSecondary(zoomOut);
        lblZoom.Text = "100%";
        lblZoom.Width = 58;
        lblZoom.Height = 30;
        lblZoom.TextAlign = ContentAlignment.MiddleCenter;
        zoomIn.Click += (_, _) => SetZoom(zoomPercent + 10);
        zoomOut.Click += (_, _) => SetZoom(zoomPercent - 10);
        lblZoom.Click += (_, _) => SetZoom(100);
        zoomBar.Controls.Add(zoomIn);
        zoomBar.Controls.Add(lblZoom);
        zoomBar.Controls.Add(zoomSlider);
        zoomBar.Controls.Add(zoomOut);
        zoomBar.Controls.Add(new Label
        {
            Text = "Thu phóng",
            Width = 82,
            Height = 30,
            TextAlign = ContentAlignment.MiddleRight,
            ForeColor = AppTheme.Muted
        });
        var saveExcel = new Button { Text = "Lưu Excel mới", Width = 118, Height = 30 };
        var deleteRows = new Button { Text = "Xóa dòng", Width = 92, Height = 30 };
        AppTheme.StyleSecondary(saveExcel);
        AppTheme.StyleSecondary(deleteRows);
        saveExcel.Click += (_, _) => SaveEditedExcel();
        deleteRows.Click += (_, _) => DeleteSelectedRows();
        editActions.Controls.Add(saveExcel);
        editActions.Controls.Add(deleteRows);
        commandBar.Controls.Add(editActions, 0, 0);
        commandBar.Controls.Add(zoomBar, 1, 0);
        panel.Controls.Add(commandBar);
        commandBar.BringToFront();

        grid.Dock = DockStyle.Fill;
        grid.ReadOnly = false;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        grid.MultiSelect = true;
        grid.BackgroundColor = AppTheme.Surface;
        grid.BorderStyle = BorderStyle.None;
        grid.GridColor = AppTheme.Border;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        grid.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.SurfaceMuted;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.Ink;
        grid.ColumnHeadersDefaultCellStyle.Font = AppTheme.Body(9F, FontStyle.Bold);
        grid.EnableHeadersVisualStyles = false;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(218, 238, 231);
        grid.DefaultCellStyle.SelectionForeColor = AppTheme.Ink;
        grid.CellValidating += Grid_CellValidating;
        grid.CellBeginEdit += (_, e) => valueBeforeEdit = grid[e.ColumnIndex, e.RowIndex].Value;
        grid.CellEndEdit += Grid_CellEndEdit;
        grid.MouseWheel += Grid_MouseWheel;
        grid.KeyDown += Grid_KeyDown;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        grid.DataBindingComplete += (_, _) => RefreshRowNumbers();
        grid.RowHeadersVisible = true;
        grid.RowHeadersWidth = 54;
        panel.Controls.Add(grid);
        grid.BringToFront();
        lblEmptyData.Text = "Chưa có dữ liệu\n\nKéo file Excel KiotViet vào cửa sổ để bắt đầu.";
        lblEmptyData.TextAlign = ContentAlignment.MiddleCenter;
        lblEmptyData.Font = AppTheme.Body(11F, FontStyle.Bold);
        lblEmptyData.ForeColor = AppTheme.Muted;
        lblEmptyData.BackColor = AppTheme.Surface;
        lblEmptyData.Dock = DockStyle.Fill;
        panel.Controls.Add(lblEmptyData);
        lblEmptyData.BringToFront();
        return panel;
    }

    private Control BuildTemplatePanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = AppTheme.Surface, Margin = new Padding(10, 14, 0, 0), Padding = new Padding(20) };
        panel.Controls.Add(new Label
        {
            Text = "03  Chọn mẫu tem",
            Dock = DockStyle.Top, Height = 34,
            Font = AppTheme.Display(14), ForeColor = AppTheme.Ink
        });
        lblTemplateState.Dock = DockStyle.Bottom;
        lblTemplateState.Height = 48;
        lblTemplateState.Padding = new Padding(10);
        lblTemplateState.Font = AppTheme.Body(9F, FontStyle.Bold);
        panel.Controls.Add(lblTemplateState);

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
        var footer = new Panel { Dock = DockStyle.Bottom, Height = 94, BackColor = AppTheme.Surface, Padding = new Padding(24) };
        btnPrint.Text = "In tem ngay";
        btnPrint.Width = 190;
        btnPrint.Height = 48;
        btnPrint.Left = footer.Width - 216;
        btnPrint.Top = 22;
        btnPrint.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StylePrimary(btnPrint);
        btnPrint.Enabled = false;
        btnPrint.Click += Print_Click;
        footer.Controls.Add(btnPrint);

        btnPreview.Text = "Xem trước bản in";
        btnPreview.Width = 176;
        btnPreview.Height = 48;
        btnPreview.Left = btnPrint.Left - 190;
        btnPreview.Top = 22;
        btnPreview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(btnPreview);
        btnPreview.Enabled = false;
        btnPreview.Click += Preview_Click;
        footer.Controls.Add(btnPreview);

        footer.Controls.Add(new Label
        {
            Text = "File KiotViet gốc luôn được giữ nguyên. App chỉ cập nhật file data trung gian.",
            Left = 26, Top = 35, Width = 620, Height = 24,
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
            selectedExcelPath = path;
            txtExcel.Text = Path.GetFileName(path);
            fileToolTip.SetToolTip(txtExcel, path);
            grid.DataSource = products;
            FormatGrid();
            lblEmptyData.Visible = false;
            lblFileState.Text = $"✓ Đã đọc {products.Count:N0} sản phẩm";
            lblFileState.ForeColor = AppTheme.AccentDark;
            UpdateSummary();
            ConfigService.Instance.Config.LastExcelFile = path;
            ConfigService.Instance.Config.LastFolder = Path.GetDirectoryName(path) ?? "";
            ConfigService.Instance.Save();
            UpdateActions();
        }
        catch (Exception ex)
        {
            products.Clear();
            selectedExcelPath = "";
            grid.DataSource = null;
            lblEmptyData.Visible = true;
            lblFileState.Text = "Không đọc được file";
            lblFileState.ForeColor = AppTheme.Danger;
            MessageBox.Show(ex.Message, "Không đọc được dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            UpdateActions();
        }
    }

    private void FormatGrid()
    {
        foreach (DataGridViewColumn column in grid.Columns)
            column.Visible = false;

        RenameColumn("ProductCode", "Mã hàng", 130);
        RenameColumn("ProductNameWithAttr", "Tên in trên tem", 360);
        RenameColumn("Unit", "Đơn vị tính", 110);
        RenameColumn("Quantity", "Số lượng", 100);
        RenameColumn("Price", "Giá bán", 130);

        foreach (DataGridViewColumn column in grid.Columns)
            column.ReadOnly = true;

        SetEditable("ProductNameWithAttr");
        SetEditable("Unit");
        SetEditable("Quantity");
        SetEditable("Price");

        if (grid.Columns["Price"] is { } price) price.DefaultCellStyle.Format = "N0";
        baseColumnWidths.Clear();
        foreach (DataGridViewColumn column in grid.Columns.Cast<DataGridViewColumn>().Where(column => column.Visible))
            baseColumnWidths[column.Name] = column.Width;
        SetZoom(zoomPercent);
    }

    private void RenameColumn(string name, string text, int width)
    {
        if (grid.Columns[name] is not { } column) return;
        column.Visible = true;
        column.HeaderText = text;
        column.Width = width;
    }

    private void SetEditable(string columnName)
    {
        if (grid.Columns[columnName] is not { } column) return;
        column.ReadOnly = false;
        column.DefaultCellStyle.BackColor = Color.FromArgb(255, 249, 218);
    }

    private void Grid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
    {
        string? columnName = grid.Columns[e.ColumnIndex].Name;
        if (columnName is not ("Quantity" or "Price"))
            return;

        if (!double.TryParse(e.FormattedValue?.ToString(), out double value) || value < 0)
        {
            e.Cancel = true;
            MessageBox.Show(
                columnName == "Price"
                    ? "Giá bán phải là một số từ 0 trở lên."
                    : "Số lượng phải là một số từ 0 trở lên.",
                "Dữ liệu chưa đúng",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private void Grid_MouseWheel(object? sender, MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Control) != Keys.Control)
            return;

        if (e is HandledMouseEventArgs handled)
            handled.Handled = true;
        SetZoom(zoomPercent + (e.Delta > 0 ? 10 : -10));
    }

    private void SetZoom(int percent)
    {
        zoomPercent = Math.Clamp(percent, 50, 200);
        lblZoom.Text = $"{zoomPercent}%";
        if (zoomSlider.Value != zoomPercent)
            zoomSlider.Value = zoomPercent;
        float scale = zoomPercent / 100F;
        grid.DefaultCellStyle.Font = AppTheme.Body(9.5F * scale);
        grid.ColumnHeadersDefaultCellStyle.Font = AppTheme.Body(9F * scale, FontStyle.Bold);
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        grid.ColumnHeadersHeight = Math.Max(28, (int)(34 * scale));
        grid.RowTemplate.Height = Math.Max(24, (int)(28 * scale));
        foreach ((string name, int width) in baseColumnWidths)
        {
            if (grid.Columns[name] is { } column)
                column.Width = Math.Max(55, (int)(width * scale));
        }
        foreach (DataGridViewRow row in grid.Rows)
            row.Height = grid.RowTemplate.Height;
    }

    private void RefreshRowNumbers()
    {
        for (int index = 0; index < grid.Rows.Count; index++)
            grid.Rows[index].HeaderCell.Value = (index + 1).ToString();
    }

    private void Grid_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.V)
        {
            PasteClipboard();
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }

        if (e.Control && e.KeyCode == Keys.Z)
        {
            UndoLastEdit();
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }

        if (e.Control && e.KeyCode == Keys.D)
        {
            FillSelection(down: true);
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }

        if (e.Control && e.KeyCode == Keys.R)
        {
            FillSelection(down: false);
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }

        if (e.Control && e.KeyCode == Keys.S)
        {
            SaveEditedExcel();
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }

        if (e.KeyCode == Keys.Delete)
        {
            ClearSelectedCells();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (!grid.ContainsFocus)
            return base.ProcessCmdKey(ref msg, keyData);

        if (!grid.IsCurrentCellInEditMode &&
            keyData is Keys.Left or Keys.Right or Keys.Up or Keys.Down)
        {
            MoveCurrentCell(keyData);
            return true;
        }

        if (keyData is Keys.Enter or (Keys.Shift | Keys.Enter))
        {
            int direction = keyData == Keys.Enter ? 1 : -1;
            int row = grid.CurrentCell?.RowIndex ?? 0;
            int column = grid.CurrentCell?.ColumnIndex ?? 0;
            grid.EndEdit();
            int nextRow = Math.Clamp(row + direction, 0, Math.Max(0, grid.Rows.Count - 1));
            if (grid.Rows.Count > 0)
                grid.CurrentCell = grid[column, nextRow];
            return true;
        }

        if (keyData == Keys.F2 && grid.CurrentCell is { ReadOnly: false })
        {
            grid.BeginEdit(true);
            return true;
        }

        if (keyData == (Keys.Control | Keys.D0) || keyData == (Keys.Control | Keys.NumPad0))
        {
            SetZoom(100);
            return true;
        }

        if (keyData is (Keys.Control | Keys.Add) or (Keys.Control | Keys.Oemplus))
        {
            SetZoom(zoomPercent + 10);
            return true;
        }

        if (keyData is (Keys.Control | Keys.Subtract) or (Keys.Control | Keys.OemMinus))
        {
            SetZoom(zoomPercent - 10);
            return true;
        }

        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void MoveCurrentCell(Keys direction)
    {
        if (grid.CurrentCell == null || grid.Rows.Count == 0)
            return;

        int row = grid.CurrentCell.RowIndex;
        int column = grid.CurrentCell.ColumnIndex;

        if (direction == Keys.Up)
            row = Math.Max(0, row - 1);
        else if (direction == Keys.Down)
            row = Math.Min(grid.Rows.Count - 1, row + 1);
        else
        {
            int step = direction == Keys.Left ? -1 : 1;
            int candidate = column + step;
            while (candidate >= 0 && candidate < grid.Columns.Count &&
                   !grid.Columns[candidate].Visible)
            {
                candidate += step;
            }

            if (candidate >= 0 && candidate < grid.Columns.Count)
                column = candidate;
        }

        grid.CurrentCell = grid[column, row];
    }

    private void Grid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
    {
        object? newValue = grid[e.ColumnIndex, e.RowIndex].Value;
        if (!applyingUndo && !Equals(valueBeforeEdit, newValue))
            undoStack.Push(new CellEdit(e.RowIndex, e.ColumnIndex, valueBeforeEdit));
        valueBeforeEdit = null;
        UpdateSummary();
    }

    private void RememberCell(DataGridViewCell cell)
    {
        if (!applyingUndo)
            undoStack.Push(new CellEdit(cell.RowIndex, cell.ColumnIndex, cell.Value));
    }

    private void UndoLastEdit()
    {
        if (!undoStack.TryPop(out CellEdit? edit))
            return;
        if (edit.RowIndex >= grid.Rows.Count || edit.ColumnIndex >= grid.Columns.Count)
            return;

        applyingUndo = true;
        try
        {
            grid[edit.ColumnIndex, edit.RowIndex].Value = edit.PreviousValue;
            grid.EndEdit();
            grid.CurrentCell = grid[edit.ColumnIndex, edit.RowIndex];
        }
        finally
        {
            applyingUndo = false;
        }
        UpdateSummary();
    }

    private void ClearSelectedCells()
    {
        foreach (DataGridViewCell cell in grid.SelectedCells.Cast<DataGridViewCell>())
        {
            if (cell.ReadOnly)
                continue;
            RememberCell(cell);
            string name = grid.Columns[cell.ColumnIndex].Name;
            cell.Value = name is "Quantity" or "Price" ? 0D : "";
        }
        grid.EndEdit();
        UpdateSummary();
    }

    private void FillSelection(bool down)
    {
        List<DataGridViewCell> cells = grid.SelectedCells
            .Cast<DataGridViewCell>()
            .Where(cell => !cell.ReadOnly)
            .ToList();
        if (cells.Count < 2)
            return;

        if (down)
        {
            foreach (IGrouping<int, DataGridViewCell> column in cells.GroupBy(cell => cell.ColumnIndex))
            {
                DataGridViewCell source = column.OrderBy(cell => cell.RowIndex).First();
                foreach (DataGridViewCell target in column.Where(cell => cell.RowIndex != source.RowIndex))
                {
                    RememberCell(target);
                    target.Value = source.Value;
                }
            }
        }
        else
        {
            foreach (IGrouping<int, DataGridViewCell> row in cells.GroupBy(cell => cell.RowIndex))
            {
                DataGridViewCell source = row.OrderBy(cell => cell.ColumnIndex).First();
                foreach (DataGridViewCell target in row.Where(cell => cell.ColumnIndex != source.ColumnIndex))
                {
                    RememberCell(target);
                    target.Value = source.Value;
                }
            }
        }

        grid.EndEdit();
        UpdateSummary();
    }

    private void PasteClipboard()
    {
        if (grid.CurrentCell == null || !Clipboard.ContainsText())
            return;

        string[] rows = Clipboard.GetText().Replace("\r", "").Split('\n');
        int startRow = grid.CurrentCell.RowIndex;
        int startColumn = grid.CurrentCell.ColumnIndex;

        for (int rowOffset = 0; rowOffset < rows.Length; rowOffset++)
        {
            if (string.IsNullOrEmpty(rows[rowOffset]) || startRow + rowOffset >= grid.Rows.Count)
                continue;

            string[] values = rows[rowOffset].Split('\t');
            for (int columnOffset = 0; columnOffset < values.Length; columnOffset++)
            {
                int targetColumn = startColumn + columnOffset;
                if (targetColumn >= grid.Columns.Count)
                    break;

                DataGridViewCell cell = grid[targetColumn, startRow + rowOffset];
                if (cell.ReadOnly || !grid.Columns[targetColumn].Visible)
                    continue;

                string columnName = grid.Columns[targetColumn].Name;
                if (columnName is "Quantity" or "Price")
                {
                    if (double.TryParse(values[columnOffset], out double number) && number >= 0)
                    {
                        RememberCell(cell);
                        cell.Value = number;
                    }
                }
                else
                {
                    RememberCell(cell);
                    cell.Value = values[columnOffset];
                }
            }
        }

        grid.EndEdit();
        UpdateSummary();
    }

    private void DeleteSelectedRows()
    {
        List<ProductRow> selected = grid.SelectedCells
            .Cast<DataGridViewCell>()
            .Select(cell => grid.Rows[cell.RowIndex].DataBoundItem)
            .OfType<ProductRow>()
            .Distinct()
            .ToList();

        if (selected.Count == 0)
            return;

        foreach (ProductRow item in selected)
            products.Remove(item);

        grid.DataSource = null;
        grid.DataSource = products;
        FormatGrid();
        UpdateSummary();
    }

    private void SaveEditedExcel()
    {
        if (products.Count == 0)
            return;

        using SaveFileDialog dialog = new()
        {
            Filter = "Excel Workbook|*.xlsx",
            FileName = $"DuLieuInTem_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
            Title = "Lưu dữ liệu đã sửa"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;

        grid.EndEdit();
        excelService.ExportProducts(dialog.FileName, products);
        MessageBox.Show(
            $"Đã lưu file:\n{dialog.FileName}",
            "Đã lưu dữ liệu",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void UpdateSummary()
    {
        lblSummary.Text =
            $"{products.Count:N0} sản phẩm • Tổng số lượng {products.Sum(x => x.Quantity):N0} • Ô màu vàng có thể sửa";
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
        else
        {
            List<Button> buttons = templateList.Controls.OfType<Button>().ToList();
            if (buttons.Count == 1)
            {
                LabelDefinition onlyTemplate = (LabelDefinition)buttons[0].Tag!;
                SelectTemplate(onlyTemplate, buttons[0]);
            }
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
            Enabled = true,
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
        UpdateActions();
    }

    private void UpdateActions()
    {
        List<string> templateIssues = selectedLabel?.GetReadinessIssues() ?? [];
        var ready = products.Count > 0 && selectedLabel != null && templateIssues.Count == 0;
        btnPreview.Enabled = ready;
        btnPrint.Enabled = ready;
        lblTemplateState.Text = selectedLabel == null
            ? "Chọn một mẫu để tiếp tục."
            : templateIssues.Count > 0
                ? "Cần sửa mẫu: " + string.Join(" • ", templateIssues)
                : products.Count == 0
                    ? $"✓ {selectedLabel.Name} đã sẵn sàng. Hãy chọn file Excel."
                    : $"✓ Sẽ in bằng: {selectedLabel.Name}";
        lblTemplateState.BackColor = ready ? Color.FromArgb(224, 243, 235) : AppTheme.SurfaceMuted;
        lblTemplateState.ForeColor = ready ? AppTheme.AccentDark : AppTheme.Muted;
    }

    private void Preview_Click(object? sender, EventArgs e)
    {
        if (!EnsureReady()) return;
        grid.EndEdit();
        using var form = new FormPreview(
            selectedExcelPath,
            selectedLabel!.Code,
            "",
            products.Select(CloneProduct).ToList());
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
            grid.EndEdit();
            var printProducts = products.Select(CloneProduct).ToList();
            var count = await Task.Run(() => labelService.PrintProducts(
                printProducts,
                selectedExcelPath,
                selectedLabel.Code,
                ""));
            MessageBox.Show($"Đã gửi lệnh in cho {count:N0} sản phẩm.", "In thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Không thể in", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
            btnPrint.Text = "In tem ngay";
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
        return true;
    }

    private void RestoreLastFile()
    {
        var last = ConfigService.Instance.Config.LastExcelFile;
        if (File.Exists(last)) LoadExcel(last);
        else
        {
            lblFileState.Text = "Có thể kéo-thả file vào cửa sổ";
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

    private static ProductRow CloneProduct(ProductRow item) => new()
    {
        StoreName = item.StoreName,
        Category = item.Category,
        ProductCode = item.ProductCode,
        Barcode = item.Barcode,
        ProductName = item.ProductName,
        ProductNameWithAttr = item.ProductNameWithAttr,
        Unit = item.Unit,
        Quantity = item.Quantity,
        Price = item.Price,
        Description = item.Description,
        Attribute = item.Attribute,
        Attribute2 = item.Attribute2,
        Position = item.Position
    };
}
