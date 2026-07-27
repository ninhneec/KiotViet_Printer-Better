using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using KiotVietLabelPrinter.Ui;
using System.Globalization;
using System.Text;

namespace KiotVietLabelPrinter.Forms;

public class FormFlowWizard : Form
{
    private readonly DataFlowService service = new();
    private readonly TextBox txtName = new();
    private readonly TextBox txtFile1 = new();
    private readonly TextBox txtFile2 = new();
    private readonly ComboBox cmbKey1 = NewCombo();
    private readonly ComboBox cmbKey2 = NewCombo();
    private readonly ComboBox cmbName = NewCombo();
    private readonly ComboBox cmbPrice = NewCombo();
    private readonly ComboBox cmbUnit = NewCombo();
    private readonly Label lblFile1 = StateLabel();
    private readonly Label lblFile2 = StateLabel();
    private readonly Label lblJoin = StateLabel();
    private readonly Label lblReady = StateLabel();
    private FlowTable? table1;
    private FlowTable? table2;

    public FormFlowWizard()
    {
        Text = "Tạo cách ghép dữ liệu";
        Size = new Size(1120, 790);
        MinimumSize = new Size(980, 700);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = AppTheme.Canvas;
        Font = AppTheme.Body();
        BuildUi();
    }

    private void BuildUi()
    {
        Controls.Add(BuildFooter());
        Controls.Add(BuildBody());
        Controls.Add(BuildHeader());
    }

    private Control BuildHeader()
    {
        Panel header = new() { Dock = DockStyle.Top, Height = 105, BackColor = AppTheme.Ink };
        header.Controls.Add(new Label
        {
            Text = "Tạo cách ghép dữ liệu",
            Left = 30, Top = 15, Width = 520, Height = 40,
            Font = AppTheme.Display(23), ForeColor = Color.White
        });
        header.Controls.Add(new Label
        {
            Text = "Chỉ cần chọn file. App sẽ đọc tiêu đề và gợi ý phần còn lại.",
            Left = 32, Top = 59, Width = 650, Height = 25,
            Font = AppTheme.Body(10), ForeColor = Color.FromArgb(195, 211, 206)
        });
        Button advanced = new() { Text = "Chế độ nâng cao", Width = 155, Height = 40, Top = 30 };
        advanced.Left = header.Width - 185;
        advanced.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleHeaderButton(advanced);
        advanced.Click += (_, _) =>
        {
            using var form = new FormFlowDesigner();
            form.ShowDialog(this);
        };
        header.Controls.Add(advanced);
        return header;
    }

    private Control BuildBody()
    {
        Panel scroll = new()
        {
            Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(28, 20, 28, 30),
            BackColor = AppTheme.Canvas
        };
        TableLayoutPanel layout = new()
        {
            Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 4
        };
        layout.Controls.Add(BuildFileStep(), 0, 0);
        layout.Controls.Add(BuildJoinStep(), 0, 1);
        layout.Controls.Add(BuildOutputStep(), 0, 2);
        layout.Controls.Add(BuildReadyStep(), 0, 3);
        scroll.Controls.Add(layout);
        return scroll;
    }

    private Control BuildFileStep()
    {
        Panel card = StepCard(170, "1", "Chọn dữ liệu", "File đầu tiên là bảng sản phẩm. File thứ hai chỉ dùng khi cần bổ sung thông tin.");
        AddFileRow(card, "File sản phẩm", txtFile1, lblFile1, 78, () => ChooseFile(1));
        AddFileRow(card, "File bổ sung (không bắt buộc)", txtFile2, lblFile2, 122, () => ChooseFile(2));
        Button clearFile2 = new() { Text = "Bỏ", Left = 892, Top = 122, Width = 52, Height = 32 };
        AppTheme.StyleSecondary(clearFile2);
        clearFile2.Click += (_, _) =>
        {
            txtFile2.Clear();
            table2 = null;
            cmbKey2.Items.Clear();
            lblFile2.Text = "Không dùng";
            lblFile2.ForeColor = AppTheme.Muted;
            RefreshOutputColumns();
            UpdateJoinState();
        };
        card.Controls.Add(clearFile2);
        return card;
    }

    private Control BuildJoinStep()
    {
        Panel card = StepCard(145, "2", "App nối hai file bằng gì?", "Chọn cột có cùng giá trị ở cả hai file, thường là Mã hàng hoặc Barcode.");
        AddCombo(card, "File sản phẩm", cmbKey1, 78, 160);
        Label arrow = new()
        {
            Text = "nối với", Left = 450, Top = 86, Width = 70, Height = 28,
            TextAlign = ContentAlignment.MiddleCenter, ForeColor = AppTheme.Muted
        };
        card.Controls.Add(arrow);
        AddCombo(card, "File bổ sung", cmbKey2, 78, 530);
        lblJoin.SetBounds(160, 112, 700, 24);
        card.Controls.Add(lblJoin);
        cmbKey1.SelectedIndexChanged += (_, _) => UpdateJoinState();
        cmbKey2.SelectedIndexChanged += (_, _) => UpdateJoinState();
        return card;
    }

    private Control BuildOutputStep()
    {
        Panel card = StepCard(160, "3", "Chọn nội dung sẽ in", "App chỉ tạo ba trường ổn định để mẫu BarTender luôn hiểu.");
        AddCombo(card, "Tên hàng", cmbName, 78, 160);
        AddCombo(card, "Giá bán", cmbPrice, 78, 465);
        AddCombo(card, "Đơn vị tính", cmbUnit, 78, 770);
        foreach (ComboBox combo in new[] { cmbName, cmbPrice, cmbUnit })
            combo.SelectedIndexChanged += (_, _) => UpdateReadyState();
        return card;
    }

    private Control BuildReadyStep()
    {
        Panel card = StepCard(132, "4", "Đặt tên và kiểm tra", "Tên dễ nhớ, ví dụ: Tem giá cửa hàng hoặc Tem phụ kiện.");
        Label label = new()
        {
            Text = "Tên cách ghép", Left = 160, Top = 72, Width = 120, Height = 24,
            ForeColor = AppTheme.Muted
        };
        txtName.SetBounds(280, 68, 330, 32);
        txtName.Text = "Tem giá sản phẩm";
        txtName.TextChanged += (_, _) => UpdateReadyState();
        lblReady.SetBounds(630, 72, 350, 25);
        card.Controls.AddRange([label, txtName, lblReady]);
        return card;
    }

    private Control BuildFooter()
    {
        Panel footer = new() { Dock = DockStyle.Bottom, Height = 82, BackColor = AppTheme.Surface };
        Label hint = new()
        {
            Text = "Không cần kéo dây. Sau khi lưu, app sẽ nhớ toàn bộ cách ghép.",
            Left = 28, Top = 31, Width = 530, Height = 24, ForeColor = AppTheme.Muted
        };
        Button preview = new() { Text = "Xem kết quả", Width = 145, Height = 44, Top = 19 };
        Button save = new() { Text = "Lưu cách ghép", Width = 155, Height = 44, Top = 19 };
        Button close = new() { Text = "Đóng", Width = 90, Height = 44, Top = 19 };
        close.Left = footer.Width - 115;
        save.Left = close.Left - 170;
        preview.Left = save.Left - 160;
        close.Anchor = save.Anchor = preview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        AppTheme.StyleSecondary(close);
        AppTheme.StylePrimary(save);
        AppTheme.StyleSecondary(preview);
        close.Click += (_, _) => Close();
        preview.Click += (_, _) => Preview();
        save.Click += (_, _) => SaveFlow();
        footer.Controls.AddRange([hint, preview, save, close]);
        return footer;
    }

    private static Panel StepCard(int height, string number, string title, string description)
    {
        Panel card = new()
        {
            Dock = DockStyle.Top, Height = height, BackColor = AppTheme.Surface,
            Margin = new Padding(0, 0, 0, 14)
        };
        Label badge = new()
        {
            Text = number, Left = 22, Top = 22, Width = 48, Height = 48,
            BackColor = AppTheme.AccentSoft, ForeColor = AppTheme.AccentDark,
            TextAlign = ContentAlignment.MiddleCenter, Font = AppTheme.Display(16)
        };
        card.Controls.Add(badge);
        card.Controls.Add(new Label
        {
            Text = title, Left = 88, Top = 20, Width = 500, Height = 30,
            Font = AppTheme.Display(15), ForeColor = AppTheme.Ink
        });
        card.Controls.Add(new Label
        {
            Text = description, Left = 89, Top = 50, Width = 800, Height = 25,
            Font = AppTheme.Body(9), ForeColor = AppTheme.Muted
        });
        return card;
    }

    private void AddFileRow(Panel card, string labelText, TextBox text, Label state, int top, Action choose)
    {
        Label label = new()
        {
            Text = labelText, Left = 88, Top = top + 5, Width = 210, Height = 25,
            ForeColor = AppTheme.Ink, Font = AppTheme.Body(9, FontStyle.Bold)
        };
        text.SetBounds(300, top, 470, 32);
        text.ReadOnly = true;
        text.BackColor = Color.White;
        Button button = new() { Text = "Chọn file", Left = 782, Top = top, Width = 105, Height = 32 };
        AppTheme.StyleSecondary(button);
        button.Click += (_, _) => choose();
        state.SetBounds(950, top + 5, 100, 24);
        card.Controls.AddRange([label, text, button, state]);
    }

    private static void AddCombo(Panel card, string labelText, ComboBox combo, int top, int left)
    {
        Label label = new()
        {
            Text = labelText, Left = left, Top = top, Width = 190, Height = 20,
            ForeColor = AppTheme.Muted, Font = AppTheme.Body(8.5F)
        };
        combo.SetBounds(left, top + 23, 235, 32);
        card.Controls.AddRange([label, combo]);
    }

    private void ChooseFile(int source)
    {
        using OpenFileDialog dialog = new()
        {
            Filter = "Bảng dữ liệu|*.xls;*.xlsx;*.csv|Tất cả file|*.*",
            Title = source == 1 ? "Chọn file sản phẩm" : "Chọn file bổ sung"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            FlowTable table = service.ReadSource(dialog.FileName);
            if (source == 1)
            {
                txtFile1.Text = dialog.FileName;
                table1 = table;
                lblFile1.Text = $"✓ {table.Rows.Count:N0} dòng";
                lblFile1.ForeColor = AppTheme.AccentDark;
                SetItems(cmbKey1, table.Columns);
            }
            else
            {
                txtFile2.Text = dialog.FileName;
                table2 = table;
                lblFile2.Text = $"✓ {table.Rows.Count:N0} dòng";
                lblFile2.ForeColor = AppTheme.AccentDark;
                SetItems(cmbKey2, table.Columns);
            }
            AutoChooseJoin();
            RefreshOutputColumns();
            UpdateReadyState();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Không đọc được file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void AutoChooseJoin()
    {
        if (table1 == null || table2 == null) return;
        string[][] preferredGroups =
        [
            ["Mã hàng", "Mã sản phẩm", "Mã SP", "SKU"],
            ["Barcode", "Mã vạch"]
        ];
        foreach (string[] group in preferredGroups)
        {
            string? left = group.Select(candidate => FindColumn(table1.Columns, candidate))
                .FirstOrDefault(column => column != null);
            string? right = group.Select(candidate => FindColumn(table2.Columns, candidate))
                .FirstOrDefault(column => column != null);
            if (left == null || right == null) continue;
            cmbKey1.SelectedItem = left;
            cmbKey2.SelectedItem = right;
            UpdateJoinState();
            return;
        }
        string? common = table1.Columns.FirstOrDefault(left =>
            table2.Columns.Any(right => Normalize(left) == Normalize(right)));
        if (common != null)
        {
            cmbKey1.SelectedItem = common;
            cmbKey2.SelectedItem = table2.Columns.First(right => Normalize(right) == Normalize(common));
        }
    }

    private void RefreshOutputColumns()
    {
        if (table1 == null) return;
        List<string> columns = [.. table1.Columns];
        if (table2 != null)
        {
            foreach (string column in table2.Columns)
                columns.Add(columns.Contains(column, StringComparer.OrdinalIgnoreCase)
                    ? $"{column} (File 2)"
                    : column);
        }
        SetItems(cmbName, columns);
        SetItems(cmbPrice, columns);
        SetItems(cmbUnit, columns);
        SelectSuggested(cmbName, ["Tên hàng thuộc tính", "Tên hàng", "Tên sản phẩm"]);
        SelectSuggested(cmbPrice, ["Giá bán", "Giá", "Đơn giá"]);
        SelectSuggested(cmbUnit, ["Đơn vị tính (File 2)", "Đơn vị tính", "ĐVT"]);
    }

    private void UpdateJoinState()
    {
        if (table2 == null)
        {
            lblJoin.Text = "Không có File 2 — app sẽ dùng trực tiếp File sản phẩm.";
            lblJoin.ForeColor = AppTheme.Muted;
        }
        else if (cmbKey1.SelectedItem != null && cmbKey2.SelectedItem != null)
        {
            HashSet<string> rightValues = table2.Rows
                .Select(row => Normalize(row.GetValueOrDefault(cmbKey2.Text, "")))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToHashSet();
            int matched = table1!.Rows.Count(row =>
                rightValues.Contains(Normalize(row.GetValueOrDefault(cmbKey1.Text, ""))));
            lblJoin.Text = matched > 0
                ? $"✓ Nối được {matched:N0}/{table1.Rows.Count:N0} dòng"
                : "Không tìm thấy giá trị trùng — hãy chọn lại hai cột";
            lblJoin.ForeColor = matched > 0 ? AppTheme.AccentDark : AppTheme.Danger;
        }
        UpdateReadyState();
    }

    private void UpdateReadyState()
    {
        bool ready = table1 != null &&
                     (table2 == null || (cmbKey1.SelectedItem != null && cmbKey2.SelectedItem != null)) &&
                     cmbName.SelectedItem != null && cmbPrice.SelectedItem != null &&
                     cmbUnit.SelectedItem != null && !string.IsNullOrWhiteSpace(txtName.Text);
        lblReady.Text = ready ? "✓ Có thể xem kết quả và lưu" : "Chọn đủ các mục ở phía trên";
        lblReady.ForeColor = ready ? AppTheme.AccentDark : AppTheme.Muted;
    }

    private DataFlowDefinition BuildFlow()
    {
        if (table1 == null || string.IsNullOrWhiteSpace(txtFile1.Text))
            throw new InvalidOperationException("Hãy chọn File sản phẩm ở bước 1.");
        if (table2 != null && (cmbKey1.SelectedItem == null || cmbKey2.SelectedItem == null))
            throw new InvalidOperationException("Hãy chọn cột dùng để nối hai file ở bước 2.");
        if (cmbName.SelectedItem == null || cmbPrice.SelectedItem == null || cmbUnit.SelectedItem == null)
            throw new InvalidOperationException("Hãy chọn đủ Tên hàng, Giá bán và Đơn vị tính ở bước 3.");
        return service.CreateSimpleFlow(
            txtName.Text,
            txtFile1.Text,
            table2 == null ? null : txtFile2.Text,
            cmbKey1.Text,
            cmbKey2.Text,
            cmbName.Text,
            cmbPrice.Text,
            cmbUnit.Text);
    }

    private void Preview()
    {
        try
        {
            DataFlowDefinition flow = BuildFlow();
            FlowTable result = service.Execute(flow);
            using FormFlowPreview preview = new(flow.Name, result);
            preview.ShowDialog(this);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Chưa xem được kết quả", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void SaveFlow()
    {
        try
        {
            DataFlowDefinition flow = BuildFlow();
            FlowTable result = service.Execute(flow);
            service.Save(flow);
            MessageBox.Show(
                $"Đã lưu “{flow.Name}”.\n\nKết quả thử: {result.Rows.Count:N0} dòng.",
                "Đã lưu cách ghép",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Chưa lưu được", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static ComboBox NewCombo() => new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        IntegralHeight = false,
        MaxDropDownItems = 12
    };

    private static Label StateLabel() => new()
    {
        ForeColor = AppTheme.Muted,
        Font = AppTheme.Body(8.5F, FontStyle.Bold)
    };

    private static void SetItems(ComboBox combo, IEnumerable<string> items)
    {
        combo.BeginUpdate();
        combo.Items.Clear();
        combo.Items.AddRange(items.Cast<object>().ToArray());
        combo.EndUpdate();
    }

    private static void SelectSuggested(ComboBox combo, IEnumerable<string> candidates)
    {
        foreach (string candidate in candidates)
        {
            object? match = combo.Items.Cast<object>().FirstOrDefault(item =>
                Normalize(item.ToString() ?? "") == Normalize(candidate));
            if (match == null) continue;
            combo.SelectedItem = match;
            return;
        }
    }

    private static string? FindColumn(IEnumerable<string> columns, string wanted) =>
        columns.FirstOrDefault(column => Normalize(column) == Normalize(wanted));

    private static string Normalize(string value)
    {
        string decomposed = value.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
        StringBuilder result = new();
        foreach (char character in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
                result.Append(character == 'đ' ? 'd' : character);
        }
        return result.ToString().Normalize(NormalizationForm.FormC);
    }
}
