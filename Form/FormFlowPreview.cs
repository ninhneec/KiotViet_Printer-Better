using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Ui;

namespace KiotVietLabelPrinter.Forms;

public class FormFlowPreview : Form
{
    public FormFlowPreview(string flowName, FlowTable table)
    {
        Text = $"Kết quả chạy thử - {flowName}";
        Size = new Size(1100, 680);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = AppTheme.Canvas;
        Label summary = new()
        {
            Dock = DockStyle.Top, Height = 66, Padding = new Padding(22, 18, 0, 0),
            Text = $"{table.Rows.Count:N0} dòng · {table.Columns.Count:N0} cột · Dòng đỏ là dữ liệu cần kiểm tra",
            Font = AppTheme.Display(13), ForeColor = AppTheme.Ink
        };
        DataGridView grid = new()
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            BackgroundColor = AppTheme.Surface,
            BorderStyle = BorderStyle.None,
            RowHeadersVisible = false
        };
        foreach (string column in table.Columns)
            grid.Columns.Add(column, column == "__Lỗi" ? "Cảnh báo" : column);
        foreach (Dictionary<string, string> row in table.Rows.Take(1000))
        {
            int index = grid.Rows.Add(table.Columns.Select(column => row.GetValueOrDefault(column, "")).ToArray());
            if (!string.IsNullOrWhiteSpace(row.GetValueOrDefault("__Lỗi", "")))
            {
                grid.Rows[index].DefaultCellStyle.BackColor = Color.FromArgb(255, 232, 229);
                grid.Rows[index].DefaultCellStyle.ForeColor = AppTheme.Danger;
            }
        }
        Controls.Add(grid);
        Controls.Add(summary);
    }
}
