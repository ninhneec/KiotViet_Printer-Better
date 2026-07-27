using System.Reflection;
using KiotVietLabelPrinter.Forms;
using KiotVietLabelPrinter.Models;
using KiotVietLabelPrinter.Services;
using NPOI.SS.UserModel;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
                throw new Exception("Cần đường dẫn Excel mẫu và file .btw mẫu.");

            string sourceExcel = Path.GetFullPath(args[0]);
            string sourceTemplate = Path.GetFullPath(args[1]);
            Assert(File.Exists(sourceExcel), "Không tìm thấy Excel mẫu.");
            Assert(File.Exists(sourceTemplate), "Không tìm thấy .btw mẫu.");

            string runtime = Path.Combine(
                Path.GetTempPath(),
                "KiotVietPrinterBetter-SmokeTests",
                DateTime.Now.ToString("yyyyMMdd-HHmmss-fff"));
            Directory.CreateDirectory(runtime);
            Environment.SetEnvironmentVariable("KIOTVIET_TEST_DATA_DIRECTORY", runtime);

            var excelService = new ExcelService();
            List<ProductRow> products = excelService.ReadProducts(sourceExcel);
            Assert(products.Count > 0, "Excel phải có ít nhất một sản phẩm.");
            Assert(!string.IsNullOrWhiteSpace(products[0].ProductName), "Thiếu Tên hàng.");

            string dataFile = Path.Combine(runtime, "fixed-data.xls");
            excelService.WriteDirectPriceDataFile(dataFile, products[0]);
            using (FileStream input = new(dataFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IWorkbook workbook = WorkbookFactory.Create(input))
            {
                ISheet sheet = workbook.GetSheetAt(0);
                IRow header = sheet.GetRow(0);
                Assert(header.LastCellNum == 3, "File data phải có đúng 3 cột.");
                Assert(header.GetCell(0).StringCellValue == "Tên hàng", "Sai cột Tên hàng.");
                Assert(header.GetCell(1).StringCellValue == "Giá bán", "Sai cột Giá bán.");
                Assert(header.GetCell(2).StringCellValue == "Đơn vị tính", "Sai cột Đơn vị tính.");
                Assert(sheet.LastRowNum == 1, "File data phải có đúng một dòng sản phẩm.");
            }

            ProductRow secondProduct = Clone(products[0]);
            secondProduct.ProductCode += "-2";
            secondProduct.ProductNameWithAttr += " dòng 2";
            string uiExcel = Path.Combine(runtime, "ui-data.xlsx");
            excelService.ExportProducts(uiExcel, [products[0], secondProduct]);

            var configService = ConfigService.Instance;
            string templateCopy = Path.Combine(runtime, "source-template.btw");
            File.Copy(sourceTemplate, templateCopy);
            string fakeBarTender = Path.Combine(runtime, "bartend.exe");
            File.Copy(sourceTemplate, fakeBarTender);
            configService.Config.BarTenderExe = fakeBarTender;
            configService.Config.LastExcelFile = uiExcel;
            configService.Config.LastFolder = Path.GetDirectoryName(sourceExcel) ?? "";
            configService.Config.Labels =
            [
                new LabelDefinition
                {
                    Code = "SMOKE",
                    Name = "Tem kiểm thử",
                    Description = "Ba cột",
                    TemplatePath = templateCopy,
                    DataFilePath = dataFile,
                    HandlerType = "DIRECT_PRICE",
                    IsEnabled = true
                }
            ];
            configService.Save();

            ApplicationConfiguration.Initialize();
            using (var main = new TestableFormMain())
            {
                main.Show();
                Application.DoEvents();
                List<ProductRow> loadedProducts = ReadField<List<ProductRow>>(main, "products");
                LabelDefinition? selectedLabel = ReadField<LabelDefinition?>(main, "selectedLabel");
                Button preview = ReadField<Button>(main, "btnPreview");
                Button print = ReadField<Button>(main, "btnPrint");
                DataGridView grid = ReadField<DataGridView>(main, "grid");
                Assert(loadedProducts.Count == 2, "Màn chính không nạp đúng Excel.");
                Assert(selectedLabel?.Code == "SMOKE", "Mẫu duy nhất không tự được chọn.");
                Assert(preview.Enabled, "Nút Xem trước phải được bật.");
                Assert(print.Enabled, "Nút In phải được bật.");

                grid.CurrentCell = grid.Rows[0].Cells["ProductNameWithAttr"];
                grid.Focus();
                Assert(main.SendKey(Keys.Enter), "Enter chưa được app xử lý.");
                Assert(grid.CurrentCell?.RowIndex == 1, "Enter không xuống dòng.");
                Assert(main.SendKey(Keys.Shift | Keys.Enter), "Shift+Enter chưa được app xử lý.");
                Assert(grid.CurrentCell?.RowIndex == 0, "Shift+Enter không đi lên.");
                int startingColumn = grid.CurrentCell!.ColumnIndex;
                main.SendKey(Keys.Right);
                Assert(grid.CurrentCell?.ColumnIndex > startingColumn, "Mũi tên phải không chuyển ô.");
                main.SendKey(Keys.Left);
                Assert(grid.CurrentCell?.ColumnIndex == startingColumn, "Mũi tên trái không trở lại ô.");
                main.SendKey(Keys.Down);
                Assert(grid.CurrentCell?.RowIndex == 1, "Mũi tên xuống không chuyển dòng.");
                main.SendKey(Keys.Up);
                Assert(grid.CurrentCell?.RowIndex == 0, "Mũi tên lên không trở lại dòng.");
                main.Close();
            }

            using (var config = new FormConfig())
            {
                ListBox list = ReadField<ListBox>(config, "lstLabels");
                TextBox name = ReadField<TextBox>(config, "txtName");
                TextBox template = ReadField<TextBox>(config, "txtTemplate");
                TextBox data = ReadField<TextBox>(config, "txtDataFile");
                object? selectedBefore = list.SelectedItem;
                string templateBefore = template.Text;
                string dataBefore = data.Text;

                name.Text = "Tên mới không nhảy";
                Assert(ReferenceEquals(selectedBefore, list.SelectedItem), "Selection bị nhảy khi sửa tên.");
                Assert(template.Text == templateBefore, "File .btw bị đổi khi sửa tên.");
                Assert(data.Text == dataBefore, "File data bị đổi khi sửa tên.");

                Invoke(config, "PersistConfig", false);
                LabelDefinition saved = ConfigService.Instance.Config.Labels.Single();
                Assert(saved.Name == "Tên mới không nhảy", "Tên mẫu không được lưu.");
                Assert(saved.DataFilePath == dataBefore, "Lưu sai file data đã chọn.");
                Assert(File.Exists(saved.TemplatePath), "Bản sao .btw được lưu không tồn tại.");
                Assert(((LabelDefinition?)list.SelectedItem)?.Code == "SMOKE", "Lưu xong không giữ mẫu đang chọn.");
            }

            Console.WriteLine("PASS: Excel import");
            Console.WriteLine("PASS: Fixed data file has exactly 3 columns");
            Console.WriteLine("PASS: Single template auto-selection");
            Console.WriteLine("PASS: Preview/print readiness");
            Console.WriteLine("PASS: Enter/Shift+Enter cell navigation");
            Console.WriteLine("PASS: Four-direction arrow cell navigation");
            Console.WriteLine("PASS: Template name editing remains stable");
            Console.WriteLine("PASS: Settings preserve selected data/template");
            Console.WriteLine($"Runtime: {runtime}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("FAIL: " + ex);
            return 1;
        }
    }

    private static T ReadField<T>(object instance, string name)
    {
        Type? type = instance.GetType();
        FieldInfo? field = null;
        while (type != null && field == null)
        {
            field = type.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
            type = type.BaseType;
        }
        if (field == null)
            throw new MissingFieldException(instance.GetType().Name, name);
        return (T)field.GetValue(instance)!;
    }

    private static void Invoke(object instance, string name, params object[] args)
    {
        MethodInfo method = instance.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)
            ?? throw new MissingMethodException(instance.GetType().Name, name);
        method.Invoke(instance, args);
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
    }

    private static ProductRow Clone(ProductRow item) => new()
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

    private sealed class TestableFormMain : FormMain
    {
        public bool SendKey(Keys keys)
        {
            Message message = new();
            return ProcessCmdKey(ref message, keys);
        }
    }
}
