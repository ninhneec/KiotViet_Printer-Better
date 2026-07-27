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

            var flowService = new DataFlowService();
            DataFlowDefinition flow = DataFlowService.CreateStarterFlow();
            List<DataFlowNode> sources = flow.Nodes
                .Where(node => node.Type == DataFlowNodeType.ExcelSource)
                .ToList();
            sources[0].Settings["FilePath"] = sourceExcel;
            sources[1].Settings["FilePath"] = sourceExcel;
            DataFlowNode selectNode = flow.Nodes.Single(node => node.Type == DataFlowNodeType.SelectColumns);
            selectNode.Settings["Mappings"] =
                "Tên hàng=Tên hàng;Giá bán=Giá bán;Đơn vị tính=Đơn vị tính (File 2)";
            FlowTable flowResult = flowService.Execute(flow);
            Assert(flowResult.Rows.Count == products.Count, "Flow nối hai file sai số dòng.");
            Assert(flowResult.Columns.Take(3).SequenceEqual(["Tên hàng", "Giá bán", "Đơn vị tính"]) &&
                   flowResult.Columns.Contains("__Lỗi"),
                "Flow chưa ánh xạ đúng ba cột đầu ra.");
            string savedFlow = flowService.Save(flow);
            Assert(File.Exists(savedFlow), "Flow chưa được lưu lâu dài.");
            Assert(flowService.Load(savedFlow)?.Nodes.Count == flow.Nodes.Count, "Không mở lại được Flow đã lưu.");

            string csvSource = Path.Combine(runtime, "flow-source.csv");
            File.WriteAllText(csvSource, "Mã hàng,Tên hàng\nSP01,\"Sản phẩm, có dấu phẩy\"");
            DataFlowDefinition csvFlow = DataFlowService.CreateStarterFlow();
            DataFlowNode csvNode = csvFlow.Nodes.First(node => node.Type == DataFlowNodeType.ExcelSource);
            csvNode.Settings["FilePath"] = csvSource;
            DataFlowNode secondCsvSource = csvFlow.Nodes
                .Where(node => node.Type == DataFlowNodeType.ExcelSource)
                .Skip(1).First();
            secondCsvSource.Settings["FilePath"] = csvSource;
            DataFlowNode csvSelect = csvFlow.Nodes.Single(node => node.Type == DataFlowNodeType.SelectColumns);
            csvSelect.Settings["Mappings"] = "Tên hàng=Tên hàng";
            DataFlowNode csvValidate = csvFlow.Nodes.Single(node => node.Type == DataFlowNodeType.Validate);
            csvValidate.Settings["Required"] = "Tên hàng";
            FlowTable csvResult = flowService.Execute(csvFlow);
            Assert(csvResult.Rows.Single()["Tên hàng"] == "Sản phẩm, có dấu phẩy",
                "Flow đọc CSV có dấu phẩy trong chuỗi chưa đúng.");

            string dataFile = args.Length >= 3
                ? Path.GetFullPath(args[2])
                : Path.Combine(runtime, "fixed-data.xls");
            excelService.WriteDirectPriceDataFile(dataFile, products[0]);
            string copiedName;
            double copiedPrice;
            string copiedUnit;
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
                IRow copiedRow = sheet.GetRow(1);
                copiedName = copiedRow.GetCell(0).ToString() ?? "";
                copiedPrice = copiedRow.GetCell(1).NumericCellValue;
                copiedUnit = copiedRow.GetCell(2).ToString() ?? "";
                string expectedName = string.IsNullOrWhiteSpace(products[0].ProductNameWithAttr)
                    ? products[0].ProductName
                    : products[0].ProductNameWithAttr;
                Assert(copiedName == expectedName, "Tên hàng bị sai sau khi copy.");
                Assert(Math.Abs(copiedPrice - products[0].Price) < 0.01, "Giá bán bị sai sau khi copy.");
                Assert(copiedUnit == products[0].Unit, "Đơn vị tính bị sai sau khi copy.");
            }

            ProductRow secondProduct = Clone(products[0]);
            secondProduct.ProductCode += "-2";
            secondProduct.ProductNameWithAttr += " dòng 2";
            secondProduct.Unit = string.IsNullOrWhiteSpace(products[0].Unit) ? "Cái" : "";
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
            using (var designer = new FormFlowDesigner())
            {
                designer.Show();
                Application.DoEvents();
                Assert(designer.Controls.Count > 0, "Trình thiết kế Flow không khởi tạo được.");
                designer.Close();
            }
            using (var main = new TestableFormMain())
            {
                main.Show();
                Application.DoEvents();
                List<ProductRow> loadedProducts = ReadField<List<ProductRow>>(main, "products");
                LabelDefinition? selectedLabel = ReadField<LabelDefinition?>(main, "selectedLabel");
                Button preview = ReadField<Button>(main, "btnPreview");
                Button print = ReadField<Button>(main, "btnPrint");
                DataGridView grid = ReadField<DataGridView>(main, "grid");
                TextBox filter = ReadField<TextBox>(main, "txtFilter");
                ComboBox specialFilter = ReadField<ComboBox>(main, "cmbSpecialFilter");
                Assert(loadedProducts.Count == 2, $"Màn chính không nạp đúng Excel: {loadedProducts.Count} dòng.");
                Assert(selectedLabel?.Code == "SMOKE", "Mẫu duy nhất không tự được chọn.");
                Assert(preview.Enabled, "Nút Xem trước phải được bật.");
                Assert(print.Enabled, "Nút In phải được bật.");

                specialFilter.SelectedIndex = 1;
                Application.DoEvents();
                Assert(grid.Rows.Count == 1, "Lọc thiếu đơn vị tính không đúng.");
                Assert(grid.Rows[0].Cells["Unit"].InheritedStyle.BackColor.R > 240,
                    "Ô thiếu đơn vị tính chưa được tô cảnh báo.");
                specialFilter.SelectedIndex = 2;
                Application.DoEvents();
                Assert(grid.Rows.Count == 1, "Lọc có đơn vị tính không đúng.");
                specialFilter.SelectedIndex = 0;
                filter.Text = "-2";
                Application.DoEvents();
                Assert(grid.Rows.Count == 1, "Tìm theo mã hàng không đúng.");
                filter.Clear();
                Application.DoEvents();

                Invoke(main, "SetFilteredPrintState", false);
                Assert(!preview.Enabled && !print.Enabled, "Bỏ chọn in chưa khóa Preview/In.");
                Invoke(main, "SetFilteredPrintState", true);
                Assert(preview.Enabled && print.Enabled, "Chọn in lại chưa bật Preview/In.");

                bool previewOpened = false;
                using System.Windows.Forms.Timer closePreview = new() { Interval = 250 };
                closePreview.Tick += (_, _) =>
                {
                    Form? previewForm = Application.OpenForms
                        .Cast<Form>()
                        .FirstOrDefault(form => form is FormPreview);
                    if (previewForm == null)
                        return;
                    previewOpened = true;
                    previewForm.Close();
                };
                closePreview.Start();
                preview.PerformClick();
                closePreview.Stop();
                Assert(previewOpened, "Bấm Xem trước không mở cửa sổ preview.");

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

            configService.Config.BarTenderExe = sourceTemplate;
            configService.Save();
            using (var mainWithoutBarTender = new FormMain())
            {
                Button preview = ReadField<Button>(mainWithoutBarTender, "btnPreview");
                Button print = ReadField<Button>(mainWithoutBarTender, "btnPrint");
                Assert(preview.Enabled, "Xem trước không được phụ thuộc bartend.exe.");
                Assert(!print.Enabled, "Nút In phải khóa khi chưa chọn đúng bartend.exe.");
            }
            configService.Config.BarTenderExe = fakeBarTender;
            configService.Save();

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
                Assert(saved.TemplatePath == templateBefore, "Lưu xong bị đổi đường dẫn .btw đã chọn.");
                Assert(File.Exists(ConfigService.Instance.GetStoredTemplatePath("SMOKE")),
                    "Bản dự phòng .btw không được tạo.");
                Assert(((LabelDefinition?)list.SelectedItem)?.Code == "SMOKE", "Lưu xong không giữ mẫu đang chọn.");
                Assert(FindControl<Button>(config, button => button.Text == "Mở thiết kế Flow") != null,
                    "Settings chưa hiển thị nút mở Flow.");
            }

            string missingTemplate = Path.Combine(runtime, "duong-dan-nguoi-dung", "mau-khong-ton-tai.btw");
            string missingData = Path.Combine(runtime, "duong-dan-nguoi-dung", "data-khong-ton-tai.xls");
            string missingBarTender = Path.Combine(runtime, "BarTender", "bartend.exe");
            configService.Config.BarTenderExe = missingBarTender;
            configService.Config.Labels[0].TemplatePath = missingTemplate;
            configService.Config.Labels[0].DataFilePath = missingData;
            configService.Save();
            configService.Load();
            Assert(configService.Config.BarTenderExe == missingBarTender,
                "Load config tự đổi đường dẫn BarTender người dùng đã chọn.");
            Assert(configService.Config.Labels[0].TemplatePath == missingTemplate,
                "Load config tự đổi đường dẫn .btw người dùng đã chọn.");
            Assert(configService.Config.Labels[0].DataFilePath == missingData,
                "Load config tự đổi đường dẫn data người dùng đã chọn.");

            Console.WriteLine("PASS: Excel import");
            Console.WriteLine("PASS: Two-file flow join, mapping and persistence");
            Console.WriteLine("PASS: Flow designer UI opens");
            Console.WriteLine("PASS: Fixed data file has exactly 3 columns");
            Console.WriteLine($"COPIED FILE: {dataFile}");
            Console.WriteLine($"COPIED VALUES: Tên hàng={copiedName} | Giá bán={copiedPrice:N0} | Đơn vị tính={copiedUnit}");
            Console.WriteLine("PASS: Single template auto-selection");
            Console.WriteLine("PASS: Preview/print readiness");
            Console.WriteLine("PASS: Preview button opens preview window");
            Console.WriteLine("PASS: Preview remains available without bartend.exe");
            Console.WriteLine("PASS: Search and special data filters");
            Console.WriteLine("PASS: Missing-unit warning color");
            Console.WriteLine("PASS: Include/exclude rows from print");
            Console.WriteLine("PASS: Enter/Shift+Enter cell navigation");
            Console.WriteLine("PASS: Four-direction arrow cell navigation");
            Console.WriteLine("PASS: Template name editing remains stable");
            Console.WriteLine("PASS: Settings preserve selected data/template");
            Console.WriteLine("PASS: Missing user paths never reset to managed defaults");
            Console.WriteLine("PASS: Flow designer is visible from Settings");
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

    private static T? FindControl<T>(Control root, Func<T, bool> predicate) where T : Control
    {
        foreach (Control child in root.Controls)
        {
            if (child is T typed && predicate(typed))
                return typed;
            T? nested = FindControl(child, predicate);
            if (nested != null)
                return nested;
        }
        return null;
    }

    private static void Invoke(object instance, string name, params object[] args)
    {
        Type? type = instance.GetType();
        MethodInfo? method = null;
        while (type != null && method == null)
        {
            method = type.GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic);
            type = type.BaseType;
        }
        if (method == null)
            throw new MissingMethodException(instance.GetType().Name, name);
        method.Invoke(instance, args);
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
    }

    private static ProductRow Clone(ProductRow item) => new()
    {
        IncludeForPrint = item.IncludeForPrint,
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
