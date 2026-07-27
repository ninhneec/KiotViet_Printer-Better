using System.Diagnostics;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace KiotVietLabelPrinter.Services;

public class BarTenderService
{
    private static readonly SemaphoreSlim PrintGate = new(1, 1);

    public void Print(string btwFile)
    {
        Print(btwFile, null, null);
    }

    public void Print(string btwFile, Dictionary<string, string>? namedSubStrings)
    {
        Print(btwFile, namedSubStrings, null);
    }

    public void Print(
        string btwFile,
        Dictionary<string, string>? namedSubStrings,
        int? recordCount)
    {
        PrintGate.Wait();
        string stage = "Khởi tạo";
        string? xmlPath = null;
        try
        {
            WritePrintLog($"START template={btwFile}");
            stage = "Tìm BarTender";
            string configuredPath = ConfigService.Instance.Config.BarTenderExe;
            string bartenderExe = ConfigService.Instance.ResolveBarTenderExecutable(configuredPath);
            if (!File.Exists(bartenderExe) ||
                !Path.GetFileName(bartenderExe).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase))
                throw new Exception(ConfigService.Instance.GetBarTenderDiagnostic());
            if (string.IsNullOrWhiteSpace(btwFile) || !File.Exists(btwFile))
                throw new Exception($"Không tìm thấy file tem:\n{btwFile}");

            string printerName = ConfigService.Instance.Config.PrinterName?.Trim() ?? "";
            if (!string.IsNullOrWhiteSpace(printerName) &&
                !PrinterSettings.InstalledPrinters.Cast<string>()
                    .Contains(printerName, StringComparer.OrdinalIgnoreCase))
                throw new Exception($"Không tìm thấy máy in đã chọn:\n{printerName}");

            stage = "Tạo XML";
            if (!string.IsNullOrWhiteSpace(printerName) && IsVirtualPrinter(printerName))
                throw new Exception(
                    $"“{printerName}” là máy in ảo nên sẽ yêu cầu lưu file.\n\n" +
                    "Mở Quản lý mẫu tem và chọn máy in tem thật.");

            xmlPath = CreatePrintXmlNearApp(
                btwFile,
                namedSubStrings,
                printerName,
                recordCount);
            WritePrintLog(
                $"XML={xmlPath} printer={(string.IsNullOrWhiteSpace(printerName) ? "(trong file .btw)" : printerName)}");

            bool databaseBatch = recordCount is > 0 &&
                                 (namedSubStrings == null || namedSubStrings.Count == 0);
            bool hasRunningBarTender = Process.GetProcessesByName("bartend").Length > 0;
            if (databaseBatch && hasRunningBarTender)
            {
                RecoverHiddenBarTenderProcesses();
                hasRunningBarTender = Process.GetProcessesByName("bartend").Length > 0;
            }
            if (databaseBatch && hasRunningBarTender)
                throw new Exception(
                    "BarTender đang mở hoặc đang đứng ở cửa sổ Print.\n\n" +
                    "Hãy đóng toàn bộ BarTender rồi bấm In lại. App không chạy " +
                    "hai phiên cùng lúc để tránh treo máy và in lặp.");
            // BarTender 10.1 không nhận /RecordRange ở command line (Error #3113).
            // Dùng RecordRange trong BTXML; /XMLScript bắt buộc là tham số cuối.
            string arguments = (hasRunningBarTender ? "" : "/X ") +
                               $"/XMLScript=\"{xmlPath}\"";
            if (databaseBatch)
            {
                stage = "Gửi lệnh qua BarTender Automation";
                WritePrintLog("METHOD=ACTIVEX_FORMAT10");
                string response = RunFormatViaActiveX(
                    btwFile,
                    printerName,
                    $"1-{recordCount!.Value}");
                WritePrintLog($"ACTIVEX_RESPONSE {LimitLog(response)}");
                if (ContainsBarTenderError(response))
                    throw new Exception(
                        "BarTender từ chối lệnh in.\n\n" + response);
                WritePrintLog($"SUCCESS template={btwFile}");
                return;
            }

            WritePrintLog("METHOD=BTXML10");
            WritePrintLog($"COMMAND exe={bartenderExe} args={arguments}");

            ProcessStartInfo psi = new()
            {
                FileName = bartenderExe,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            stage = "Mở BarTender";
            using Process? process = Process.Start(psi);
            if (process == null)
                throw new Exception("Không thể mở tiến trình BarTender.");

            // Đọc song song để tránh deadlock khi bộ đệm stdout/stderr đầy.
            Task<string> stdOutTask = process.StandardOutput.ReadToEndAsync();
            Task<string> stdErrTask = process.StandardError.ReadToEndAsync();
            stage = "Chờ BarTender xử lý";
            // In cả lô có thể cần lâu hơn trên máy yếu. Luồng này chạy trong
            // Task nền nên giao diện vẫn phản hồi trong lúc chờ BarTender.
            bool exited = process.WaitForExit(databaseBatch ? 60000 : 180000);
            if (!exited)
                throw new Exception(
                    "BarTender chưa phản hồi trong thời gian cho phép. " +
                    "Có thể đang chờ một hộp thoại xác nhận.");
            string stdOut = stdOutTask.GetAwaiter().GetResult();
            string stdErr = stdErrTask.GetAwaiter().GetResult();
            WritePrintLog($"RESPONSE stdout={LimitLog(stdOut)} stderr={LimitLog(stdErr)}");
            if (process.ExitCode != 0)
                throw new Exception(
                    $"BarTender trả ExitCode {process.ExitCode}.\n\nSTDERR:\n{stdErr}\n\nSTDOUT:\n{stdOut}");

            if (ContainsBarTenderError(stdOut) || ContainsBarTenderError(stdErr))
                throw new Exception(
                    "BarTender nhận lệnh nhưng từ chối tạo job in.\n\n" +
                    $"STDERR:\n{stdErr}\n\nSTDOUT:\n{stdOut}");

            if (databaseBatch)
            {
                // /X chỉ trả về sau khi BarTender xử lý xong lệnh. Không chờ thêm
                // 8 giây ở spooler vì khoảng chờ này làm app có cảm giác bị treo.
            }
            else if (!string.IsNullOrWhiteSpace(printerName))
            {
                stage = "Chờ hàng đợi máy in";
                WaitForPrintJobToFinish(printerName);
            }
            else
            {
                // Khi dùng máy in lưu trong .btw, app không biết tên queue.
                // Chỉ chờ BarTender đọc xong file data, không khóa app 10 phút.
                Thread.Sleep(1200);
            }
            WritePrintLog($"SUCCESS template={btwFile}");
        }
        catch (Exception ex)
        {
            WritePrintLog($"ERROR stage={stage} xml={xmlPath ?? "(chưa tạo)"} message={ex}");
            throw new Exception(
                $"Lỗi tại bước: {stage}\n\n{ex.Message}\n\nLog: {GetPrintLogPath()}",
                ex);
        }
        finally
        {
            PrintGate.Release();
        }
    }

    private static string GetPrintLogPath()
    {
        string folder = Path.Combine(ConfigService.Instance.DataDirectory, "Logs");
        Directory.CreateDirectory(folder);
        return Path.Combine(folder, "printing.log");
    }

    private static void WritePrintLog(string message)
    {
        try
        {
            File.AppendAllText(
                GetPrintLogPath(),
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}",
                new UTF8Encoding(false));
        }
        catch
        {
        }
    }

    private static bool ContainsBarTenderError(string value)
    {
        return value.Contains("Severity=\"Error\"", StringComparison.OrdinalIgnoreCase) ||
               value.Contains("<Error", StringComparison.OrdinalIgnoreCase) ||
               value.Contains("ErrorCode", StringComparison.OrdinalIgnoreCase);
    }

    private static void RecoverHiddenBarTenderProcesses()
    {
        Process[] processes = Process.GetProcessesByName("bartend");
        bool hasVisibleWindow = processes.Any(process =>
        {
            try
            {
                return !process.HasExited && process.MainWindowHandle != IntPtr.Zero;
            }
            catch
            {
                return false;
            }
        });
        if (hasVisibleWindow)
        {
            foreach (Process process in processes)
                process.Dispose();
            return;
        }

        foreach (Process process in processes)
        {
            try
            {
                WritePrintLog($"RECOVER hidden_process pid={process.Id}");
                process.Kill(entireProcessTree: true);
                process.WaitForExit(5000);
            }
            catch (Exception ex)
            {
                WritePrintLog($"RECOVER_FAILED pid={process.Id} message={ex.Message}");
            }
            finally
            {
                process.Dispose();
            }
        }

        Thread.Sleep(500);
    }

    private static string RunFormatViaActiveX(
        string btwFile,
        string printerName,
        string recordRange)
    {
        string runnerPath = Path.Combine(
            ConfigService.Instance.DataDirectory,
            "bartender_automation.vbs");
        File.WriteAllText(
            runnerPath,
            """
            Option Explicit
            Dim btApp, btFormat, btPrintSetup
            Dim btwPath, printerName, recordRange
            btwPath = WScript.Arguments(0)
            printerName = WScript.Arguments(1)
            recordRange = WScript.Arguments(2)

            On Error Resume Next
            Set btApp = CreateObject("BarTender.Application")
            If Err.Number <> 0 Then
              WScript.StdErr.WriteLine "CREATE_COM: " & Err.Description
              WScript.Quit 2
            End If

            Err.Clear
            btApp.Visible = False
            Set btFormat = btApp.Formats.Open(btwPath, False, "")
            If Err.Number <> 0 Then
              WScript.StdErr.WriteLine "OPEN_FORMAT: " & Err.Description
              btApp.Quit 1
              WScript.Quit 3
            End If

            Err.Clear
            btFormat.RecordRange = recordRange
            If Err.Number <> 0 Then
              WScript.StdErr.WriteLine "RECORD_RANGE: " & Err.Description
              btFormat.Close 1
              btApp.Quit 1
              WScript.Quit 4
            End If

            If Len(printerName) > 0 Then
              Err.Clear
              Set btPrintSetup = btFormat.PrintSetup
              btPrintSetup.Printer = printerName
              If Err.Number <> 0 Then
                WScript.StdErr.WriteLine "PRINTER: " & Err.Description
                btFormat.Close 1
                btApp.Quit 1
                WScript.Quit 5
              End If
            End If

            Err.Clear
            btFormat.PrintOut False, True
            If Err.Number <> 0 Then
              WScript.StdErr.WriteLine "PRINT_OUT: " & Err.Description
              btFormat.Close 1
              btApp.Quit 1
              WScript.Quit 6
            End If

            btFormat.Close 1
            btApp.Quit 1
            WScript.Echo "OK range=" & recordRange
            """,
            new UTF8Encoding(false));

        ProcessStartInfo startInfo = new()
        {
            FileName = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                "cscript.exe"),
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("//NoLogo");
        startInfo.ArgumentList.Add(runnerPath);
        startInfo.ArgumentList.Add(btwFile);
        startInfo.ArgumentList.Add(printerName);
        startInfo.ArgumentList.Add(recordRange);

        using Process? process = Process.Start(startInfo);
        if (process == null)
            throw new Exception("Không thể mở Windows Script Host.");

        Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(180000))
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch
            {
            }
            throw new Exception(
                "BarTender Automation không phản hồi sau 60 giây.");
        }

        string response = outputTask.GetAwaiter().GetResult();
        string error = errorTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
            throw new Exception(
                $"BarTender Automation trả mã {process.ExitCode}.\n\n{error}");

        return response;
    }

    private static string LimitLog(string value)
    {
        string singleLine = (value ?? string.Empty)
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Trim();
        return singleLine.Length <= 4000
            ? singleLine
            : singleLine[..4000] + "...";
    }

    //---------------------------------------------------------
    // Chờ máy in xử lý xong job (dựa vào hàng đợi Windows Spooler)
    //---------------------------------------------------------

    private static void WaitForPrintJobToFinish(
        string printerName,
        int startupGraceMs = 8000,
        int maxWaitMs = 120000,
        int pollIntervalMs = 200)
    {
        if (!OpenPrinter(printerName, out IntPtr hPrinter, IntPtr.Zero))
            return; // Không lấy được handle máy in — bỏ qua, không chặn luồng in

        try
        {
            Stopwatch sw = Stopwatch.StartNew();
            bool sawJob = false;

            // 1) Chờ job thực sự xuất hiện trong hàng đợi (BarTender vừa
            // nhận lệnh và bắt đầu tạo job). Nếu không thấy trong khoảng
            // grace này thì có thể job đã in xong quá nhanh (số lượng nhỏ)
            // — bỏ qua bước chờ rỗng để không trễ vô ích.
            while (sw.ElapsedMilliseconds < startupGraceMs)
            {
                if (GetQueuedJobCount(hPrinter) > 0)
                {
                    sawJob = true;
                    break;
                }

                Thread.Sleep(pollIntervalMs);
            }

            if (!sawJob)
                return;

            // 2) Chờ hàng đợi rỗng trở lại = máy in đã nhận xong toàn bộ
            // số lượng nhãn của job này.
            while (sw.ElapsedMilliseconds < maxWaitMs)
            {
                if (GetQueuedJobCount(hPrinter) == 0)
                    return;

                Thread.Sleep(pollIntervalMs);
            }
        }
        finally
        {
            ClosePrinter(hPrinter);
        }
    }

    private static int GetQueuedJobCount(IntPtr hPrinter)
    {
        GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out uint neededBytes);

        if (neededBytes == 0)
            return 0;

        IntPtr buffer = Marshal.AllocHGlobal((int)neededBytes);

        try
        {
            if (!GetPrinter(hPrinter, 2, buffer, neededBytes, out _))
                return 0;

            PRINTER_INFO_2 info = Marshal.PtrToStructure<PRINTER_INFO_2>(buffer);

            return (int)info.cJobs;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    private struct PRINTER_INFO_2
    {
        public string? pServerName;
        public string? pPrinterName;
        public string? pShareName;
        public string? pPortName;
        public string? pDriverName;
        public string? pComment;
        public string? pLocation;
        public IntPtr pDevMode;
        public string? pSepFile;
        public string? pPrintProcessor;
        public string? pDatatype;
        public string? pParameters;
        public IntPtr pSecurityDescriptor;
        public uint Attributes;
        public uint Priority;
        public uint DefaultPriority;
        public uint StartTime;
        public uint UntilTime;
        public uint Status;
        public uint cJobs;
        public uint AveragePPM;
    }

    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern bool GetPrinter(IntPtr hPrinter, uint dwLevel, IntPtr pPrinter, uint cbBuf, out uint pcbNeeded);

    private static string CreatePrintXmlNearApp(
        string btwFile,
            Dictionary<string, string>? namedSubStrings,
            string printerName,
            int? recordCount)
    {
        string debugFolder = Path.Combine(ConfigService.Instance.DataDirectory, "debug_xml");
        Directory.CreateDirectory(debugFolder);

        string xmlPath = Path.Combine(
            debugFolder,
            $"print_{DateTime.Now:yyyyMMdd_HHmmss_fff}.xml");

        StringBuilder sb = new();

        sb.AppendLine("""<?xml version="1.0" encoding="utf-8"?>""");
        sb.AppendLine("""<XMLScript Version="2.0">""");
        sb.AppendLine("""  <Command>""");
        // Chờ BarTender đọc và gửi xong đúng dòng hiện tại trước khi app
        // được phép ghi dữ liệu của sản phẩm tiếp theo.
        // Không yêu cầu BarTender trả dữ liệu/summary của hàng trăm tem:
        // response nhỏ hơn đáng kể trên máy cấu hình thấp.
        sb.AppendLine("""    <Print WaitForJobToComplete="true" Timeout="180000" ReturnPrintData="false" ReturnSummary="false" ReturnLabelData="false">""");
        // BarTender giữ kết nối Excel chừng nào tài liệu .btw còn mở.
        // Đóng tài liệu sau mỗi job để nhả file data cho sản phẩm kế tiếp.
        sb.AppendLine($"      <Format CloseAtEndOfJob=\"true\">{EscapeXml(btwFile)}</Format>");
        sb.AppendLine("""      <PrintSetup>""");
        if (!string.IsNullOrWhiteSpace(printerName))
            sb.AppendLine($"        <Printer>{EscapeXml(printerName)}</Printer>");
        // Ghi đè setting "First Record Only" có thể đã được lưu trong .btw.
        // Range lớn chỉ in đến record thực tế đang có trong file data.
        if (recordCount is > 0)
            sb.AppendLine($"        <RecordRange>1-{recordCount.Value}</RecordRange>");
        sb.AppendLine("""        <UseDatabase>true</UseDatabase>""");
        sb.AppendLine("""        <EnablePrompting>false</EnablePrompting>""");
        sb.AppendLine("""      </PrintSetup>""");

        if (namedSubStrings != null && namedSubStrings.Count > 0)
        {
            foreach (var kv in namedSubStrings)
            {
                if (string.IsNullOrWhiteSpace(kv.Key))
                    continue;

                string value = kv.Value ?? string.Empty;

                sb.AppendLine($"      <NamedSubString Name=\"{EscapeXml(kv.Key)}\">");
                sb.AppendLine($"        <Value>{EscapeXml(value)}</Value>");
                sb.AppendLine("      </NamedSubString>");
            }
        }

        sb.AppendLine("""    </Print>""");
        sb.AppendLine("""  </Command>""");
        sb.AppendLine("""</XMLScript>""");

        string xml = sb.ToString();

        using (FileStream fs = new(xmlPath, FileMode.Create, FileAccess.Write, FileShare.Read))
        using (StreamWriter writer = new(fs, new UTF8Encoding(false)))
        {
            writer.Write(xml);
            writer.Flush();
            fs.Flush(true);
        }

        File.WriteAllText(Path.Combine(debugFolder, "last_print_debug.xml"), xml, new UTF8Encoding(false));

        return xmlPath;
    }

    private static bool IsVirtualPrinter(string printerName)
    {
        string value = printerName.Trim().ToLowerInvariant();
        return value.Contains("pdf") ||
               value.Contains("xps") ||
               value.Contains("onenote") ||
               value.Contains("fax") ||
               value.Contains("document writer");
    }

    private static string GetDefaultPrinterName()
    {
        PrinterSettings settings = new();

        if (string.IsNullOrWhiteSpace(settings.PrinterName))
            throw new Exception("Không xác định được máy in mặc định của Windows.");

        if (!settings.IsValid)
            throw new Exception($"Máy in mặc định không hợp lệ: {settings.PrinterName}");

        return settings.PrinterName;
    }

    private static string EscapeXml(string value)
    {
        return System.Security.SecurityElement.Escape(value) ?? value;
    }

    private static string SafeReadAllText(string path)
    {
        try
        {
            if (!File.Exists(path))
                return "(Không tìm thấy file XML)";

            return File.ReadAllText(path);
        }
        catch (Exception ex)
        {
            return $"(Không đọc được file XML: {ex.Message})";
        }
    }
}
