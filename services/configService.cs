using Newtonsoft.Json;
using KiotVietLabelPrinter.Models;

namespace KiotVietLabelPrinter.Services;

public class ConfigService
{
    private static ConfigService? _instance;
    public static ConfigService Instance => _instance ??= new ConfigService();

    private readonly string _configPath;
    public string ConfigFilePath => _configPath;
    public string DataDirectory { get; }
    public string TemplatesDirectory => Path.Combine(DataDirectory, "Templates");
    public string DataFilesDirectory => Path.Combine(DataDirectory, "Data");

    public AppConfig Config { get; private set; } = new();

    private ConfigService()
    {
        string? overrideDirectory = Environment.GetEnvironmentVariable("KIOTVIET_TEST_DATA_DIRECTORY");
        DataDirectory = string.IsNullOrWhiteSpace(overrideDirectory)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KiotVietPrinterBetter")
            : Path.GetFullPath(overrideDirectory);
        _configPath = Path.Combine(DataDirectory, "config.json");

        Load();
    }

    public void Load()
    {
        try
        {
            string? folder = Path.GetDirectoryName(_configPath);

            if (!string.IsNullOrWhiteSpace(folder) && !Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            if (!File.Exists(_configPath))
            {
                string legacyPath = Path.Combine(Application.StartupPath, "Config", "config.json");
                Config = File.Exists(legacyPath)
                    ? JsonConvert.DeserializeObject<AppConfig>(File.ReadAllText(legacyPath)) ?? CreateDefaultConfig()
                    : CreateDefaultConfig();
                Save();
                return;
            }

            string json = File.ReadAllText(_configPath);

            if (string.IsNullOrWhiteSpace(json))
            {
                Config = CreateDefaultConfig();
                Save();
                return;
            }

            Config = JsonConvert.DeserializeObject<AppConfig>(json)
                     ?? CreateDefaultConfig();

            // Phòng trường hợp config cũ hoặc labels null
            if (Config.Labels == null)
                Config.Labels = new List<LabelDefinition>();

            if (Config.Labels.Count == 0)
            {
                Config = CreateDefaultConfig();
            }

            // Mỗi mẫu dùng một file data trung gian có đường dẫn cố định.
            foreach (LabelDefinition label in Config.Labels)
            {
                label.HandlerType = "DIRECT_PRICE";
                if (string.IsNullOrWhiteSpace(label.DataFilePath))
                    label.DataFilePath = GetManagedDataFilePath(label.Code);
                label.RequiresEmployeeCode = false;
                label.UseBarcodeParser = false;
                label.AppendEmployeeCode = false;
            }

            Save();
        }
        catch
        {
            Config = CreateDefaultConfig();
            Save();
        }
    }

    public void Save()
    {
        string? folder = Path.GetDirectoryName(_configPath);

        if (!string.IsNullOrWhiteSpace(folder) && !Directory.Exists(folder))
            Directory.CreateDirectory(folder);

        string json = JsonConvert.SerializeObject(
            Config,
            Formatting.Indented);

        File.WriteAllText(_configPath, json);
    }

    public bool IsConfigured()
    {
        if (string.IsNullOrWhiteSpace(Config.BarTenderExe))
            return false;

        if (Config.Labels == null || Config.Labels.Count == 0)
            return false;

        return Config.Labels.Any(x =>
            !string.IsNullOrWhiteSpace(x.TemplatePath) &&
            File.Exists(x.TemplatePath));
    }

    public bool IsBarTenderExecutableValid()
    {
        string path = ResolveBarTenderExecutable(Config.BarTenderExe);
        return File.Exists(path) &&
               Path.GetFileName(path).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase);
    }

    public string ResolveBarTenderExecutable(string configuredPath)
    {
        string path = NormalizeConfiguredPath(configuredPath);
        if (string.IsNullOrWhiteSpace(path))
            return "";
        if (File.Exists(path))
        {
            if (Path.GetExtension(path).Equals(".lnk", StringComparison.OrdinalIgnoreCase))
            {
                string shortcutTarget = ResolveWindowsShortcut(path);
                if (IsBarTenderExecutable(shortcutTarget))
                    return shortcutTarget;
                return FindBarTenderExecutables().FirstOrDefault() ?? shortcutTarget;
            }
            return path;
        }
        if (Directory.Exists(path))
        {
            string? insideFolder = FindExecutableUnderDirectory(path).FirstOrDefault();
            if (insideFolder != null)
                return insideFolder;
            foreach (string shortcut in SafeEnumerateFiles(path, "*.lnk", SearchOption.AllDirectories))
            {
                string target = ResolveWindowsShortcut(shortcut);
                if (IsBarTenderExecutable(target))
                    return target;
            }
        }

        if (Path.GetFileName(path).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase) ||
            path.Contains("bartender", StringComparison.OrdinalIgnoreCase) ||
            path.Contains("seagull", StringComparison.OrdinalIgnoreCase))
            return FindBarTenderExecutables().FirstOrDefault() ?? path;
        return path;
    }

    public string GetBarTenderDiagnostic()
    {
        string selected = Config.BarTenderExe;
        string resolved = ResolveBarTenderExecutable(selected);
        IReadOnlyList<string> candidates = FindBarTenderExecutables();
        return $"Đã chọn: {selected}\n" +
               $"Thực thi: {resolved}\n" +
               $"Tồn tại: {(File.Exists(resolved) ? "Có" : "Không")}\n" +
               $"Ứng viên tự tìm: {(candidates.Count == 0 ? "(không có)" : string.Join(" | ", candidates))}";
    }

    private static string NormalizeConfiguredPath(string value)
    {
        string path = Environment.ExpandEnvironmentVariables(value ?? "")
            .Replace("\r", "")
            .Replace("\n", "")
            .Trim()
            .Trim('"');
        if (path.StartsWith("file:///", StringComparison.OrdinalIgnoreCase) &&
            Uri.TryCreate(path, UriKind.Absolute, out Uri? uri) && uri.IsFile)
            path = uri.LocalPath;
        return path;
    }

    private static bool IsBarTenderExecutable(string path) =>
        File.Exists(path) &&
        Path.GetFileName(path).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase);

    private static string ResolveWindowsShortcut(string shortcutPath)
    {
        string installerTarget = ResolveAdvertisedInstallerShortcut(shortcutPath);
        if (File.Exists(installerTarget) &&
            Path.GetFileName(installerTarget).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase))
            return installerTarget;

        object? shell = null;
        object? shortcut = null;
        try
        {
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null) return shortcutPath;
            shell = Activator.CreateInstance(shellType);
            shortcut = shellType.InvokeMember(
                "CreateShortcut",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                shell,
                [shortcutPath]);
            return shortcut?.GetType().InvokeMember(
                "TargetPath",
                System.Reflection.BindingFlags.GetProperty,
                null,
                shortcut,
                null)?.ToString() ?? shortcutPath;
        }
        catch
        {
            return shortcutPath;
        }
        finally
        {
            if (shortcut != null && System.Runtime.InteropServices.Marshal.IsComObject(shortcut))
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shortcut);
            if (shell != null && System.Runtime.InteropServices.Marshal.IsComObject(shell))
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shell);
        }
    }

    private static string ResolveAdvertisedInstallerShortcut(string shortcutPath)
    {
        try
        {
            System.Text.StringBuilder productCode = new(39);
            System.Text.StringBuilder featureId = new(256);
            System.Text.StringBuilder componentCode = new(39);
            uint result = MsiGetShortcutTarget(
                shortcutPath,
                productCode,
                featureId,
                componentCode);
            if (result != 0 || productCode.Length == 0 || componentCode.Length == 0)
                return "";

            uint pathLength = 2048;
            System.Text.StringBuilder componentPath = new((int)pathLength);
            int state = MsiGetComponentPath(
                productCode.ToString(),
                componentCode.ToString(),
                componentPath,
                ref pathLength);
            return state is 3 or 4 && componentPath.Length > 0
                ? componentPath.ToString()
                : "";
        }
        catch
        {
            return "";
        }
    }

    [System.Runtime.InteropServices.DllImport(
        "msi.dll",
        CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern uint MsiGetShortcutTarget(
        string shortcutTarget,
        System.Text.StringBuilder productCode,
        System.Text.StringBuilder featureId,
        System.Text.StringBuilder componentCode);

    [System.Runtime.InteropServices.DllImport(
        "msi.dll",
        CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int MsiGetComponentPath(
        string productCode,
        string componentCode,
        System.Text.StringBuilder pathBuffer,
        ref uint pathBufferLength);

    public IReadOnlyList<string> FindBarTenderExecutables()
    {
        List<string> results = new();
        AddRunningBarTender(results);
        AddRegistryBarTender(results);
        AddPathBarTender(results);

        string[] roots =
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        };
        foreach (string root in roots.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            foreach (string vendor in new[] { "Seagull", "Seagull Scientific", "BarTender" })
                results.AddRange(FindExecutableUnderDirectory(Path.Combine(root, vendor)));
        }

        string[] startMenus =
        {
            Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu),
            Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
        };
        foreach (string startMenu in startMenus.Where(Directory.Exists))
        {
            foreach (string shortcut in SafeEnumerateFiles(startMenu, "*.lnk", SearchOption.AllDirectories)
                         .Where(file => file.Contains("bartender", StringComparison.OrdinalIgnoreCase)))
            {
                string target = ResolveWindowsShortcut(shortcut);
                if (IsBarTenderExecutable(target))
                    results.Add(target);
            }
        }
        return results
            .Select(NormalizeConfiguredPath)
            .Where(IsBarTenderExecutable)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(path => GetFileVersion(path))
            .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IEnumerable<string> FindExecutableUnderDirectory(string directory)
    {
        if (!Directory.Exists(directory)) return [];
        return SafeEnumerateFiles(directory, "bartend.exe", SearchOption.AllDirectories);
    }

    private static IEnumerable<string> SafeEnumerateFiles(
        string directory,
        string pattern,
        SearchOption searchOption)
    {
        try
        {
            return Directory.EnumerateFiles(directory, pattern, searchOption).ToList();
        }
        catch (UnauthorizedAccessException)
        {
            return [];
        }
        catch (IOException)
        {
            return [];
        }
    }

    private static void AddRunningBarTender(List<string> results)
    {
        try
        {
            foreach (System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("bartend"))
            {
                try
                {
                    string? path = process.MainModule?.FileName;
                    if (!string.IsNullOrWhiteSpace(path)) results.Add(path);
                }
                catch
                {
                }
                finally
                {
                    process.Dispose();
                }
            }
        }
        catch
        {
        }
    }

    private static void AddPathBarTender(List<string> results)
    {
        string? pathValue = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(pathValue)) return;
        foreach (string folder in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                string candidate = Path.Combine(folder.Trim().Trim('"'), "bartend.exe");
                if (File.Exists(candidate)) results.Add(candidate);
            }
            catch
            {
            }
        }
    }

    private static void AddRegistryBarTender(List<string> results)
    {
        foreach (Microsoft.Win32.RegistryHive hive in new[]
                 {
                     Microsoft.Win32.RegistryHive.LocalMachine,
                     Microsoft.Win32.RegistryHive.CurrentUser
                 })
        foreach (Microsoft.Win32.RegistryView view in new[]
                 {
                     Microsoft.Win32.RegistryView.Registry64,
                     Microsoft.Win32.RegistryView.Registry32
                 })
        {
            try
            {
                using Microsoft.Win32.RegistryKey baseKey =
                    Microsoft.Win32.RegistryKey.OpenBaseKey(hive, view);
                using (Microsoft.Win32.RegistryKey? appPath =
                       baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\bartend.exe"))
                {
                    AddRegistryCandidate(results, appPath?.GetValue(null)?.ToString());
                }
                using Microsoft.Win32.RegistryKey? uninstall =
                    baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                if (uninstall == null) continue;
                foreach (string keyName in uninstall.GetSubKeyNames())
                {
                    using Microsoft.Win32.RegistryKey? entry = uninstall.OpenSubKey(keyName);
                    string displayName = entry?.GetValue("DisplayName")?.ToString() ?? "";
                    string publisher = entry?.GetValue("Publisher")?.ToString() ?? "";
                    if (!displayName.Contains("bartender", StringComparison.OrdinalIgnoreCase) &&
                        !publisher.Contains("seagull", StringComparison.OrdinalIgnoreCase))
                        continue;
                    AddRegistryCandidate(results, entry?.GetValue("DisplayIcon")?.ToString());
                    string installLocation = entry?.GetValue("InstallLocation")?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(installLocation))
                        results.AddRange(FindExecutableUnderDirectory(installLocation));
                }
            }
            catch
            {
            }
        }
    }

    private static void AddRegistryCandidate(List<string> results, string? rawValue)
    {
        if (string.IsNullOrWhiteSpace(rawValue)) return;
        string value = rawValue.Trim().Trim('"');
        int iconIndex = value.LastIndexOf(',');
        if (iconIndex > 2 && int.TryParse(value[(iconIndex + 1)..], out _))
            value = value[..iconIndex].Trim().Trim('"');
        if (Directory.Exists(value))
            results.AddRange(FindExecutableUnderDirectory(value));
        else if (File.Exists(value))
            results.Add(value);
    }

    private static Version GetFileVersion(string path)
    {
        try
        {
            return Version.TryParse(
                System.Diagnostics.FileVersionInfo.GetVersionInfo(path).FileVersion,
                out Version? version)
                ? version
                : new Version();
        }
        catch
        {
            return new Version();
        }
    }

    public string StoreTemplate(string sourcePath, string labelCode)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            return sourcePath;

        Directory.CreateDirectory(TemplatesDirectory);
        string targetPath = GetStoredTemplatePath(labelCode);

        if (!Path.GetFullPath(sourcePath).Equals(Path.GetFullPath(targetPath), StringComparison.OrdinalIgnoreCase))
            File.Copy(sourcePath, targetPath, true);

        return targetPath;
    }

    public string GetStoredTemplatePath(string labelCode)
    {
        string safeCode = string.Concat((string.IsNullOrWhiteSpace(labelCode) ? "LABEL" : labelCode)
            .Select(character => Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        return Path.Combine(TemplatesDirectory, $"{safeCode}.btw");
    }

    public string GetManagedDataFilePath(string labelCode)
    {
        Directory.CreateDirectory(DataFilesDirectory);
        string safeCode = string.Concat((string.IsNullOrWhiteSpace(labelCode) ? "PRICE_LABEL" : labelCode)
            .Select(character => Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        return Path.Combine(DataFilesDirectory, $"{safeCode}_data.xls");
    }

    private AppConfig CreateDefaultConfig()
    {
        string dataFilePath = GetManagedDataFilePath("PRICE_LABEL");
        new ExcelService().EnsureDirectPriceDataFile(dataFilePath);
        return new AppConfig
        {
            BarTenderExe = "",
            LastFolder = "",
            LastExcelFile = "",
            AutoOpenLastFolder = true,
            RememberEmployee = true,
            DefaultEmployee = "",
            Labels =
            [
                new LabelDefinition
                {
                    Code = "PRICE_LABEL",
                    Name = "Tem giá sản phẩm",
                    Description = "Tên hàng, giá bán và đơn vị tính",
                    IconText = "🏷",
                    IsEnabled = true,
                    HandlerType = "DIRECT_PRICE",
                    DataFilePath = dataFilePath
                }
            ]
        };
    }
}
