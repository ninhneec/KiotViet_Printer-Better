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
        string path = (configuredPath ?? "").Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(path))
            return "";
        if (File.Exists(path))
        {
            if (Path.GetExtension(path).Equals(".lnk", StringComparison.OrdinalIgnoreCase))
                return ResolveWindowsShortcut(path);
            return path;
        }
        if (!Directory.Exists(path))
            return path;

        string direct = Path.Combine(path, "bartend.exe");
        if (File.Exists(direct))
            return direct;
        try
        {
            foreach (string shortcut in Directory.EnumerateFiles(path, "*.lnk", SearchOption.TopDirectoryOnly))
            {
                string target = ResolveWindowsShortcut(shortcut);
                if (File.Exists(target) &&
                    Path.GetFileName(target).Equals("bartend.exe", StringComparison.OrdinalIgnoreCase))
                    return target;
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
        return path;
    }

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
        string[] roots =
        [
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        ];
        List<string> results = new();
        foreach (string root in roots.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            string[] likelyDirectories =
            [
                Path.Combine(root, "Seagull"),
                Path.Combine(root, "Seagull Scientific"),
                Path.Combine(root, "BarTender")
            ];
            foreach (string directory in likelyDirectories.Where(Directory.Exists))
            {
                try
                {
                    results.AddRange(Directory.EnumerateFiles(
                        directory, "bartend.exe", SearchOption.AllDirectories));
                }
                catch (UnauthorizedAccessException)
                {
                }
                catch (IOException)
                {
                }
            }
        }
        return results.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
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
