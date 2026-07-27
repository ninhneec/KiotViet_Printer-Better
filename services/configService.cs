using Newtonsoft.Json;
using KiotVietLabelPrinter.Models;

namespace KiotVietLabelPrinter.Services;

public class ConfigService
{
    private static ConfigService? _instance;
    public static ConfigService Instance => _instance ??= new ConfigService();

    private readonly string _configPath;
    public string DataDirectory { get; }
    public string TemplatesDirectory => Path.Combine(DataDirectory, "Templates");

    public AppConfig Config { get; private set; } = new();

    private ConfigService()
    {
        DataDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "KiotVietPrinterBetter");
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

            // Phiên bản Better chỉ dùng luồng in trực tiếp ba trường.
            foreach (LabelDefinition label in Config.Labels)
            {
                label.HandlerType = "DIRECT_PRICE";
                label.DataFilePath = "";
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

    public string StoreTemplate(string sourcePath, string labelCode)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            return sourcePath;

        Directory.CreateDirectory(TemplatesDirectory);
        string safeCode = string.Concat((labelCode.Length == 0 ? "LABEL" : labelCode)
            .Select(character => Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        string targetPath = Path.Combine(TemplatesDirectory, $"{safeCode}.btw");

        if (!Path.GetFullPath(sourcePath).Equals(Path.GetFullPath(targetPath), StringComparison.OrdinalIgnoreCase))
            File.Copy(sourcePath, targetPath, true);

        return targetPath;
    }

    private AppConfig CreateDefaultConfig()
    {
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
                    HandlerType = "DIRECT_PRICE"
                }
            ]
        };
    }
}
