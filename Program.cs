using KiotVietLabelPrinter.Forms;
using KiotVietLabelPrinter.Services;

namespace KiotVietLabelPrinter;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        bool testUi = args.Contains("--test-ui", StringComparer.OrdinalIgnoreCase);
        bool isNewInstance = true;

        // Ten mutex phai khop voi AppMutex trong installer.iss de Inno Setup
        // co the phat hien va dong app dang chay truoc khi cai/ghi de.
        using var mutex = testUi
            ? null
            : new Mutex(true, "KiotVietLabelPrinterProV2Mutex", out isNewInstance);

        if (!testUi && !isNewInstance)
        {
            MessageBox.Show("Ung dung dang chay roi.", "In Tem KiotViet",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ApplicationConfiguration.Initialize();

        ConfigService.Instance.Load();

        int captureIndex = Array.FindIndex(
            args,
            value => value.Equals("--capture-ui", StringComparison.OrdinalIgnoreCase));
        if (captureIndex >= 0 && captureIndex + 1 < args.Length)
        {
            using FormMain preview = new();
            preview.Show();
            Application.DoEvents();
            using Bitmap bitmap = new(preview.ClientSize.Width, preview.ClientSize.Height);
            preview.DrawToBitmap(bitmap, new Rectangle(Point.Empty, preview.ClientSize));
            bitmap.Save(args[captureIndex + 1], System.Drawing.Imaging.ImageFormat.Png);
            return;
        }

        Application.Run(new FormMain());
    }
}
