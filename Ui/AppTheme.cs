namespace KiotVietLabelPrinter.Ui;

public static class AppTheme
{
    public static readonly Color Canvas = Color.FromArgb(244, 247, 246);
    public static readonly Color Surface = Color.White;
    public static readonly Color SurfaceMuted = Color.FromArgb(235, 241, 239);
    public static readonly Color Ink = Color.FromArgb(28, 43, 39);
    public static readonly Color Muted = Color.FromArgb(99, 116, 111);
    public static readonly Color Accent = Color.FromArgb(18, 122, 91);
    public static readonly Color AccentDark = Color.FromArgb(13, 94, 70);
    public static readonly Color Danger = Color.FromArgb(181, 66, 59);
    public static readonly Color Border = Color.FromArgb(211, 221, 218);

    public static Font Display(float size, FontStyle style = FontStyle.Bold) =>
        new("Segoe UI Variable Display", size, style);

    public static Font Body(float size = 10F, FontStyle style = FontStyle.Regular) =>
        new("Segoe UI Variable Text", size, style);

    public static void StylePrimary(Button button)
    {
        button.BackColor = Accent;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.Font = Body(10F, FontStyle.Bold);
        button.Cursor = Cursors.Hand;
    }

    public static void StyleSecondary(Button button)
    {
        button.BackColor = Surface;
        button.ForeColor = Ink;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Border;
        button.FlatAppearance.BorderSize = 1;
        button.Font = Body(9.5F, FontStyle.Bold);
        button.Cursor = Cursors.Hand;
    }

    public static Label Caption(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Font = Body(9F, FontStyle.Bold),
        ForeColor = Ink
    };
}
