namespace KiotVietLabelPrinter.Ui;

public static class AppTheme
{
    public static readonly Color Canvas = Color.FromArgb(244, 246, 243);
    public static readonly Color Surface = Color.FromArgb(253, 254, 252);
    public static readonly Color SurfaceMuted = Color.FromArgb(237, 241, 237);
    public static readonly Color Ink = Color.FromArgb(20, 34, 29);
    public static readonly Color Muted = Color.FromArgb(92, 107, 103);
    public static readonly Color Accent = Color.FromArgb(18, 122, 87);
    public static readonly Color AccentDark = Color.FromArgb(9, 84, 61);
    public static readonly Color AccentSoft = Color.FromArgb(222, 242, 232);
    public static readonly Color Danger = Color.FromArgb(181, 66, 59);
    public static readonly Color Border = Color.FromArgb(215, 223, 218);

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
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(11, 101, 73);
        button.FlatAppearance.MouseDownBackColor = AccentDark;
        button.Font = Body(10F, FontStyle.Bold);
        button.Padding = new Padding(10, 0, 10, 1);
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
    }

    public static void StyleSecondary(Button button)
    {
        button.BackColor = Surface;
        button.ForeColor = Ink;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Border;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = SurfaceMuted;
        button.FlatAppearance.MouseDownBackColor = Color.FromArgb(220, 228, 225);
        button.Font = Body(9.5F, FontStyle.Bold);
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
    }

    public static void StyleHeaderButton(Button button)
    {
        button.BackColor = Surface;
        button.ForeColor = Ink;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Border;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = SurfaceMuted;
        button.FlatAppearance.MouseDownBackColor = AccentSoft;
        button.Font = Body(9.5F, FontStyle.Bold);
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
    }

    public static Label Caption(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Font = Body(9F, FontStyle.Bold),
        ForeColor = Ink
    };
}
