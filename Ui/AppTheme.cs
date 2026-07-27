namespace KiotVietLabelPrinter.Ui;

public static class AppTheme
{
    public static readonly Color Canvas = Color.FromArgb(242, 245, 244);
    public static readonly Color Surface = Color.White;
    public static readonly Color SurfaceMuted = Color.FromArgb(232, 238, 236);
    public static readonly Color Ink = Color.FromArgb(24, 37, 34);
    public static readonly Color Muted = Color.FromArgb(92, 107, 103);
    public static readonly Color Accent = Color.FromArgb(14, 116, 84);
    public static readonly Color AccentDark = Color.FromArgb(9, 84, 61);
    public static readonly Color AccentSoft = Color.FromArgb(220, 240, 232);
    public static readonly Color Danger = Color.FromArgb(181, 66, 59);
    public static readonly Color Border = Color.FromArgb(205, 216, 212);

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
        button.BackColor = Ink;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Color.FromArgb(69, 88, 82);
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(39, 58, 52);
        button.FlatAppearance.MouseDownBackColor = Color.FromArgb(15, 27, 24);
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
