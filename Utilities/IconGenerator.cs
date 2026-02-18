namespace CloudAdmin365.Utilities;

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Windows.Forms;

/// <summary>
/// Icon generator for CloudAdmin365.
/// Produces a cog/gear icon with "365" text in the centre — an agnostic
/// representation of an Office 365 administration platform.
/// The design intentionally avoids service-specific imagery (envelopes,
/// mailboxes, etc.) so it remains appropriate as new PowerShell modules are
/// added to provide Exchange, Teams, SharePoint, Security, and other features.
/// </summary>
public static class IconGenerator
{
    // ── Colour palette ────────────────────────────────────────────────────
    private static readonly Color BgBlue      = Color.FromArgb(0, 120, 212);   // Microsoft blue
    private static readonly Color CogWhite    = Color.White;
    private static readonly Color TextColor   = Color.FromArgb(0, 120, 212);   // Blue text on white cog
    private static readonly Color CogShadow   = Color.FromArgb(40, 0, 0, 0);  // Subtle depth

    /// <summary>
    /// Returns the application icon. Since CloudAdmin365 does not embed an .ico
    /// resource (the icon is generated programmatically), this always delegates
    /// to <see cref="GenerateApplicationIcon"/>.
    /// </summary>
    public static Icon GetAppIcon() => GenerateApplicationIcon();

    /// <summary>
    /// Generates the CloudAdmin365 application icon: a white cog/gear with
    /// bold "365" text centred inside, on a Microsoft-blue rounded-square
    /// background. Returns a 32×32 icon suitable for window title bars.
    /// </summary>
    public static Icon GenerateApplicationIcon()
    {
        const int size = 32;
        var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);

        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode     = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            // ── 1. Background: rounded square in Microsoft blue ───────────
            using var bgBrush = new SolidBrush(BgBlue);
            FillRoundedRectangle(g, bgBrush, 0, 0, size, size, 7);

            // ── 2. Cog / gear (white, centred) ───────────────────────────
            const float cx = 16f, cy = 16f;
            const float outerR = 13.5f, innerR = 9.5f;
            const int   teeth  = 8;

            // Subtle shadow offset for depth
            using var shadowBrush = new SolidBrush(CogShadow);
            DrawGear(g, shadowBrush, cx + 0.5f, cy + 0.5f, outerR, innerR, teeth);

            using var cogBrush = new SolidBrush(CogWhite);
            DrawGear(g, cogBrush, cx, cy, outerR, innerR, teeth);

            // ── 3. "365" text centred inside the cog ──────────────────────
            using var textBrush = new SolidBrush(TextColor);
            using var textFont  = new Font("Segoe UI", 8.5f, FontStyle.Bold, GraphicsUnit.Pixel);

            var textSize = g.MeasureString("365", textFont);
            float tx = cx - textSize.Width  / 2f;
            float ty = cy - textSize.Height / 2f;
            g.DrawString("365", textFont, textBrush, tx, ty);
        }

        return Icon.FromHandle(bitmap.GetHicon());
    }

    /// <summary>
    /// Generates a small 16×16 tab icon: a cog with "365" text.
    /// </summary>
    public static Icon GenerateTabIcon()
    {
        const int size = 16;
        var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);

        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode     = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            const float cx = 8f, cy = 8f;
            const float outerR = 7f, innerR = 4.8f;
            const int   teeth  = 8;

            using var cogBrush = new SolidBrush(BgBlue);
            DrawGear(g, cogBrush, cx, cy, outerR, innerR, teeth);

            // "365" text inside
            using var textBrush = new SolidBrush(Color.White);
            using var textFont  = new Font("Segoe UI", 5f, FontStyle.Bold, GraphicsUnit.Pixel);

            var textSize = g.MeasureString("365", textFont);
            float tx = cx - textSize.Width  / 2f;
            float ty = cy - textSize.Height / 2f;
            g.DrawString("365", textFont, textBrush, tx, ty);
        }

        return Icon.FromHandle(bitmap.GetHicon());
    }

    // ──────────────────────────────────────────────────────────────────────
    // Helpers
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Draws a gear/cog wheel centred at (<paramref name="cx"/>, <paramref name="cy"/>).
    /// Uses a smooth star-polygon approach: alternate points sit on the outer and
    /// inner radii to form a continuous, round-tipped gear outline.
    /// </summary>
    private static void DrawGear(Graphics g, Brush brush, float cx, float cy,
                                 float outerR, float innerR, int teeth)
    {
        // Each tooth occupies two points (outer tip, inner valley).
        int points = teeth * 2;
        var pts = new PointF[points];
        double angleStep = Math.PI * 2.0 / points;

        for (int i = 0; i < points; i++)
        {
            // Start at -90° so the first tooth faces up.
            double angle = i * angleStep - Math.PI / 2.0;
            float r = (i % 2 == 0) ? outerR : innerR;
            pts[i] = new PointF(
                cx + (float)(r * Math.Cos(angle)),
                cy + (float)(r * Math.Sin(angle)));
        }

        using var path = new GraphicsPath();
        path.AddPolygon(pts);
        g.FillPath(brush, path);
    }

    /// <summary>
    /// Fills a rounded rectangle.
    /// </summary>
    private static void FillRoundedRectangle(Graphics g, Brush brush,
                                             float x, float y, float w, float h, float r)
    {
        using var path = new GraphicsPath();
        path.AddArc(x,            y,            r * 2, r * 2, 180, 90);
        path.AddArc(x + w - r * 2, y,           r * 2, r * 2, 270, 90);
        path.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0,   90);
        path.AddArc(x,            y + h - r * 2, r * 2, r * 2, 90,  90);
        path.CloseFigure();
        g.FillPath(brush, path);
    }
}
