namespace CloudAdmin365.Utilities;

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

/// <summary>
/// Icon generator for CloudAdmin365.
/// Produces a cloud + cog icon representing cloud administration.
/// </summary>
public static class IconGenerator
{
    private static readonly Color IconBlue  = Color.FromArgb(0, 120, 212);   // Microsoft blue
    private static readonly Color CloudWhite = Color.White;
    private static readonly Color CogColor  = Color.FromArgb(200, 230, 255); // Pale blue cog

    /// <summary>
    /// Returns the icon embedded in the application executable.
    /// Falls back to generating one programmatically if not found.
    /// </summary>
    public static Icon GetAppIcon()
    {
        try
        {
            var exePath = Application.ExecutablePath;
            if (!string.IsNullOrWhiteSpace(exePath) && File.Exists(exePath))
            {
                var icon = Icon.ExtractAssociatedIcon(exePath);
                if (icon != null)
                    return icon;
            }
        }
        catch { }

        return GenerateApplicationIcon();
    }

    /// <summary>
    /// Generates the CloudAdmin365 application icon: a cloud with a gear overlay.
    /// Returns a 32x32 icon suitable for window title bars.
    /// </summary>
    public static Icon GenerateApplicationIcon()
    {
        const int size = 32;
        var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);

        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            // ── Background: rounded square in Microsoft blue ──────────────
            using var bgBrush = new SolidBrush(IconBlue);
            FillRoundedRectangle(g, bgBrush, 0, 0, size, size, 7);

            // ── Cloud silhouette (white, upper area) ──────────────────────
            using var cloudBrush = new SolidBrush(CloudWhite);
            using var cloudPath  = BuildCloudPath(4, 5, 24, 16);
            g.FillPath(cloudBrush, cloudPath);

            // ── Gear (pale blue, overlapping cloud bottom) ────────────────
            using var cogBrush = new SolidBrush(CogColor);
            DrawGear(g, cogBrush, cx: 16, cy: 21, outerR: 7f, innerR: 3.5f, teeth: 6);
        }

        return Icon.FromHandle(bitmap.GetHicon());
    }

    /// <summary>
    /// Generates a small 16x16 tab icon: cloud + gear.
    /// </summary>
    public static Icon GenerateTabIcon()
    {
        const int size = 16;
        var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb);

        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            using var cloudBrush = new SolidBrush(IconBlue);
            using var cloudPath  = BuildCloudPath(1, 2, 14, 8);
            g.FillPath(cloudBrush, cloudPath);

            using var cogBrush = new SolidBrush(Color.FromArgb(0, 80, 160));
            DrawGear(g, cogBrush, cx: 8, cy: 11, outerR: 4f, innerR: 1.8f, teeth: 6);
        }

        return Icon.FromHandle(bitmap.GetHicon());
    }

    // ──────────────────────────────────────────────────────────────────────
    // Helpers
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds a cloud GraphicsPath using three overlapping ellipses merged with a base rectangle.
    /// </summary>
    private static GraphicsPath BuildCloudPath(float x, float y, float w, float h)
    {
        var path = new GraphicsPath();

        float baseTop    = y + h * 0.52f;
        float baseHeight = h - (baseTop - y);

        // Three bumps (left, centre, right)
        path.AddEllipse(x,              baseTop - h * 0.40f, w * 0.40f, h * 0.50f);
        path.AddEllipse(x + w * 0.28f,  y,                   w * 0.44f, h * 0.60f);
        path.AddEllipse(x + w * 0.58f,  baseTop - h * 0.30f, w * 0.40f, h * 0.46f);

        // Filled base — rounds off the bottom of the cloud
        path.AddRectangle(new RectangleF(x, baseTop, w, baseHeight));

        return path;
    }

    /// <summary>
    /// Draws a gear/cog wheel centred at (cx, cy).
    /// </summary>
    private static void DrawGear(Graphics g, Brush brush, float cx, float cy, float outerR, float innerR, int teeth)
    {
        var path = new GraphicsPath();

        double step    = Math.PI * 2.0 / teeth;
        float  toothW  = (float)(step * 0.40);
        float  toothH  = outerR - innerR + 1.5f;

        // Teeth: one small rectangle per tooth, rotated around the centre
        for (int i = 0; i < teeth; i++)
        {
            double angle = i * step - Math.PI / 2.0;

            using var toothPath = new GraphicsPath();
            toothPath.AddRectangle(new RectangleF(-toothW / 2f, -(innerR + toothH), toothW, toothH));

            using var m = new Matrix();
            m.RotateAt((float)(angle * 180.0 / Math.PI), PointF.Empty);
            m.Translate(cx, cy, MatrixOrder.Append);
            toothPath.Transform(m);

            path.AddPath(toothPath, false);
        }

        // Ring body (annulus represented as filled circle; hole punched below)
        path.AddEllipse(cx - innerR - 0.5f, cy - innerR - 0.5f, (innerR + 0.5f) * 2, (innerR + 0.5f) * 2);

        g.FillPath(brush, path);

        // Punch the centre hole with transparent composite
        using var holeBrush = new SolidBrush(Color.FromArgb(0, 0, 0, 0));
        var previous = g.CompositingMode;
        g.CompositingMode = CompositingMode.SourceCopy;
        g.FillEllipse(holeBrush, cx - innerR * 0.55f, cy - innerR * 0.55f, innerR * 1.1f, innerR * 1.1f);
        g.CompositingMode = previous;
    }

    /// <summary>
    /// Fills a rounded rectangle.
    /// </summary>
    private static void FillRoundedRectangle(Graphics g, Brush brush, float x, float y, float w, float h, float r)
    {
        using var path = new GraphicsPath();
        path.AddArc(x,          y,          r * 2, r * 2, 180, 90);
        path.AddArc(x + w - r*2, y,         r * 2, r * 2, 270, 90);
        path.AddArc(x + w - r*2, y + h - r*2, r * 2, r * 2, 0,   90);
        path.AddArc(x,          y + h - r*2, r * 2, r * 2, 90,  90);
        path.CloseFigure();
        g.FillPath(brush, path);
    }
}
