using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

// Draws the SPOReplaceLink icon on a 256 grid: a spreadsheet on a SharePoint-teal to Excel-green tile with a link badge.
// Small sizes (up to 24 px) get a simplified drawing, only the chain, so they stay readable.
// Writes app.ico (PNG frames 16-256), app.svg (full drawing), app-small.svg and icon-<size>.png to look at.
// Copy app.ico to src\SPOReplaceLink.Desktop, app.svg to src\SPOReplaceLink.Ui\wwwroot\logo.svg (top bar of the app)
// and app-small.svg next to this file.
if (args.Length != 1)
{
    Console.Error.WriteLine("dotnet run --project tools/IconGen -- <output folder>");
    return 2;
}
string outDir = args[0];
Directory.CreateDirectory(outDir);
int[] sizes = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256];
var frames = new List<(int Size, byte[] Png)>();
foreach (int size in sizes)
{
    var canvas = new Canvas();
    var visual = new DrawingVisual();
    using (var dc = visual.RenderOpen())
    {
        dc.PushTransform(new ScaleTransform(size / 256.0, size / 256.0));
        canvas.Dc = dc;
        Icon.Draw(canvas, size <= 24);
    }
    var bmp = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
    bmp.Render(visual);
    var enc = new PngBitmapEncoder();
    enc.Frames.Add(BitmapFrame.Create(bmp));
    using var ms = new MemoryStream();
    enc.Save(ms);
    frames.Add((size, ms.ToArray()));
    File.WriteAllBytes(Path.Combine(outDir, $"icon-{size}.png"), ms.ToArray());
}

using (var ico = new BinaryWriter(File.Create(Path.Combine(outDir, "app.ico"))))
{
    ico.Write((ushort)0);
    ico.Write((ushort)1);
    ico.Write((ushort)frames.Count);
    int offset = 6 + 16 * frames.Count;
    foreach (var (size, png) in frames)
    {
        ico.Write((byte)(size >= 256 ? 0 : size));
        ico.Write((byte)(size >= 256 ? 0 : size));
        ico.Write((byte)0);
        ico.Write((byte)0);
        ico.Write((ushort)1);
        ico.Write((ushort)32);
        ico.Write(png.Length);
        ico.Write(offset);
        offset += png.Length;
    }
    foreach (var (_, png) in frames)
        ico.Write(png);
}

foreach (var (name, small) in new[] { ("app.svg", false), ("app-small.svg", true) })
{
    var svg = new Canvas();
    Icon.Draw(svg, small);
    File.WriteAllText(Path.Combine(outDir, name), svg.Svg());
}
Console.WriteLine($"Записано в {outDir}");
return 0;

static class Icon
{
    private const string Badge = "#0F6CBD";
    private const string Excel = "#107C41";

    public static void Draw(Canvas c, bool small)
    {
        c.Gradient("bg", "#05727A", "#0E7A3E");
        if (small)
        {
            c.Rect(4, 4, 248, 248, 56, "url(#bg)");
            Chain(c, 128, 128, scale: 2.35, stroke: 25);
            return;
        }
        c.Rect(8, 8, 240, 240, 52, "url(#bg)");

        // Spreadsheet with a folded corner.
        c.Path("M66,40 H148 L182,74 V186 A14,14 0 0 1 168,200 H66 A14,14 0 0 1 52,186 V54 A14,14 0 0 1 66,40 Z", "#000000", opacity: 0.18, dx: 3, dy: 5);
        c.Path("M66,40 H148 L182,74 V186 A14,14 0 0 1 168,200 H66 A14,14 0 0 1 52,186 V54 A14,14 0 0 1 66,40 Z", "#FFFFFF");
        c.Path("M148,40 V62 A12,12 0 0 0 160,74 H182 Z", "#C9E7D6");
        c.Rect(68, 58, 62, 16, 5, Excel);
        foreach (int y in new[] { 88, 112, 136, 160 })
        {
            c.Rect(68, y, 30, 14, 4, "#21A366");
            c.Rect(106, y, 60, 14, 4, "#C4E5D2");
        }

        // Link badge.
        c.Circle(182, 182, 62, "#FFFFFF");
        c.Circle(182, 182, 53, Badge);
        Chain(c, 182, 182, scale: 1, stroke: 10);
    }

    /// <summary>Two interlocked links at 45°, centred on (cx, cy); at scale 1 about 88 wide.</summary>
    private static void Chain(Canvas c, double cx, double cy, double scale, double stroke)
    {
        double w = 48 * scale, h = 26 * scale, off = 17 * scale;
        c.Rect(cx - off - w / 2, cy - h / 2, w, h, h / 2, null, stroke: "#FFFFFF", strokeWidth: stroke, rotate: (-45, cx, cy));
        c.Rect(cx + off - w / 2, cy - h / 2, w, h, h / 2, null, stroke: "#FFFFFF", strokeWidth: stroke, rotate: (-45, cx, cy));
    }
}

/// <summary>Draws to WPF when Dc is set and always records the same shapes as SVG.</summary>
sealed class Canvas
{
    public DrawingContext? Dc;
    private readonly StringBuilder _body = new();
    private readonly StringBuilder _defs = new();
    private readonly Dictionary<string, Brush> _gradients = [];

    private static string F(double v) => v.ToString("0.##", CultureInfo.InvariantCulture);

    public void Gradient(string id, string from, string to)
    {
        _defs.Append($"<linearGradient id=\"{id}\" x1=\"0\" y1=\"0\" x2=\"1\" y2=\"1\"><stop offset=\"0\" stop-color=\"{from}\"/><stop offset=\"1\" stop-color=\"{to}\"/></linearGradient>");
        _gradients[$"url(#{id})"] = new LinearGradientBrush(Hex(from), Hex(to), new Point(0, 0), new Point(1, 1));
    }

    public void Rect(double x, double y, double w, double h, double r, string? fill, string? stroke = null, double strokeWidth = 0,
        double opacity = 1, (double A, double X, double Y)? rotate = null)
    {
        _body.Append($"<rect x=\"{F(x)}\" y=\"{F(y)}\" width=\"{F(w)}\" height=\"{F(h)}\" rx=\"{F(r)}\"{Attrs(fill, stroke, strokeWidth, opacity, rotate)}/>");
        Draw(new RectangleGeometry(new Rect(x, y, w, h), r, r), fill, stroke, strokeWidth, opacity, rotate);
    }

    public void Circle(double cx, double cy, double r, string fill)
    {
        _body.Append($"<circle cx=\"{F(cx)}\" cy=\"{F(cy)}\" r=\"{F(r)}\"{Attrs(fill, null, 0, 1, null)}/>");
        Draw(new EllipseGeometry(new Point(cx, cy), r, r), fill, null, 0, 1, null);
    }

    public void Path(string d, string fill, double opacity = 1, double dx = 0, double dy = 0)
    {
        string move = dx != 0 || dy != 0 ? $" transform=\"translate({F(dx)} {F(dy)})\"" : "";
        _body.Append($"<path d=\"{d}\"{Attrs(fill, null, 0, opacity, null)}{move}/>");
        var g = Geometry.Parse(d).Clone();
        g.Transform = new TranslateTransform(dx, dy);
        Draw(g, fill, null, 0, opacity, null);
    }

    private static string Attrs(string? fill, string? stroke, double sw, double opacity, (double A, double X, double Y)? rotate) =>
        $" fill=\"{fill ?? "none"}\"" +
        (stroke is null ? "" : $" stroke=\"{stroke}\" stroke-width=\"{F(sw)}\"") +
        (opacity < 1 ? $" opacity=\"{F(opacity)}\"" : "") +
        (rotate is { } t ? $" transform=\"rotate({F(t.A)} {F(t.X)} {F(t.Y)})\"" : "");

    private void Draw(Geometry g, string? fill, string? stroke, double sw, double opacity, (double A, double X, double Y)? rotate)
    {
        if (Dc is null)
            return;
        if (rotate is { } t)
            Dc.PushTransform(new RotateTransform(t.A, t.X, t.Y));
        Dc.PushOpacity(opacity);
        Brush? brush = fill is null ? null : _gradients.TryGetValue(fill, out var b) ? b : new SolidColorBrush(Hex(fill));
        Pen? pen = stroke is null ? null : new Pen(new SolidColorBrush(Hex(stroke)), sw);
        Dc.DrawGeometry(brush, pen, g);
        Dc.Pop();
        if (rotate is not null)
            Dc.Pop();
    }

    private static Color Hex(string s) => (Color)ColorConverter.ConvertFromString(s);

    public string Svg() =>
        $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 256 256\" width=\"256\" height=\"256\"><defs>{_defs}</defs>{_body}</svg>\n";
}
