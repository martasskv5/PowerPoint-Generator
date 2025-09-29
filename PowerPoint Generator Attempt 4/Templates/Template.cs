using System.Drawing;
using System.Text.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Data = PowerPoint_Generator_Attempt_4.Data;
using Slide = PowerPoint_Generator_Attempt_4.Data.Slide;

namespace PowerPoint_Generator_Attempt_4.Templates;

public class Template
{
    public Layout Title { get; set; }
    public Layout LongText { get; set; }
    public Layout ShortTextImageLeft { get; set; }
    public Layout ShortTextImageRight { get; set; }
    public Layout SingleLongImageBottom { get; set; }
    public Layout SplitShortTextImageLeft { get; set; }
    public Layout SplitShortTextImageRight { get; set; }


    public class Layout
    {
        public Shape TitleShape { get; set; }
        public Shape ContentShape { get; set; }
        public Shape? ImageShape { get; set; }
        public string Transition { get; set; }
        public float Duration { get; set; }
        public bool AdvanceOnTime { get; set; }
        public float? AdvanceTime { get; set; }
        public int? MaxParagraphs { get; set; }

    }
    public class Shape
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Height { get; set; }
        public float Width { get; set; }
        public int[]? _color { get; set; }
        public string? Font { get; set; }
        public int FontSize { get; set; }
        public float BorderSize { get; set; }
        public float Transparency { get; set; }
        public bool CenterH { get; set; }
        public bool CenterV { get; set; }
        public bool Bullets { get; set; }
        public string? _bulletChar { get; set; }
        public float? Padding { get; set; }

        public float cmToPoints = 28.3465f;
        public int? Color { get { return _color != null ? (int)Data.RgbToOleColor(_color[0], _color[1], _color[2]) : null; } }
        public int BulletChar { get { return !string.IsNullOrEmpty(_bulletChar) ? Convert.ToInt32(_bulletChar, 16) : 0; } }

        public void UpdateSize(string text, string font)
        {
            string f;
            if (font != null)
            {
                f = font;
            }
            else
            {
                f = Font;
            }
            SizeF textSize = TextMeasurement.MeasureText(text, f, FontSize);
            Width = textSize.Width;
            Height = textSize.Height;
        }
    }

    public class TextMeasurement
    {
        public static SizeF MeasureText(string text, string fontName, float fontSize)
        {
            using Font font = new Font(fontName, fontSize);
            using Bitmap bitmap = new Bitmap(1, 1);

            using Graphics graphics = Graphics.FromImage(bitmap);
            return graphics.MeasureString(text, font);

        }
    }

    public static Template LoadFromJson(string filePath)
    {
        string jsonString = File.ReadAllText(filePath);
        Template data = JsonSerializer.Deserialize<Template>(jsonString) ?? throw new InvalidOperationException("Failed to deserialize JSON to Template object.");
        return data;
    }

    public Layout GetLayout(string layoutName)
    {
        return layoutName switch
        {
            "Title" => Title,
            "LongText" => LongText,
            "ShortTextImageLeft" => ShortTextImageLeft,
            "ShortTextImageRight" => ShortTextImageRight,
            "SingleLongImageBottom" => SingleLongImageBottom,
            "SplitShortTextImageLeft" => SplitShortTextImageLeft,
            "SplitShortTextImageRight" => SplitShortTextImageRight,
            _ => throw new ArgumentException($"Layout '{layoutName}' not found."),
        };

    }
}
