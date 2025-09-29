namespace PowerPoint_Generator_Attempt_4;
using System.Text.Json;

public class Data
{
    public string Name { get; set; }
    public string Background { get; set; }
    public string Background_blur { get; set; }
    public string Font { get; set; }
    public int[] _color { get; set; }
    public List<Slide>? Slides { get; set; }
    public int Color { get { return (int)RgbToOleColor(_color[0], _color[1], _color[2]); } }
    //public Template Templates { get; set; }


    public class Slide
    {
        public string Layout { get; set; }
        public string Title { get; set; }
        public string[] Content { get; set; }
        public string? Image { get; set; }
        public string? Font { get; set; }
    }

    public static long RgbToOleColor(int red, int green, int blue)
    {
        return ((long)red << 16) | ((long)green << 8) | blue;
    }

    public static Data LoadFromJson(string filePath)
    {
        string jsonString = File.ReadAllText(filePath);
        Data data = JsonSerializer.Deserialize<Data>(jsonString) ?? throw new InvalidOperationException("Failed to deserialize JSON to Data object.");
        return data;
    }
}
