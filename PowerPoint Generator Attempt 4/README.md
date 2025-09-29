# PowerPoint Generator (Attempt 4)

A small utility to generate PowerPoint presentations from JSON data and JSON templates using the PowerPoint Interop API.

## What this does

- Loads `data.json` (presentation content) and `templates.json` (layout & styling).
- Creates a new PowerPoint presentation using `Microsoft.Office.Interop.PowerPoint`.
- Places text, images, background and formatting according to the template definitions.
- Saves the generated `.pptx` to your Desktop and writes basic file metadata.

## Requirements

- Windows (the project relies on COM interop with Microsoft PowerPoint and some Windows-only APIs).
- Microsoft PowerPoint installed (required for `Microsoft.Office.Interop.PowerPoint`).
- .NET SDK 8.0 (project targets `net8.0`).
- NuGet packages referenced in the project (these are included in the project file):
  - Microsoft.Office.Interop.PowerPoint
  - Microsoft.Windows.Compatibility
  - System.Drawing.Common
  - System.Reactive.Windows.Forms
  - ThammimTech.Microsoft.Office.Core
  - WindowsAPICodePack-Shell

## Files of interest

- `Program.cs` — main program, orchestrates loading data/templates and generating the presentation.
- `data.json` — example data (presentation name, background paths, slides array).
- `templates.json` — layout and shape definitions used when placing content on slides.
- `Templates/` — helper classes (`Template.cs`, `Content.cs`, `Formatting.cs`) that implement placement and formatting logic.

## How to build

1. Open the solution in Visual Studio 2022/2023, or use the command line with the .NET SDK 8 installed.
2. From the project directory (where `PowerPoint Generator Attempt 4.csproj` lives) run:

```powershell
dotnet restore
dotnet build -c Debug
dotnet run --project "PowerPoint Generator Attempt 4.csproj"
```

Note: run in a PowerShell / Windows environment. If running from Visual Studio, start the project normally (the build will pull the referenced packages).

## How to use

1. Edit `data.json` to contain your presentation content. Minimal structure (see full file for example):

```json
{
  "Name": "Presentation name",
  "Background": "C:\\path\\to\\background.png",
  "Background_blur": "C:\\path\\to\\background_blur.png",
  "Font": "Arial",
  "_color": [255,255,255],
  "Slides": [ { "Layout": "Title", "Title": "Hello", "Content": ["line1"], "Image": null } ]
}
```

2. Edit `templates.json` to adjust layout properties. Each layout contains two `Shape` objects (TitleShape and ContentShape) and optional `ImageShape`. Important shape properties:

- `Left`, `Top`, `Width`, `Height` — position and size (points).
- `_color` — RGB color array (e.g. `[255,255,255]`).
- `Font`, `FontSize` — font family & size.
- `BorderSize`, `Transparency`, `CenterH`, `CenterV`, `Bullets`, `_bulletChar`, `Padding` — visual formatting.

See `Templates/Template.cs` for the model used to deserialize `templates.json`.

3. Run the program. By default the app reads the hard-coded paths at the top of `Program.cs` for `data.json` and `templates.json`. The generated `.pptx` is saved to your Desktop as `{data.Name}.pptx`.

## Customization notes

- To pick a data file at runtime, uncomment/use the `GetFilePathFromUser()` OpenFileDialog code in `Program.cs`.
- If you want the generator to place a full-slide image behind other content, the code inserts a rectangle shape sized to the slide and fills it with `UserPicture`, then sends it to the back.
- The code currently measures text sizes using `System.Drawing.Graphics.MeasureString`. That works on Windows desktop environments; if you need cross-platform measurement, consider replacing the `TextMeasurement` implementation with SkiaSharp (example notes in code comments).

## Troubleshooting

- NullReferenceException when iterating slides: ensure `data.json` `Slides` array exists and the path used to load `data.json` is correct. The program prints `No slides found in data.json` if it's null.
- `PlatformNotSupportedException: Text measurement is only supported on Windows.` — ensure you're running on Windows and building with the Windows desktop/.NET SDK. If the exception originates from `Template.TextMeasurement`, confirm the environment is Windows and fonts used are installed. To support other platforms, use SkiaSharp as documented in the source comments.
- Font mismatch / weird sizes: ensure the font names used in `data.json` and `templates.json` are installed on your machine. If a font isn't found, measurement and rendering will fall back and sizes may change.
- PowerPoint COM errors: PowerPoint must be installed and available for COM interop. Run the app on a machine with PowerPoint and ensure no modal PowerPoint dialogs block automation.

## Logging & Debugging

- The program writes some console messages (e.g. invalid slide types, missing templates). Consider adding more `Console.WriteLine` calls around JSON loading to validate the content if something behaves unexpectedly.

## Where the output goes

- By default the presentation is saved to the current user's Desktop at `C:\Users\<you>\Desktop\{data.Name}.pptx`. You can change the save path in `Program.cs` before `presentation.SaveAs(savePath)`.

## License & Attribution

This repository contains user code. The project uses Microsoft Office Interop libraries which are subject to Microsoft licensing. The sample code here is provided as-is; if you redistribute or publish, take care to comply with dependent libraries' licenses.

## Next steps / suggestions

- I want global restructure to improve readability, maintainability and have better ability to create different templates.
- Add a small CLI to accept `--data <file>` and `--templates <file>` to avoid hard-coded paths.
- Add unit tests for JSON parsing and basic template validation.
- Consider SkiaSharp for cross-platform text measurement if plan to run outside Windows.

