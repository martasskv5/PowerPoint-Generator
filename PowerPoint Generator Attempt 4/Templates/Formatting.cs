using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Data = PowerPoint_Generator_Attempt_4.Data;
using Slide = PowerPoint_Generator_Attempt_4.Data.Slide;
using Template = PowerPoint_Generator_Attempt_4.Templates.Template;
using Shape = PowerPoint_Generator_Attempt_4.Templates.Template.Shape;

namespace PowerPoint_Generator_Attempt_4.Templates;

public class Formatting
{
    /// <summary>
    /// Set the background of a slide to an image.
    /// </summary>
    /// <param name="s">Slide to set background</param>
    /// <param name="data">Presentation data</param>
    public static void SetBackground(PowerPoint.Slide s, Data data)
    {
        s.FollowMasterBackground = Office.MsoTriState.msoFalse; // Ensure the slide does not follow the master background
        s.Background.Fill.UserPicture(data.Background_blur); // Actual background
        // Add a new shape for the image that covers the entire slide
        PowerPoint.Shape imageShape = s.Shapes.AddShape(
            Office.MsoAutoShapeType.msoShapeRectangle, // Shape type
            0, // Left position
            0, // Top position
            s.Master.Width, // Width
            s.Master.Height  // Height
        );

        imageShape.Fill.UserPicture(data.Background); // Insert an image into the new shape

        imageShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack); // Send the new shape to the back

        imageShape.Line.Visible = Office.MsoTriState.msoFalse; // Hide the border
    }
    /// <summary>
    /// Sets the position, size, and border of a shape.
    /// </summary>
    /// <param name="shape">Shape to format</param>
    /// <param name="sData">Shape data from template</param>
    /// <param name="data">Presentation data</param>
    public static void FormatShape(PowerPoint.Shape shape, Shape sData, Data data)
    {
        //Set the position and size of the shape
        shape.Left = sData.Left * sData.cmToPoints;
        shape.Top = sData.Top * sData.cmToPoints;
        shape.Height = sData.Height * sData.cmToPoints;
        shape.Width = sData.Width * sData.cmToPoints;
        shape.AutoShapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle;

        //Set Border style
        shape.Line.Visible = Office.MsoTriState.msoCTrue;
        if (sData.Color != null) { shape.Line.ForeColor.RGB = (int)sData.Color; }
        else { shape.Line.ForeColor.RGB = data.Color; }
        shape.Line.Weight = sData.BorderSize;
        shape.Line.Transparency = sData.Transparency;

        if (sData.CenterV) { shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle; } // Center vertically
    }
    /// <summary>
    /// Sets the position, size, and fill of an image shape.
    /// </summary>
    /// <param name="shape"></param>
    /// <param name="sData"></param>
    /// <param name="data"></param>
    /// <param name="image"></param>
    public static void FormatImageShape(PowerPoint.Shape shape, Shape sData, Data data, string image)
    {
        // Set the maximum width and height for the image
        float maxWidth = sData.Width * sData.cmToPoints;
        float maxHeight = sData.Height * sData.cmToPoints;

        // Load the image to get its original dimensions
        using System.Drawing.Image img = System.Drawing.Image.FromFile(image);
        float originalWidth = img.Width;
        float originalHeight = img.Height;

        float aspectRatio = originalWidth / originalHeight; // Calculate the aspect ratio

        // Calculate the new dimensions while keeping the aspect ratio
        float newWidth = maxWidth;
        float newHeight = maxHeight;

        if (originalWidth > originalHeight)
        {
            newHeight = maxWidth / aspectRatio;
            if (newHeight > maxHeight)
            {
                newHeight = maxHeight;
                newWidth = maxHeight * aspectRatio;
            }
        }
        else
        {
            newWidth = maxHeight * aspectRatio;
            if (newWidth > maxWidth)
            {
                newWidth = maxWidth;
                newHeight = maxWidth / aspectRatio;
            }
        }

        // Set the position and size of the shape
        shape.Left = sData.Left * sData.cmToPoints + (maxWidth - newWidth) / 2;
        shape.Top = sData.Top * sData.cmToPoints + (maxHeight - newHeight) / 2;
        shape.Height = newHeight;
        shape.Width = newWidth;
        shape.AutoShapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle;
        shape.Line.Visible = Office.MsoTriState.msoFalse; // Hide the border

        shape.Fill.UserPicture(image);// Fill the shape with the image
    }
    /// <summary>
    /// Sets font, size and color of text in a shape. Optionally sets bullet points.  
    /// </summary>
    /// <param name="textRange">Text to format</param>
    /// <param name="sData">Shape data form template</param>
    /// <param name="data">Presentation data</param>
    public static void FormatText(Office.TextRange2 textRange, Shape sData, Data data)
    {
        textRange.Font.Name = data.Font; // Font name
        textRange.Font.Size = sData.FontSize; // Font size
        textRange.Font.Fill.ForeColor.RGB = data.Color; // Font color

        SetBullet(textRange, sData); // Set bullet point properties if bullets are to be used

        // Add a shadow to the text
        textRange.Font.Shadow.Visible = Office.MsoTriState.msoTrue;
        textRange.Font.Shadow.Type = Office.MsoShadowType.msoShadow1;
        textRange.Font.Shadow.Transparency = 0.4f;
        textRange.Font.Shadow.Size = 102;
        textRange.Font.Shadow.Blur = 5.0f;
        textRange.Font.Shadow.OffsetX = 3.0f;
        textRange.Font.Shadow.OffsetY = 3.0f;

        if (sData.CenterH) { textRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignCenter; } // Center horizontally
        else { textRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignLeft; } // Align left
    }
    /// <summary>
    /// Apply a gradient fill to a top text shape.
    /// </summary>
    /// <param name="shape">Top text shape</param>
    /// <param name="color">Color of gradient in OLE format</param>
    public static void CreateGradientFill(PowerPoint.Shape shape, int color)
    {
        shape.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1); // Set the fill to a gradient

        // Insert 4 gradient stops
        shape.Fill.GradientStops.Insert(color, 0.0f);
        shape.Fill.GradientStops.Insert(color, 0.2f, 0.9f);
        shape.Fill.GradientStops.Insert(color, 0.75f, 0.9f);
        shape.Fill.GradientStops.Insert(color, 1.0f);

        shape.Fill.GradientAngle = 45; // Set the gradient angle
    }
    /// <summary>
    /// Apply a transition to a slide.
    /// </summary>
    /// <param name="slide">Slide to apply transition</param>
    /// <param name="layout">Slide layout from template</param>
    public static void ApplyTransition(PowerPoint.Slide slide, Template.Layout layout)
    {
        PowerPoint.SlideShowTransition transition = slide.SlideShowTransition;
        transition.EntryEffect = GetEntryEffect(layout.Transition);
        transition.Duration = layout.Duration;
        if (layout.AdvanceOnTime) // Check if the slide should advance automatically
        {
            transition.AdvanceOnTime = layout.AdvanceOnTime ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            transition.AdvanceTime = layout.AdvanceTime ?? 1f;
        }
    }
    /// <summary>
    /// Get the entry effect (transition) based on the provided effect name.
    /// More info https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
    /// </summary>
    /// <param name="effectName">Name of effect from template</param>
    /// <returns>PowerPoint Transition effect</returns>
    private static PowerPoint.PpEntryEffect GetEntryEffect(string effectName)
    {
        try
        {
            string fullEffectName = "ppEffect" + effectName; // Combine the prefix "ppEffect" with the provided effect name

            return (PowerPoint.PpEntryEffect)Enum.Parse(typeof(PowerPoint.PpEntryEffect), fullEffectName, true); // Parse the combined name to get the corresponding enum value
        }
        catch
        {
            Console.WriteLine($"Entry effect '{effectName}' not found. Defaulting to 'ppEffectNone'.");
            return PowerPoint.PpEntryEffect.ppEffectNone;
        }
    }
    /// <summary>
    /// Set the bullet character to a character from template.
    /// Unicode character list https://en.wikipedia.org/wiki/List_of_Unicode_characters
    /// </summary>
    /// <param name="text">TextRange to apply bullets</param>
    private static void SetBullet(Office.TextRange2 text, Shape sData)
    {
        if (sData.Bullets)
        {
            // Iterate through each paragraph in the text range
            for (int i = 1; i <= text.Paragraphs.Count; i++)
            {
                Office.TextRange2 paragraph = text.Paragraphs[i];
                Office.ParagraphFormat2 paragraphFormat = paragraph.ParagraphFormat;

                // Enable bullets and set the bullet character to an empty square
                paragraphFormat.Bullet.Type = Office.MsoBulletType.msoBulletUnnumbered;
                paragraphFormat.Bullet.Font.Name = "Arial";
                paragraphFormat.Bullet.Character = sData.BulletChar;

                paragraphFormat.Bullet.RelativeSize = 1.0f; // Adjust the size of the bullet relative to the text

            }
        }
        else
        {
            // Disable bullets if useBullets is false
            for (int i = 1; i <= text.Paragraphs.Count; i++)
            {
                Office.TextRange2 paragraph = text.Paragraphs[i];
                Office.ParagraphFormat2 paragraphFormat = paragraph.ParagraphFormat;

                paragraphFormat.Bullet.Type = Office.MsoBulletType.msoBulletNone;
            }
        }
    }
}
