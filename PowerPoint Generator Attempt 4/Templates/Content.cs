using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Data = PowerPoint_Generator_Attempt_4.Data;
using Slide = PowerPoint_Generator_Attempt_4.Data.Slide;
using Template = PowerPoint_Generator_Attempt_4.Templates.Template;
using IndexManager = PowerPoint_Generator_Attempt_4.IndexManager;
using Layout = PowerPoint_Generator_Attempt_4.Templates.Template.Layout;
using Shape = PowerPoint_Generator_Attempt_4.Templates.Template.Shape;
using F = PowerPoint_Generator_Attempt_4.Templates.Formatting;

namespace PowerPoint_Generator_Attempt_4.Templates;

public class Content
{
    public static void CreateTitle(PowerPoint.Slide s, Slide slide, Data data, Shape titleShape)
    {
        // Insert text into the presentation slide
        PowerPoint.Shape title = s.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 200, 200);
        PowerPoint.Shape titleBottom = title.Duplicate()[1];
        PowerPoint.ShapeRange titleRange = s.Shapes.Range(new[] { title.Name, titleBottom.Name });
        //titleRange.Group(); // Attempt to group the shapes

        F.FormatShape(titleBottom, titleShape, data);

        titleBottom.Fill.Background(); //Set background

        F.FormatShape(title, titleShape, data);

        F.CreateGradientFill(title, data.Color); //Set background

        title.TextFrame2.TextRange.Text = slide.Title; // add title text
        F.FormatText(title.TextFrame2.TextRange, titleShape, data);
        title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
        title.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
    }

    public static void CreateContent(PowerPoint.Slide s, Slide slide, Data data, Shape contentShape, Layout layout, string layoutName)
    {
        string[] splitLayouts = ["SplitShortTextImageLeft", "SplitShortTextImageRight"];
        if (splitLayouts.Contains(layoutName)) // Handle different processing for split layouts
        {
            float originalTop = contentShape.Top;
            foreach (string p in slide.Content)
            {
                PowerPoint.Shape content = s.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 200, 200);
                PowerPoint.Shape contentBottom = content.Duplicate()[1];
                PowerPoint.ShapeRange contentRange = s.Shapes.Range(new[] { content.Name, contentBottom.Name });
                //contentRange.Group(); // Group the shapes
                F.FormatShape(contentBottom, contentShape, data);

                contentBottom.Fill.Background(); //Set background

                F.FormatShape(content, contentShape, data); //Format shape

                F.CreateGradientFill(content, data.Color); //Set fill

                content.TextFrame2.TextRange.Text = p; // add title text
                F.FormatText(content.TextFrame2.TextRange, contentShape, data);
                content.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                contentShape.Top += (float)(contentShape.Height + contentShape.Padding);
            }
            contentShape.Top = originalTop;
        }
        else
        {
            // Insert text into the presentation slide

            PowerPoint.Shape content = s.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 200, 200);
            PowerPoint.Shape contentBottom = content.Duplicate()[1];
            PowerPoint.ShapeRange contentRange = s.Shapes.Range(new[] { content.Name, contentBottom.Name });
            //contentRange.Group(); // Group the shapes
            F.FormatShape(contentBottom, contentShape, data);

            contentBottom.Fill.Background(); //Set background

            F.FormatShape(content, contentShape, data); //Format shape

            F.CreateGradientFill(content, data.Color); //Set fill

            // add title text

            content.TextFrame2.TextRange.Text = string.Empty; // Clear existing text

            // Determine the number of paragraphs to add
            int paragraphLimit = slide.Content.Length;
            if (layout.MaxParagraphs != null) { paragraphLimit = Math.Min((int)layout.MaxParagraphs, slide.Content.Length); }
            // Check if the content exceeds the paragraph limit
            //if (slide.Content.Length > paragraphLimit) { Console.WriteLine("Content exceeds the paragraph limit of 8. Only the first 8 paragraphs will be added."); }
            // Iterate through the content and append text without adding an empty line at the end
            for (int i = 0; i < paragraphLimit; i++)
            {
                content.TextFrame2.TextRange.Text += "  " + slide.Content[i];
                if (i < paragraphLimit - 1)
                {
                    content.TextFrame2.TextRange.Text += "\n";
                }
            }
            F.FormatText(content.TextFrame2.TextRange, contentShape, data);
            content.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }
    }

    public static void EditContent(PowerPoint.Shape content, Slide slide, Data data, Shape contentShape, Layout layout)
    {
        content.TextFrame2.TextRange.Text = string.Empty; // Clear existing text

        // Determine the number of paragraphs to add
        int paragraphLimit = slide.Content.Length;
        if (layout.MaxParagraphs != null) { paragraphLimit = Math.Min((int)layout.MaxParagraphs, slide.Content.Length); }
        // Check if the content exceeds the paragraph limit
        //if (slide.Content.Length > paragraphLimit) { Console.WriteLine("Content exceeds the paragraph limit of 8. Only the first 8 paragraphs will be added."); }
        // Iterate through the content and append text without adding an empty line at the end
        for (int i = 0; i < paragraphLimit; i++)
        {
            content.TextFrame2.TextRange.Text += "  " + slide.Content[i];
            if (i < paragraphLimit - 1)
            {
                content.TextFrame2.TextRange.Text += "\n";
            }
        }
        F.FormatText(content.TextFrame2.TextRange, contentShape, data);
    }

    public static void ProcessParagraphs(PowerPoint.Slide s, Slide slide, Data data, Shape contentShape, Layout layout, PowerPoint.Presentation presentation, string layoutName)
    {
        if (slide.Content.Length > layout.MaxParagraphs)
        {
            string[] originalContent = [.. slide.Content]; // Store original content
            int maxParagraphs = (int)layout.MaxParagraphs;
            int a = 0;
            while (originalContent.Length > 0)
            {
                if (a == 0) // Handle the first MaxParagraphs
                {
                    a++;
                    CreateContent(s, slide, data, contentShape, layout, layoutName);
                    originalContent = originalContent.Skip(maxParagraphs).ToArray();
                }
                else
                {
                    // Update slide.Content with the first MaxParagraphs
                    slide.Content = originalContent.Take(maxParagraphs).ToArray();

                    // Remove processed paragraphs from original content
                    originalContent = originalContent.Skip(maxParagraphs).ToArray();

                    // Process the updated slide (this part depends on your specific requirements)
                    // For example, you might want to update the layout or perform some other operation
                    PowerPoint.Slide newS = presentation.Slides[IndexManager.Instance.Index].Duplicate()[1];
                    IndexManager.Instance.Index++;
                    List<PowerPoint.Shape> shapesWithText = new List<PowerPoint.Shape>();

                    foreach (PowerPoint.Shape shape in newS.Shapes)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                        {
                            shapesWithText.Add(shape);
                        }
                    }
                    //Console.WriteLine(shapesWithText.Count);
                    for (int i = 1; i < shapesWithText.Count; i++)
                    {
                        EditContent(shapesWithText[i], slide, data, contentShape, layout);
                    }
                    //EditContent(newS.Shapes[5], slide, data, contentShape, layout);
                }
            }
        }
        else { CreateContent(s, slide, data, contentShape, layout, layoutName); }
    }

    public static void ProcessImage(Slide slide, PowerPoint.Slide s, Data data, Template templates, Shape contentShape)
    {
        if (slide.Image != null)
        {
            PowerPoint.Shape image = s.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 200, 200);
            if (templates.GetLayout(slide.Layout).ImageShape != null)
            {
                Shape imageShape = templates.GetLayout(slide.Layout).ImageShape;
                if (imageShape.Bullets) // using existing value to set bg because lazy
                {
                    PowerPoint.Shape imageBG = image.Duplicate()[1];
                    F.FormatShape(imageBG, imageShape, data);
                    imageBG.Fill.Background();
                    image.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
                F.FormatImageShape(image, imageShape, data, slide.Image);
            }
            else
            {
                Console.WriteLine("Image shape not defined in template. Using default settings.");
                F.FormatImageShape(image, contentShape, data, slide.Image);
            }
            //F.FormatImageShape(image, contentShape, data, slide.Image);
        }
        else { Console.WriteLine("This slide type does not support images."); }
    }
}