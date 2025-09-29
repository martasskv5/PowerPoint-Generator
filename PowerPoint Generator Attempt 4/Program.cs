using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Data = PowerPoint_Generator_Attempt_4.Data;
using Slide = PowerPoint_Generator_Attempt_4.Data.Slide;
using Template = PowerPoint_Generator_Attempt_4.Templates.Template;
using Shape = PowerPoint_Generator_Attempt_4.Templates.Template.Shape;
using Layout = PowerPoint_Generator_Attempt_4.Templates.Template.Layout;
using C = PowerPoint_Generator_Attempt_4.Templates.Content;
using F = PowerPoint_Generator_Attempt_4.Templates.Formatting;
using PowerPoint_Generator_Attempt_4;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System;
using System.Windows.Forms;



namespace PowerPoint_Generator_Attempt_4
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //string imagePath = @"C:\Users\martinko\OneDrive - MSFT\Documents\csharp\PowerPoint Generator Attempt 4\PowerPoint Generator Attempt 4\Green.png";
            string dataPath = @"C:\Users\martinko\OneDrive - MSFT\Documents\csharp\PowerPoint Generator Attempt 4\PowerPoint Generator Attempt 4\data.json";
            // string dataPath = GetFilePathFromUser();
            if (string.IsNullOrEmpty(dataPath))
            {
                Console.WriteLine("No file selected. Exiting program.");
                return;
            }
            string templatePath = @"C:\Users\martinko\OneDrive - MSFT\Documents\csharp\PowerPoint Generator Attempt 4\PowerPoint Generator Attempt 4\templates.json";
            IndexManager.Instance.Index = 1;
            Data data = Data.LoadFromJson(dataPath);
            //Console.WriteLine(data.Background);
            Template templates = Template.LoadFromJson(templatePath);
            //Console.WriteLine($"{data._color}, {data.color}");
            // Create an instance of PowerPoint application
            PowerPoint.Application powerpointApp = new PowerPoint.Application();
            // Create powerpoint presentation
            PowerPoint.Presentation presentation = powerpointApp.Presentations.Add();

            string[] slideTypes1 = ["Title", "LongText", "ShortTextImageRight", "ShortTextImageLeft", "SingleLongImageBottom", "SplitShortTextImageLeft", "SplitShortTextImageRight"];

            if (data.Slides != null)
            {
                foreach (Slide slide in data.Slides)
                {
                    // Customize the presentation
                    // Add slides, insert content, apply formatting, etc.
                    // Add a new slide
                    if (slideTypes1.Contains(slide.Layout))
                    {
                        PowerPoint.Slide s = presentation.Slides.Add(IndexManager.Instance.Index, PowerPoint.PpSlideLayout.ppLayoutBlank);
                        Layout layout = templates.GetLayout(slide.Layout);
                        Shape titleShape = layout.TitleShape;
                        Shape contentShape = layout.ContentShape;

                        // ----------------- Title -----------------
                        C.CreateTitle(s, slide, data, titleShape);
                        // ----------------- Title -----------------
                        F.SetBackground(s, data);
                        C.ProcessImage(slide, s, data, templates, contentShape);
                        // ----------------- Content -----------------
                        C.ProcessParagraphs(s, slide, data, contentShape, layout, presentation, slide.Layout);
                        // ----------------- Content -----------------

                        // Apply transition to the slide
                        F.ApplyTransition(s, layout);

                    }
                    else { Console.WriteLine("Invalid slide type. Try fixing your data.json file"); }
                    IndexManager.Instance.Index++;
                }

                // Save the presentation
                string savePath = $"C:\\Users\\martinko\\Desktop\\{data.Name}.pptx";
                presentation.SaveAs(savePath);

                presentation.Close();
                powerpointApp.Quit();

                ShellObject shellObject = ShellObject.FromParsingName(savePath);
                // Set new metadata
                shellObject.Properties.System.Title.Value = data.Name;
                shellObject.Properties.System.Comment.Value = "Created with PowerPoint Generator v4 by Marťas_SK";

            }
            else { Console.WriteLine("No slides found in data.json"); }

            // static string GetFilePathFromUser()
            // {
            //     using OpenFileDialog openFileDialog = new OpenFileDialog();
            //     openFileDialog.InitialDirectory = @"C:\Users\martinko\OneDrive - MSFT\Documents\csharp\PowerPoint Generator Attempt 4\PowerPoint Generator Attempt 4";
            //     openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
            //     openFileDialog.FilterIndex = 1;
            //     openFileDialog.RestoreDirectory = true;

            //     if (openFileDialog.ShowDialog() == DialogResult.OK)
            //     {
            //         // Get the path of specified file
            //         return openFileDialog.FileName;
            //     }

            //     return null;
            // }
        }

    }


    public class IndexManager
    {
        private static IndexManager _instance;
        private int _index = 1;

        private IndexManager() { }

        public static IndexManager Instance
        {
            get
            {
                _instance ??= new IndexManager();
                return _instance;
            }
        }

        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }
    }
}