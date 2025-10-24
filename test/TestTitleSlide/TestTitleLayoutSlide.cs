using DocumentFormat.OpenXml.Packaging;
using OpenXMLExtensions;

namespace DocLayer.Core.Examples
{
    public class TestTitleLayoutSlide
    {
        public static void Run()
        {
            string outputPath = "C:\\Users\\shrey\\projects\\doclayer\\test\\test_outputs\\test_title_layout_slide.pptx";

            using (var presentationDoc = PresentationHelper.CreatePresentation(outputPath, widescreen: true)) 
            {
                PresentationBuilder builder = new(presentationDoc);
                string title = "TEST";
                string subtitle = "TEST SLIDE";
                string footnote = "Placeholder for footnote";
                builder.CreateTitleSlide(title, subtitle, footnote);
            }

            Console.WriteLine($"✓ Title layout slide created successfully: {outputPath}");
            Console.WriteLine($"✓ File size: {new FileInfo(outputPath).Length} bytes");
            Console.WriteLine("\nOpen the file in PowerPoint to view the slide!");
        }
    }
}
