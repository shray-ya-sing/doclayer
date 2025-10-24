using DocumentFormat.OpenXml.Packaging;
using OpenXMLExtensions;

namespace DocLayer.Core.Examples
{
    public class TestTheme
    {
        public static void Run()
        {
            string outputPath = "C:\\Users\\shrey\\projects\\doclayer\\test\\test_outputs\\test_theme.pptx";

            using (var presentationDoc = PresentationHelper.CreatePresentation(outputPath, widescreen: true))
            {
                PresentationBuilder builder = new(presentationDoc);

                // Set custom theme BEFORE creating slides
                Console.WriteLine("Setting presentation theme...");
                builder.SetPresentationTheme(
                    fontName: "Arial",
                    accentColors: new List<string>
                    {
                        "FF5733",  // Red-Orange (Accent 1)
                        "33FF57",  // Green (Accent 2)
                        "3357FF",  // Blue (Accent 3)
                        "F3FF33"   // Yellow (Accent 4)
                    }
                );
                Console.WriteLine("✓ Theme set successfully");

                // Create a title slide with the new theme
                Console.WriteLine("Creating title slide...");
                builder.CreateTitleSlide(
                    title: "Custom Theme Test",
                    subtitle: "Arial font with custom accent colors",
                    footnote: "Source: DocLayer.Core Theme Test"
                );
                Console.WriteLine("✓ Title slide created");
            }

            Console.WriteLine($"\n✓ Presentation created successfully: {outputPath}");
            Console.WriteLine($"✓ File size: {new FileInfo(outputPath).Length} bytes");
            Console.WriteLine("\nOpen the file in PowerPoint to view the custom theme!");
            Console.WriteLine("Check Design > Variants to see the custom accent colors");
        }
    }
}
