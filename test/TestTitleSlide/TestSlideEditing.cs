using DocumentFormat.OpenXml.Packaging;
using DocLayer.Core;

namespace DocLayer.Core.Examples
{
    public static class TestSlideEditing
    {
        public static void Run()
        {
            string sourcePath = @"C:\Users\shrey\OneDrive\Desktop\docs\2024.10.27 Project Core - Valuation Analysis_v23.pptx";
            string outputPath = @"C:\Users\shrey\OneDrive\Desktop\docs\Test_SlideEditing_Output.pptx";

            Console.WriteLine($"Loading presentation from: {sourcePath}");
            Console.WriteLine();

            // Create a copy for modification
            File.Copy(sourcePath, outputPath, true);
            Console.WriteLine($"Created working copy: {outputPath}");
            Console.WriteLine();

            // Open the copy for editing
            using (PresentationDocument presentationDocument = PresentationDocument.Open(outputPath, true))
            {
                var builder = new PresentationBuilder(presentationDocument, outputPath);

                // Test 1: Extract content from slide 1
                Console.WriteLine("=== TEST 1: Extract Slide Content ===");
                TestExtractSlideContent(builder, 1);
                Console.WriteLine();

                // Test 2: Edit text on slide 1
                Console.WriteLine("=== TEST 2: Edit Slide Text ===");
                TestEditSlideText(builder, 1);
                Console.WriteLine();

                // Test 3: Extract content from slide 3
                Console.WriteLine("=== TEST 3: Extract Slide 3 Content ===");
                TestExtractSlideContent(builder, 3);
                Console.WriteLine();

                // Test 4: Refresh slide 3 with new project details
                Console.WriteLine("=== TEST 4: Refresh Slide 3 with New Project ===");
                TestRefreshSlide3(builder);
                Console.WriteLine();

                Console.WriteLine($"✓ All tests completed. Output saved to: {outputPath}");
            }

            // Test 5: Render slide to image (must be done after closing the document)
            Console.WriteLine();
            Console.WriteLine("=== TEST 5: Render Slide to Image ===");
            TestRenderSlideToImageStandalone(outputPath, 3);
        }

        private static void TestExtractSlideContent(PresentationBuilder builder, int slideNumber)
        {
            try
            {
                var content = builder.ExtractSlideContent(slideNumber);

                Console.WriteLine($"Slide {slideNumber} Content:");
                Console.WriteLine($"  Shapes: {content.Shapes.Count}");
                
                foreach (var shape in content.Shapes)
                {
                    Console.WriteLine($"    - {shape.Name ?? "(null name)"}");
                    if (!string.IsNullOrEmpty(shape.Text))
                    {
                        var preview = shape.Text.Length > 50 
                            ? shape.Text.Substring(0, 47) + "..." 
                            : shape.Text;
                        Console.WriteLine($"      Text: \"{preview}\"");
                    }
                    if (shape.Position != null)
                    {
                        Console.WriteLine($"      Position: ({shape.Position.X}, {shape.Position.Y})");
                    }
                    if (shape.Size != null)
                    {
                        Console.WriteLine($"      Size: {shape.Size.Width} x {shape.Size.Height}");
                    }
                }

                Console.WriteLine($"  Pictures: {content.Pictures.Count}");
                foreach (var picture in content.Pictures)
                {
                    Console.WriteLine($"    - {picture.Name ?? "(null name)"}");
                    if (picture.Position != null)
                    {
                        Console.WriteLine($"      Position: ({picture.Position.X}, {picture.Position.Y})");
                    }
                    if (picture.Size != null)
                    {
                        Console.WriteLine($"      Size: {picture.Size.Width} x {picture.Size.Height}");
                    }
                }

                Console.WriteLine("✓ Extract test passed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Extract test failed: {ex.Message}");
                Console.WriteLine($"   Stack trace: {ex.StackTrace}");
            }
        }

        private static void TestEditSlideText(PresentationBuilder builder, int slideNumber)
        {
            try
            {
                // First, extract to see what shapes are available
                var content = builder.ExtractSlideContent(slideNumber);
                
                if (content.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found to edit");
                    return;
                }

                // Try to edit the first shape with text
                var shapeToEdit = content.Shapes.FirstOrDefault(s => !string.IsNullOrEmpty(s.Text));
                
                if (shapeToEdit == null)
                {
                    Console.WriteLine("No shapes with text found to edit");
                    return;
                }

                Console.WriteLine($"Editing shape: {shapeToEdit.Name}");
                Console.WriteLine($"  Original text: \"{shapeToEdit.Text}\"");
                
                string newText = "TEST: This text was edited programmatically!";
                builder.EditSlideText(slideNumber, shapeToEdit.Name, newText);
                
                Console.WriteLine($"  New text: \"{newText}\"");
                Console.WriteLine("✓ Edit test passed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Edit test failed: {ex.Message}");
            }
        }

        private static void TestRefreshSlide3(PresentationBuilder builder)
        {
            try
            {
                Console.WriteLine("Refreshing slide 3 with new project details...");
                
                // New fictional project details
                string newProjectName = "Project Phoenix";
                string newAnalysisType = "Strategic Assessment";
                string newDate = "December 2025";
                string currentDate = "12/15/2025";

                // Update title (shape name: "Title 1")
                Console.WriteLine($"  Updating title to: {newProjectName}");
                builder.EditSlideText(3, "Title 1", newProjectName);

                // Update subtitle (shape name: "Subtitle 2")
                string subtitle = $"{newAnalysisType}\n{newDate}";
                Console.WriteLine($"  Updating subtitle to: {newAnalysisType} / {newDate}");
                builder.EditSlideText(3, "Subtitle 2", subtitle);

                // Update date placeholder (shape name: "Date Placeholder 6")
                Console.WriteLine($"  Updating date to: {currentDate}");
                builder.EditSlideText(3, "Date Placeholder 6", currentDate);

                Console.WriteLine("✓ Slide 3 refresh test passed");
                Console.WriteLine($"  New project: {newProjectName}");
                Console.WriteLine($"  New analysis: {newAnalysisType}");
                Console.WriteLine($"  New date: {newDate}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Slide 3 refresh test failed: {ex.Message}");
                Console.WriteLine($"   Stack trace: {ex.StackTrace}");
            }
        }

        private static void TestRenderSlideToImageStandalone(string pptxPath, int slideNumber)
        {
            try
            {
                Console.WriteLine($"Rendering slide {slideNumber} to image...");
                
                // Use Syncfusion helper directly since document needs to be closed
                string imagePath = InternalUtilities.Syncfusion.SyncfusionHelperMethods.ExportSlideToImage(pptxPath, slideNumber);
                
                if (File.Exists(imagePath))
                {
                    var fileInfo = new FileInfo(imagePath);
                    Console.WriteLine($"✓ Render test passed");
                    Console.WriteLine($"  Image saved to: {imagePath}");
                    Console.WriteLine($"  File size: {fileInfo.Length / 1024} KB");
                }
                else
                {
                    Console.WriteLine("✗ Render test failed: Image file was not created");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Render test failed: {ex.Message}");
            }
        }
    }
}
