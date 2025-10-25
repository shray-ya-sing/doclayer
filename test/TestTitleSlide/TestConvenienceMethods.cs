using DocLayer.Core;
using DocLayer.Core.Models;

namespace DocLayer.Core.Examples
{
    public static class TestConvenienceMethods
    {
        public static void Run()
        {
            string sourcePath = @"C:\Users\shrey\OneDrive\Desktop\docs\2024.10.27 Project Core - Valuation Analysis_v23.pptx";
            string outputPath = @"C:\Users\shrey\OneDrive\Desktop\docs\Test_ConvenienceMethods_Output.pptx";

            Console.WriteLine("Testing Convenience Methods for Agent Integration");
            Console.WriteLine("==================================================\n");

            // Create a working copy
            File.Copy(sourcePath, outputPath, true);
            Console.WriteLine($"Created working copy: {outputPath}\n");

            // Test 1: FromFile factory method with using statement
            Console.WriteLine("=== TEST 1: FromFile Factory Method ===");
            TestFromFileFactory(outputPath);
            Console.WriteLine();

            // Test 2: Get slide count
            Console.WriteLine("=== TEST 2: Get Slide Count ===");
            TestGetSlideCount(outputPath);
            Console.WriteLine();

            // Test 3: Extract all slides at once
            Console.WriteLine("=== TEST 3: Extract All Slides ===");
            TestExtractAllSlides(outputPath);
            Console.WriteLine();

            // Test 4: Render all slides to images
            Console.WriteLine("=== TEST 4: Render All Slides to Images ===");
            TestRenderAllSlides(outputPath);
            Console.WriteLine();

            // Test 5: End-to-end agent workflow
            Console.WriteLine("=== TEST 5: Complete Agent Workflow ===");
            TestAgentWorkflow(outputPath);
            Console.WriteLine();

            Console.WriteLine("==================================================");
            Console.WriteLine("✓ All convenience method tests completed!");
        }

        private static void TestFromFileFactory(string filePath)
        {
            try
            {
                // Simple factory method - no need to manually manage PresentationDocument
                using var builder = PresentationBuilder.FromFile(filePath);
                
                Console.WriteLine("✓ FromFile factory method works");
                Console.WriteLine("  No manual PresentationDocument.Open required");
                Console.WriteLine("  Auto-disposal with using statement");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ FromFile test failed: {ex.Message}");
            }
        }

        private static void TestGetSlideCount(string filePath)
        {
            try
            {
                using var builder = PresentationBuilder.FromFile(filePath);
                int count = builder.GetSlideCount();
                
                Console.WriteLine($"✓ GetSlideCount: {count} slides");
                Console.WriteLine("  Agent knows how many slides to process");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ GetSlideCount test failed: {ex.Message}");
            }
        }

        private static void TestExtractAllSlides(string filePath)
        {
            try
            {
                using var builder = PresentationBuilder.FromFile(filePath);
                var allSlides = builder.ExtractAllSlides();
                
                Console.WriteLine($"✓ ExtractAllSlides: Extracted {allSlides.Count} slides");
                
                foreach (var kvp in allSlides)
                {
                    var slideNum = kvp.Key;
                    var content = kvp.Value;
                    Console.WriteLine($"  Slide {slideNum}: {content.Shapes.Count} shapes, {content.Pictures.Count} pictures");
                }
                
                Console.WriteLine("  Single call extracts entire deck");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ ExtractAllSlides test failed: {ex.Message}");
            }
        }

        private static void TestRenderAllSlides(string filePath)
        {
            try
            {
                using var builder = PresentationBuilder.FromFile(filePath, isEditable: false);
                builder.Save(); // Ensure no pending changes
                
                // Close document so Syncfusion can read it
                builder.Dispose();
                
                // Now render (need to use helper directly since document is closed)
                var images = InternalUtilities.Syncfusion.SyncfusionHelperMethods.ExportPptToImages(filePath);
                
                Console.WriteLine($"✓ RenderAllSlides: Generated {images.Count} images");
                
                long totalSize = 0;
                for (int i = 0; i < images.Count; i++)
                {
                    var fileInfo = new FileInfo(images[i]);
                    totalSize += fileInfo.Length;
                    Console.WriteLine($"  Slide {i + 1}: {images[i]} ({fileInfo.Length / 1024} KB)");
                }
                
                Console.WriteLine($"  Total size: {totalSize / 1024} KB");
                Console.WriteLine("  Perfect for thumbnail galleries");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ RenderAllSlides test failed: {ex.Message}");
            }
        }

        private static void TestAgentWorkflow(string filePath)
        {
            try
            {
                Console.WriteLine("Simulating agent workflow:");
                Console.WriteLine("1. Extract all content");
                Console.WriteLine("2. Analyze and identify edits");
                Console.WriteLine("3. Make edits");
                Console.WriteLine("4. Generate new images\n");

                // Step 1: Extract
                Dictionary<int, SlideContentInfo> allContent;
                using (var builder = PresentationBuilder.FromFile(filePath))
                {
                    allContent = builder.ExtractAllSlides();
                    Console.WriteLine($"  ✓ Extracted {allContent.Count} slides");

                    // Step 2: Agent analyzes (simulated)
                    // In real scenario, agent would look at images + content to decide edits
                    
                    // Step 3: Make edits based on analysis
                    builder.EditSlideText(3, "Title 1", "Project Quantum");
                    builder.EditSlideText(3, "Subtitle 2", "AI-Powered Analysis\nJanuary 2026");
                    builder.EditSlideText(3, "Date Placeholder 6", "1/15/2026");
                    
                    builder.Save();
                    Console.WriteLine("  ✓ Applied edits to slide 3");
                }

                // Step 4: Generate new images (after closing document)
                var updatedImage = InternalUtilities.Syncfusion.SyncfusionHelperMethods.ExportSlideToImage(filePath, 3);
                Console.WriteLine($"  ✓ Generated updated image: {updatedImage}");
                Console.WriteLine("\n✓ Complete agent workflow successful!");
                Console.WriteLine("  Agent can: extract → analyze → edit → visualize");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Agent workflow test failed: {ex.Message}");
            }
        }
    }
}
