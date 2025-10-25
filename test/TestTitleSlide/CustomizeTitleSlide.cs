using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocLayer.Core.Examples
{
    public static class CustomizeTitleSlide
    {
        public static void Run()
        {
            string sourcePath = @"C:\Users\shrey\OneDrive\Desktop\docs\2024.10.27 Project Core - Valuation Analysis_v23.pptx";
            string outputPath = @"C:\Users\shrey\OneDrive\Desktop\docs\Project Genesis - Pitch Materials.pptx";

            Console.WriteLine($"Loading presentation from: {sourcePath}");
            
            // Open the presentation
            using (PresentationDocument presentationDocument = PresentationDocument.Open(sourcePath, false))
            {
                // Create a copy for modification
                File.Copy(sourcePath, outputPath, true);
            }

            // Open the copy for editing
            using (PresentationDocument presentationDocument = PresentationDocument.Open(outputPath, true))
            {
                PresentationPart? presentationPart = presentationDocument.PresentationPart;
                if (presentationPart?.Presentation?.SlideIdList == null)
                {
                    Console.WriteLine("Error: Could not find slide list in presentation");
                    return;
                }

                // Get the third slide (index 2)
                var slideIds = presentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
                if (slideIds.Count < 3)
                {
                    Console.WriteLine("Error: Presentation does not have 3 slides");
                    return;
                }

                var slideId = slideIds[2]; // Third slide (0-indexed)
                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                
                Console.WriteLine("Customizing slide 3...");
                CustomizeSlide(slidePart);
                
                presentationPart.Presentation.Save();
                Console.WriteLine($"✓ Presentation saved to: {outputPath}");
            }
        }

        private static void CustomizeSlide(SlidePart slidePart)
        {
            var slide = slidePart.Slide;
            var shapeTree = slide.CommonSlideData?.ShapeTree;
            
            if (shapeTree == null)
            {
                Console.WriteLine("Error: Could not find shape tree in slide");
                return;
            }

            // Find and update the title shape
            var titleShape = shapeTree.Elements<Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?.Type?.Value == PlaceholderValues.CenteredTitle);
            
            if (titleShape != null)
            {
                UpdateTitle(titleShape, "Project Genesis");
                Console.WriteLine("  - Updated title");
            }

            // Find and update the subtitle shape
            var subtitleShape = shapeTree.Elements<Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?.Type?.Value == PlaceholderValues.SubTitle);
            
            if (subtitleShape != null)
            {
                UpdateSubtitle(subtitleShape, "Pitch Materials", "October 2025");
                Console.WriteLine("  - Updated subtitle");
            }

            // Find and update the red rectangle with "Preliminary and Illustrative" text
            var rectangleShape = shapeTree.Elements<Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Rectangle 5");
            
            if (rectangleShape != null)
            {
                // Keep the rectangle as is - it already says "Preliminary and Illustrative"
                Console.WriteLine("  - Rectangle shape found (no changes needed)");
            }
        }

        private static void UpdateTitle(Shape titleShape, string newTitle)
        {
            var textBody = titleShape.TextBody;
            if (textBody == null) return;

            // Clear existing paragraphs and add new one
            textBody.RemoveAllChildren<A.Paragraph>();
            
            var paragraph = new A.Paragraph();
            paragraph.ParagraphProperties = new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Left };
            
            var run = new A.Run();
            var runProperties = new A.RunProperties 
            { 
                Language = "en-US", 
                FontSize = 3000, 
                Dirty = false 
            };
            runProperties.Append(new A.LatinFont 
            { 
                Typeface = "Arial", 
                Panose = "020B0604020202020204", 
                PitchFamily = 34, 
                CharacterSet = 0 
            });
            runProperties.Append(new A.ComplexScriptFont 
            { 
                Typeface = "Arial", 
                Panose = "020B0604020202020204", 
                PitchFamily = 34, 
                CharacterSet = 0 
            });
            
            run.Append(runProperties);
            run.Append(new A.Text { Text = newTitle });
            
            paragraph.Append(run);
            textBody.Append(paragraph);
        }

        private static void UpdateSubtitle(Shape subtitleShape, string line1, string line2)
        {
            var textBody = subtitleShape.TextBody;
            if (textBody == null) return;

            // Clear existing paragraphs
            textBody.RemoveAllChildren<A.Paragraph>();
            
            // Add first line (Pitch Materials)
            var paragraph1 = new A.Paragraph();
            paragraph1.ParagraphProperties = new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Left };
            
            var run1 = new A.Run();
            var runProperties1 = new A.RunProperties 
            { 
                Language = "en-US", 
                FontSize = 1600, 
                Dirty = false 
            };
            runProperties1.Append(new A.LatinFont 
            { 
                Typeface = "Arial", 
                Panose = "020B0604020202020204", 
                PitchFamily = 34, 
                CharacterSet = 0 
            });
            runProperties1.Append(new A.ComplexScriptFont 
            { 
                Typeface = "Arial", 
                Panose = "020B0604020202020204", 
                PitchFamily = 34, 
                CharacterSet = 0 
            });
            
            run1.Append(runProperties1);
            run1.Append(new A.Text { Text = line1 });
            paragraph1.Append(run1);
            
            // Add second line (October 2025) - italic
            var paragraph2 = new A.Paragraph();
            paragraph2.ParagraphProperties = new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Left };
            
            var run2 = new A.Run();
            var runProperties2 = new A.RunProperties 
            { 
                Language = "en-US", 
                FontSize = 1600, 
                Italic = true, 
                Dirty = false 
            };
            runProperties2.Append(new A.LatinFont 
            { 
                Typeface = "Arial", 
                Panose = "020B0604020202020204", 
                PitchFamily = 34, 
                CharacterSet = 0 
            });
            runProperties2.Append(new A.ComplexScriptFont 
            { 
                Typeface = "Arial", 
                Panose = "020B0604020202020204", 
                PitchFamily = 34, 
                CharacterSet = 0 
            });
            
            run2.Append(runProperties2);
            run2.Append(new A.Text { Text = line2 });
            paragraph2.Append(run2);
            
            textBody.Append(paragraph1);
            textBody.Append(paragraph2);
        }
    }
}
