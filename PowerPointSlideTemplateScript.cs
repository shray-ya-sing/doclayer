using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using OpenXMLExtensions;

namespace PowerPointTemplateScript
{
    /// <summary>
    /// PowerPoint Slide Template Script using OpenXMLExtensions
    /// Demonstrates how to create a comprehensive slide with various elements
    /// </summary>
    public class SlideTemplateBuilder
    {
        private PresentationDocument _presentationDocument;
        private string _filePath;

        public SlideTemplateBuilder(string filePath)
        {
            _filePath = filePath;
        }

        /// <summary>
        /// Creates a new PowerPoint presentation with a template slide
        /// </summary>
        public void CreatePresentationWithTemplate()
        {
            // Create a new presentation document
            _presentationDocument = PresentationDocument.Create(_filePath, PresentationDocumentType.Presentation);
            
            // Add presentation part
            PresentationPart presentationPart = _presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            // Set slide size to widescreen (16:9)
            presentationPart.Presentation.SetSlideSizeWidescreen();

            // Create slide master and layout parts (simplified)
            CreateSlideLayout(presentationPart);

            // Add the first slide
            AddTemplateSlide();

            // Save and close
            _presentationDocument.Save();
            _presentationDocument.Close();
        }

        /// <summary>
        /// Creates a basic slide layout
        /// </summary>
        private void CreateSlideLayout(PresentationPart presentationPart)
        {
            // Create slide master part
            SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup())
                )),
                new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping())
            );

            // Create slide layout part
            SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup())
                )),
                new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping())
            );

            // Link slide master to presentation
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                new SlideMasterId() { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }
            );
        }

        /// <summary>
        /// Adds a comprehensive template slide with various elements
        /// </summary>
        private void AddTemplateSlide()
        {
            // Create slide part
            SlidePart slidePart = _presentationDocument.PresentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = CreateTemplateSlide();

            // Add slide to presentation
            SlideIdList slideIdList = _presentationDocument.PresentationPart.Presentation.SlideIdList ?? new SlideIdList();
            uint slideId = (uint)(slideIdList.Count() + 256);
            slideIdList.Append(new SlideId() { Id = slideId, RelationshipId = _presentationDocument.PresentationPart.GetIdOfPart(slidePart) });
            _presentationDocument.PresentationPart.Presentation.SlideIdList = slideIdList;
        }

        /// <summary>
        /// Creates a template slide with comprehensive content using OpenXMLExtensions
        /// </summary>
        private Slide CreateTemplateSlide()
        {
            Slide slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                            new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                        ),
                        new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup())
                    )
                ),
                new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping())
            );

            // 1. Add Title
            slide.AddTitle("Template Slide - Company Presentation");

            // 2. Add various shapes using extensions
            slide.AddTextbox("Key Bullet Points:\n• Market Analysis\n• Financial Overview\n• Strategic Initiatives\n• Next Steps");
            
            slide.AddRectangle(); // Default rectangle
            
            // 3. Add a table for data presentation
            slide.AddTable(3, 4); // 3 rows, 4 columns

            // 4. Add footer and page number
            slide.AddFootnote("Source: Company Internal Data");
            slide.AddPageNumber("1");

            return slide;
        }

        /// <summary>
        /// Creates an advanced template with custom positioning and styling
        /// </summary>
        public void CreateAdvancedTemplate()
        {
            _presentationDocument = PresentationDocument.Create(_filePath.Replace(".pptx", "_advanced.pptx"), PresentationDocumentType.Presentation);
            
            PresentationPart presentationPart = _presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.SetSlideSizeWidescreen();

            CreateSlideLayout(presentationPart);

            // Create slide part
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = CreateAdvancedSlide();

            // Add slide to presentation
            SlideIdList slideIdList = new SlideIdList();
            slideIdList.Append(new SlideId() { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });
            presentationPart.Presentation.SlideIdList = slideIdList;

            _presentationDocument.Save();
            _presentationDocument.Close();
        }

        /// <summary>
        /// Creates an advanced slide with custom shapes, positioning, and styling
        /// </summary>
        private Slide CreateAdvancedSlide()
        {
            Slide slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                            new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                        ),
                        new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup())
                    )
                ),
                new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping())
            );

            // Get the shape tree for advanced operations
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // 1. Add title
            slide.AddTitle("Advanced Template - Dashboard View");

            // 2. Add various geometric shapes with custom positioning
            shapeTree.AddRectangle(1, 2, 2, 3);        // x=1, y=2, height=2, width=3
            shapeTree.AddCircle(5, 2, 2, 2);           // Circle at x=5, y=2
            shapeTree.AddTriangle(9, 2, 2, 2);         // Triangle
            shapeTree.AddRoundedRectangle(1, 5, 2, 3); // Rounded rectangle

            // 3. Add arrows for process flow
            shapeTree.AddRightArrow(5, 5, 1, 2);
            shapeTree.AddDownArrow(7, 3, 2, 1);
            shapeTree.AddChevron(9, 5, 1, 2);

            // 4. Add text boxes with specific positioning
            shapeTree.AddTextbox("Process Step 1:\nInitiation", 1, 1);
            shapeTree.AddTextbox("Process Step 2:\nAnalysis", 5, 1);
            shapeTree.AddTextbox("Process Step 3:\nImplementation", 9, 1);

            // 5. Add a data table
            slide.AddTable(4, 3); // 4 rows, 3 columns for metrics

            // 6. Add lines for visual separation
            shapeTree.AddLine(0, 4, 0, 12); // Vertical line
            shapeTree.AddLine(4, 0, 12, 0); // Horizontal line

            // 7. Add footer elements
            slide.AddFootnote("Confidential - Internal Use Only");
            slide.AddPageNumber("Dashboard - Page 1");

            return slide;
        }
    }

    /// <summary>
    /// Example usage and entry point
    /// </summary>
    public class Program
    {
        public static void Main(string[] args)
        {
            string outputPath = "C:\\Users\\shrey\\projects\\doclayer\\";
            
            // Create basic template
            var basicTemplate = new SlideTemplateBuilder($"{outputPath}BasicTemplate.pptx");
            basicTemplate.CreatePresentationWithTemplate();
            Console.WriteLine("Basic template created: BasicTemplate.pptx");

            // Create advanced template
            var advancedTemplate = new SlideTemplateBuilder($"{outputPath}AdvancedTemplate.pptx");
            advancedTemplate.CreateAdvancedTemplate();
            Console.WriteLine("Advanced template created: AdvancedTemplate_advanced.pptx");

            Console.WriteLine("\nTemplates created successfully!");
            Console.WriteLine("\nAvailable OpenXMLExtensions methods used:");
            Console.WriteLine("- Slide Extensions: AddTitle(), AddTextbox(), AddTable(), AddFootnote(), AddPageNumber()");
            Console.WriteLine("- ShapeTree Extensions: AddRectangle(), AddCircle(), AddTriangle(), AddRoundedRectangle()");
            Console.WriteLine("- ShapeTree Extensions: AddRightArrow(), AddDownArrow(), AddChevron(), AddLine()");
            Console.WriteLine("- Presentation Extensions: SetSlideSizeWidescreen()");
        }
    }

    /// <summary>
    /// Utility class for common template operations
    /// </summary>
    public static class TemplateUtilities
    {
        /// <summary>
        /// Creates a slide with company branding template
        /// </summary>
        public static Slide CreateBrandedSlide(string title, string subtitle = "")
        {
            Slide slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                            new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                        ),
                        new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup())
                    )
                ),
                new ColorMapOverride(new DocumentFormat.OpenXml.Drawing.MasterColorMapping())
            );

            // Add branding elements
            slide.AddTitle(title);
            
            if (!string.IsNullOrEmpty(subtitle))
            {
                slide.AddTextbox(subtitle);
            }

            // Add standard footer
            slide.AddFootnote("© 2024 Company Name - Confidential");
            
            return slide;
        }

        /// <summary>
        /// Creates a data visualization slide template
        /// </summary>
        public static Slide CreateDataVisualizationSlide(string title, int tableRows = 5, int tableCols = 4)
        {
            Slide slide = CreateBrandedSlide(title);
            
            // Add data table
            slide.AddTable(tableRows, tableCols);
            
            // Add chart placeholder rectangle
            slide.AddRectangle();
            
            // Add legend area
            var shapeTree = slide.CommonSlideData.ShapeTree;
            shapeTree.AddTextbox("Chart Legend:\n• Series 1\n• Series 2\n• Series 3", 8, 2);
            
            return slide;
        }

        /// <summary>
        /// Creates a process flow slide template
        /// </summary>
        public static Slide CreateProcessFlowSlide(string title)
        {
            Slide slide = CreateBrandedSlide(title);
            var shapeTree = slide.CommonSlideData.ShapeTree;

            // Create process flow with arrows and shapes
            shapeTree.AddRoundedRectangle(1, 3, 1, 2);
            shapeTree.AddRightArrow(3, 3, 1, 1);
            shapeTree.AddRoundedRectangle(4, 3, 1, 2);
            shapeTree.AddRightArrow(6, 3, 1, 1);
            shapeTree.AddRoundedRectangle(7, 3, 1, 2);

            // Add process labels
            shapeTree.AddTextbox("Step 1", 1, 2);
            shapeTree.AddTextbox("Step 2", 4, 2);
            shapeTree.AddTextbox("Step 3", 7, 2);

            return slide;
        }
    }
}