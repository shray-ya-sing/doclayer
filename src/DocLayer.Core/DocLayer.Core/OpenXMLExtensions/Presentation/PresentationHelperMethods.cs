
using System.IO.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using System.ComponentModel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Text;

namespace OpenXMLExtensions
{
    public static class PresentationHelperMethods
    {

        public static PresentationDocument CreatePresentation(string filepath)
            {
            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreatePresentationParts(presentationPart);

            //Dispose the presentation handle
            //presentationDoc.Dispose();

            return presentationDoc;
        }

        static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;


            slidePart1 = CreateSlidePart(presentationPart);
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
        }

        static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties(
                                        new D.Transform2D(
                                            new D.Offset() {X = 914400, Y = 914400 },
                                            new D.Extents() {Cx = 11 * 914400, Cy = 914400}
                                        )
                                    ),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                    new ColorMapOverride(new MasterColorMapping()));
            return slidePart1;
        }

        static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster slideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }))
              { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
              { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }


    [Description("Copies the powerpoint presentation saved at the user input source path and  saves at the user input target filepath")]
        public static void SavePresentationAs(string sourcePresentationFile, string targetPresentationFile)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(sourcePresentationFile, false))
            {
                PresentationDocument newPresentationDocument = presentationDocument.Clone(targetPresentationFile);
            }
        }

        [Description("Creates a presentation from a given theme saved at the template filepath string and saves at the user input filepath given by the the presentationFile parameter string")]
        public static void CreatePresentationFromTemplate(string template, string presentationFile)
        {
            PresentationDocument.CreateFromTemplate(template).Clone(presentationFile).Dispose();
        }

        public static void ApplyThemeToPresentation(string presentationFile, string themePresentation)
        {
            
            string info;
            using (PresentationDocument themeDocument = PresentationDocument.Open(themePresentation, true))
            {
                info = themeDocument.GetThemeInfo();
            }
            using (PresentationDocument presentationDoc = PresentationDocument.Open(presentationFile, true))
            {
                if (info != null)
                {
                    presentationDoc.PasteThemeInfo(info);
                }
            }

        }

        public static void ApplyThemeToPresentationDocument(PresentationDocument presentationDocument, PresentationDocument themeDocument)
        {

            // Get the presentation part of the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();

            // Get the existing slide master part.
            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.ElementAt(0);


            string relationshipId = presentationPart.GetIdOfPart(slideMasterPart);

            // Get the new slide master part.
            PresentationPart themeDocPresentationPart = themeDocument.PresentationPart ?? themeDocument.AddPresentationPart();
            SlideMasterPart newSlideMasterPart = themeDocPresentationPart.SlideMasterParts.ElementAt(0);
            if (presentationPart.ThemePart is not null)
            {
                // Remove the existing theme part.
                presentationPart.DeletePart(presentationPart.ThemePart);
            }

            // Remove the old slide master part.
            presentationPart.DeletePart(slideMasterPart);

            // Import the new slide master part, and reuse the old relationship ID.
            newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId);

            if (newSlideMasterPart.ThemePart is not null)
            {
                // Change to the new theme part.
                presentationPart.AddPart(newSlideMasterPart.ThemePart);
            }
            Dictionary<string, SlideLayoutPart> newSlideLayouts = new Dictionary<string, SlideLayoutPart>();

            foreach (var slideLayoutPart in newSlideMasterPart.SlideLayoutParts)
            {
                string? newLayoutType = GetSlideLayoutType(slideLayoutPart);

                if (newLayoutType is not null)
                {
                    newSlideLayouts.Add(newLayoutType, slideLayoutPart);
                }
            }

            string? layoutType = null;
            SlideLayoutPart? newLayoutPart = null;

            // Insert the code for the layout for this example.
            string defaultLayoutType = "Title and Content";
            // Remove the slide layout relationship on all slides. 
            foreach (var slidePart in presentationPart.SlideParts)
            {
                layoutType = null;

                if (slidePart.SlideLayoutPart is not null)
                {
                    // Determine the slide layout type for each slide.
                    layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart);

                    // Delete the old layout part.
                    slidePart.DeletePart(slidePart.SlideLayoutPart);
                }

                if (layoutType is not null && newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
                {
                    // Apply the new layout part.
                    slidePart.AddPart(newLayoutPart);
                }
                else
                {
                    newLayoutPart = newSlideLayouts[defaultLayoutType];

                    // Apply the new default layout part.
                    slidePart.AddPart(newLayoutPart);
                }
            }
        }

        // Get the slide layout type.
        static string? GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
        {
            CommonSlideData? slideData = slideLayoutPart.SlideLayout?.CommonSlideData;

            return slideData?.Name;
        }
        public static void AddThemeToThemePart(ThemePart themePart, D.Theme theme)
        {
            theme.Save(themePart);
        }

        public static P.Shape CreateShape(string name, UInt32 id, int x, int y, int height, int width)
        {        
            P.Shape shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties()
                    ),
                new P.ShapeProperties(
                    new D.Transform2D(
                        new D.Offset() { X = x, Y = y },
                        new D.Extents() { Cx = width, Cy = height}                        
                        ),
                    new D.PresetGeometry( new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle },
                    new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.Accent2 })
                    ),
                new P.TextBody(
                    new D.BodyProperties() { Anchor = D.TextAnchoringTypeValues.Center },
                    new D.ListStyle(),
                    new D.Paragraph(
                        new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center },
                        new D.EndParagraphRunProperties()
                        )
                    )
                );

            return shape;
        }

        public static Slide CreateNewSlide()
        {
            Slide slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties(),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                    new ColorMapOverride(new MasterColorMapping()));

            return slide;
        }

        public static void InsertSlideAtPosition(PresentationDocument presentationDocument, int position)
        {

            PresentationPart? presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;

            // Verify that the presentation is not empty.
            if (presentationPart is null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = CreateNewSlide();

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId? prevSlideId = null;

            OpenXmlElementList slideIds = slideIdList?.ChildElements ?? default;

            foreach (SlideId slideId in slideIds)
            {
                if (slideId.Id is not null && slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId is not null && prevSlideId.RelationshipId is not null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId!);
            }
            else
            {
                string? firstRelId = ((SlideId)slideIds[0]).RelationshipId;
                // If the first slide does not contain a relationship ID, throw an exception.
                if (firstRelId is null)
                {
                    throw new ArgumentNullException(nameof(firstRelId));
                }

                lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
            }

            // Use the same slide layout as that of the previous slide.
            if (lastSlidePart.SlideLayoutPart is not null)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList!.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

        }
        /// <summary>
        /// Get a list of the titles of all the slides in the presentation.
        /// </summary>
        /// <param name="presentationFile"></param>
        /// <returns></returns>

        static IList<string>? GetSlideTitles(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                return GetSlideTitlesFromPresentation(presentationDocument);
            }
        }

        
        /// <summary>
        /// Get a list of the titles of all the slides in the presentation.
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <returns></returns>
        static IList<string>? GetSlideTitlesFromPresentation(PresentationDocument presentationDocument)
        {
            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            if (presentationPart is not null && presentationPart.Presentation is not null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList is not null)
                {
                    List<string> titlesList = new List<string>();

                    // Get the title of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        if (slideId.RelationshipId is null)
                        {
                            continue;
                        }

                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);

                        // Get the slide title.
                        string title = GetSlideTitle(slidePart);

                        // An empty title can also be added.
                        titlesList.Add(title);
                    }

                    return titlesList;
                }

            }

            return null;
        }


        /// <summary>
        /// Gets the title string content of the slide.
        /// </summary>
        /// <param name="slidePart"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        static string GetSlideTitle(SlidePart slidePart)
        {
            if (slidePart is null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Declare a paragraph separator.
            string? paragraphSeparator = null;

            if (slidePart.Slide is not null)
            {
                // Find all the title shapes.
                var shapes = from shape in slidePart.Slide.Descendants<P.Shape>()
                             where IsTitleShape(shape)
                             select shape;

                StringBuilder paragraphText = new StringBuilder();

                foreach (var shape in shapes)
                {
                    var paragraphs = shape.TextBody?.Descendants<D.Paragraph>();
                    if (paragraphs is null)
                    {
                        continue;
                    }

                    // Get the text in each paragraph in this shape.
                    foreach (var paragraph in paragraphs)
                    {
                        // Add a line break.
                        paragraphText.Append(paragraphSeparator);

                        foreach (var text in paragraph.Descendants<D.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        paragraphSeparator = "\n";
                    }
                }

                return paragraphText.ToString();
            }

            return string.Empty;
        }


        /// <summary>
        /// Detects whether the shape is a title shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static bool IsTitleShape(P.Shape shape)
        {
            PlaceholderShape? placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();

            if (placeholderShape is not null && placeholderShape.Type is not null && placeholderShape.Type.HasValue)
            {
                return placeholderShape.Type == PlaceholderValues.Title || placeholderShape.Type == PlaceholderValues.CenteredTitle;
            }

            return false;
        }


        /// <summary>
        /// Get all the text content in a slide.
        /// </summary>
        /// <param name="presentationFile"></param>
        /// <param name="slideIndex"></param>
        /// <returns></returns>
        static string[]? GetAllTextInSlide(string presentationFile, int slideIndex)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation and the slide index
                // to the next GetAllTextInSlide method, and
                // then return the array of strings it returns. 
                return GetAllTextInSlideFromPresentation(presentationDocument, slideIndex);
            }
        }


        static string[]? GetAllTextInSlideFromPresentation(PresentationDocument presentationDocument, int slideIndex)
        {
            // Verify that the slide index is not out of range.
            if (slideIndex < 0)
            {
                throw new ArgumentOutOfRangeException("slideIndex");
            }

            // Get the presentation part of the presentation document.
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation part and presentation exist.
            if (presentationPart is not null && presentationPart.Presentation is not null)
            {
                // Get the Presentation object from the presentation part.
                Presentation presentation = presentationPart.Presentation;

                // Verify that the slide ID list exists.
                if (presentation.SlideIdList is not null)
                {
                    // Get the collection of slide IDs from the slide ID list.
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                    // If the slide ID is in range...
                    if (slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string? slidePartRelationshipId = ((SlideId)slideIds[slideIndex]).RelationshipId;

                        if (slidePartRelationshipId is null)
                        {
                            return null;
                        }

                        // Get the specified slide part from the relationship ID.
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                        // Pass the slide part to the next method, and
                        // then return the array of strings that method
                        // returns to the previous method.
                        return GetAllTextInSlideFromPart(slidePart);
                    }
                }
            }

            // Else, return null.
            return null;
        }
        static string[] GetAllTextInSlideFromPart(SlidePart slidePart)
        {
            // Verify that the slide part exists.
            if (slidePart is null)
            {
                throw new ArgumentNullException("slidePart");
            }

            // Create a new linked list of strings.
            LinkedList<string> texts = new LinkedList<string>();

            // If the slide exists...
            if (slidePart.Slide is not null)
            {
                // Iterate through all the paragraphs in the slide.
                foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in
                    slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    // Create a new string builder.                    
                    StringBuilder paragraphText = new StringBuilder();

                    // Iterate through the lines of the paragraph.
                    foreach (DocumentFormat.OpenXml.Drawing.Text text in
                        paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        // Append each line to the previous lines.
                        paragraphText.Append(text.Text);
                    }

                    if (paragraphText.Length > 0)
                    {
                        // Add each paragraph to the linked list.
                        texts.AddLast(paragraphText.ToString());
                    }
                }
            }

            // Return an array of strings.
            return texts.ToArray();
        }


        static void GetSlideIdAndText(out string sldText, string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart? part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part?.Presentation?.SlideIdList?.ChildElements ?? default;

                if (part is null || slideIds.Count == 0)
                {
                    sldText = "";
                    return;
                }

                string? relId = ((SlideId)slideIds[index]).RelationshipId;

                if (relId is null)
                {
                    sldText = "";
                    return;
                }

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                IEnumerable<D.Text> texts = slide.Slide.Descendants<D.Text>();
                foreach (D.Text text in texts)
                {
                    paragraphText.Append(text.Text);
                }
                sldText = paragraphText.ToString();
            }
        }

        // Move a slide to a different position in the slide order in the presentation.
        static void MoveSlide(string presentationFile, int from, int to)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                MoveSlideFromPresentation(presentationDocument, from, to);
            }
        }
        // Move a slide to a different position in the slide order in the presentation.
        static void MoveSlideFromPresentation(PresentationDocument presentationDocument, int from, int to)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Call the CountSlides method to get the number of slides in the presentation.
            int slidesCount = presentationDocument.GetSlideCount();

            // Verify that both from and to positions are within range and different from one another.
            if (from < 0 || from >= slidesCount)
            {
                throw new ArgumentOutOfRangeException("from");
            }

            if (to < 0 || from >= slidesCount || to == from)
            {
                throw new ArgumentOutOfRangeException("to");
            }

            // Get the presentation part from the presentation document.
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // The slide count is not zero, so the presentation must contain slides.            
            Presentation? presentation = presentationPart?.Presentation;

            if (presentation is null)
            {
                throw new ArgumentNullException(nameof(presentation));
            }

            SlideIdList? slideIdList = presentation.SlideIdList;

            if (slideIdList is null)
            {
                throw new ArgumentNullException(nameof(slideIdList));
            }

            // Get the slide ID of the source slide.
            SlideId? sourceSlide = slideIdList.ChildElements[from] as SlideId;

            if (sourceSlide is null)
            {
                throw new ArgumentNullException(nameof(sourceSlide));
            }

            SlideId? targetSlide = null;

            // Identify the position of the target slide after which to move the source slide.
            if (to == 0)
            {
                targetSlide = null;
            }
            else if (from < to)
            {
                targetSlide = slideIdList.ChildElements[to] as SlideId;
            }
            else
            {
                targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
            }

            // Remove the source slide from its current position.
            sourceSlide.Remove();

            // Insert the source slide at its new position after the target slide.
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            // Save the modified presentation.
            presentation.Save();
        }


        /// <summary>
        /// Detects whether the shape is a footer
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static bool IsFooter(P.Shape shape)
        {
            PlaceholderShape? placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();

            if (placeholderShape is not null && placeholderShape.Type is not null && placeholderShape.Type.HasValue)
            {
                return placeholderShape.Type == PlaceholderValues.Footer;
            }

            return false;
        }

        /// <summary>
        /// Detects whether the shape is a slide number
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static bool IsSlideNumber(P.Shape shape)
        {
            PlaceholderShape? placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();

            if (placeholderShape is not null && placeholderShape.Type is not null && placeholderShape.Type.HasValue)
            {
                return placeholderShape.Type == PlaceholderValues.SlideNumber;
            }

            return false;
        }

        /// <summary>
        /// Detects whether the shape is a header
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static bool IsHeader(P.Shape shape)
        {
            PlaceholderShape? placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();

            if (placeholderShape is not null && placeholderShape.Type is not null && placeholderShape.Type.HasValue)
            {
                return placeholderShape.Type == PlaceholderValues.Header;
            }

            return false;
        }

        /// <summary>
        /// Adds a slide with title layout (title and subtitle placeholders) to the presentation
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="position">Position to insert the slide (1-based index)</param>
        public static void AddTitleLayoutSlide(PresentationDocument presentationDocument, int position)
        {
            PresentationPart? presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;

            // Verify that the presentation is not empty.
            if (presentationPart is null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Create a new slide with title layout
            Slide slide = CreateTitleLayoutSlide();

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId? prevSlideId = null;

            OpenXmlElementList slideIds = slideIdList?.ChildElements ?? default;

            foreach (SlideId slideId in slideIds)
            {
                if (slideId.Id is not null && slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }
            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId is not null && prevSlideId.RelationshipId is not null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId!);
            }
            else
            {
                string? firstRelId = ((SlideId)slideIds[0]).RelationshipId;
                if (firstRelId is null)
                {
                    throw new ArgumentNullException(nameof(firstRelId));
                }

                lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
            }

            // Use the same slide layout as that of the previous slide.
            if (lastSlidePart.SlideLayoutPart is not null)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList!.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        /// <summary>
        /// Creates a new slide with title layout (centered title and subtitle placeholders)
        /// </summary>
        /// <returns>Slide object with title layout</returns>
        private static Slide CreateTitleLayoutSlide()
        {
            Slide slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup()),
                        // Title shape (centered)
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.CenteredTitle })),
                            new P.ShapeProperties(
                                new D.Transform2D(
                                    new D.Offset() { X = 914400, Y = 685800 },  // Centered, top third
                                    new D.Extents() { Cx = 10363200, Cy = 1371600 }  // Wide
                                )
                            ),
                            new P.TextBody(
                                new BodyProperties() { Anchor = D.TextAnchoringTypeValues.Center },
                                new ListStyle(),
                                new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))),
                        // Subtitle shape
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Subtitle 2" },
                                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U })),
                            new P.ShapeProperties(
                                new D.Transform2D(
                                    new D.Offset() { X = 914400, Y = 2057400 },  // Below title
                                    new D.Extents() { Cx = 10363200, Cy = 1371600 }
                                )
                            ),
                            new P.TextBody(
                                new BodyProperties() { Anchor = D.TextAnchoringTypeValues.Center },
                                new ListStyle(),
                                new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                new ColorMapOverride(new MasterColorMapping()));

            return slide;
        }

    }
}
