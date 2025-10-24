using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OpenXMLExtensions;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocLayer.Core
{
    /// <summary>
    /// Provides high-level methods for creating PowerPoint slides with pre-defined templates
    /// </summary>
    public class PresentationBuilder
    {
        private readonly PresentationDocument _presentationDoc;
        private PresentationPart _presentationPart;
        private uint _slideIdCounter = 256;

        public PresentationBuilder(PresentationDocument presentationDoc)
        {
            _presentationDoc = presentationDoc ?? throw new ArgumentNullException(nameof(presentationDoc));
            _presentationPart = _presentationDoc.PresentationPart ?? throw new InvalidOperationException("PresentationPart not found");
        }

        /// <summary>
        /// Creates a title slide with title, subtitle, and optional footnote
        /// </summary>
        /// <param name="title">Main title text</param>
        /// <param name="subtitle">Subtitle text (optional)</param>
        /// <param name="footnote">Footnote text (optional, defaults to "Source:")</param>
        /// <returns>The created SlidePart</returns>
        public void CreateTitleSlide(string title, string? subtitle = null, string? footnote = "Source:")
        {
            // Add a title layout slide at position 1
            PresentationHelperMethods.AddTitleLayoutSlide(_presentationDoc, 1);

            // Get the last slide (the one we just added)
            var slide = _presentationDoc.GetLastSlide();

            // Add title and subtitle text
            slide.AddTitle(title);
            if (!string.IsNullOrEmpty(subtitle))
            {
                slide.AddSubtitle(subtitle);
            }

            // Optionally add a footnote
            if (!string.IsNullOrEmpty(footnote))
            {
                slide.AddFootnote(footnote);
            }

            // Save the presentation
            _presentationDoc.Save();
        }

        /// <summary>
        /// Sets the presentation theme with custom fonts and colors
        /// </summary>
        /// <param name="fontName">Font typeface name (e.g., "Arial", "Calibri") - optional</param>
        /// <param name="accentColors">List of 4 hex color codes for accent colors (e.g., "4472C4") - optional</param>
        public void SetPresentationTheme(string? fontName = null, List<string>? accentColors = null)
        {
            // Get the theme from the first slide master
            var slideMasterPart = _presentationPart.SlideMasterParts.FirstOrDefault()
                ?? throw new InvalidOperationException("No slide master found in presentation");
            
            var themePart = slideMasterPart.ThemePart
                ?? throw new InvalidOperationException("No theme part found in slide master");
            
            var theme = themePart.Theme;

            // Set font if provided
            if (!string.IsNullOrEmpty(fontName))
            {
                if (fontName.Equals("Arial", StringComparison.OrdinalIgnoreCase))
                {
                    theme.SetFontSchemeArial();
                }
                else
                {
                    theme.SetCustomFontScheme(fontName);
                }
            }

            // Set accent colors if provided (must be exactly 4 colors)
            if (accentColors != null && accentColors.Count > 0)
            {
                if (accentColors.Count != 4)
                {
                    throw new ArgumentException("Must provide exactly 4 accent colors", nameof(accentColors));
                }
                theme.SetAccentColors(accentColors);
            }

            // Save changes
            theme.Save();
        }

        #region Helper Methods

        private P.NonVisualGroupShapeProperties CreateNonVisualGroupShapeProperties()
        {
            return new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            );
        }

        private P.Shape CreateTitleShape(string titleText)
        {
            P.Shape titleShape = new P.Shape();

            // Non-visual properties
            P.NonVisualShapeProperties nvShapeProps = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2, Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = PlaceholderValues.CenteredTitle })
            );
            titleShape.Append(nvShapeProps);

            // Shape properties - positioned at top center
            P.ShapeProperties shapeProps = new P.ShapeProperties();
            D.Transform2D transform = new D.Transform2D(
                new D.Offset { X = 914400L, Y = 685800L },  // Centered horizontally, near top
                new D.Extents { Cx = 10363200L, Cy = 1371600L }  // Wide and relatively tall for title
            );
            shapeProps.Append(transform);
            titleShape.Append(shapeProps);

            // Text body
            P.TextBody textBody = new P.TextBody(
                new D.BodyProperties { Anchor = D.TextAnchoringTypeValues.Center },
                new D.ListStyle(),
                CreateTitleParagraph(titleText)
            );
            titleShape.Append(textBody);

            return titleShape;
        }

        private D.Paragraph CreateTitleParagraph(string text)
        {
            D.Paragraph para = new D.Paragraph();
            
            D.ParagraphProperties paraProps = new D.ParagraphProperties
            {
                Alignment = D.TextAlignmentTypeValues.Center
            };
            para.Append(paraProps);

            D.Run run = new D.Run();
            D.RunProperties runProps = new D.RunProperties
            {
                Language = "en-US",
                FontSize = 4400,  // 44pt
                Bold = true,
                Dirty = false
            };
            D.SolidFill solidFill = new D.SolidFill(
                new D.SchemeColor { Val = D.SchemeColorValues.Text1 }
            );
            runProps.Append(solidFill);
            run.Append(runProps);
            run.Append(new D.Text(text));
            para.Append(run);

            para.Append(new D.EndParagraphRunProperties { Language = "en-US" });

            return para;
        }

        private P.Shape CreateSubtitleShape(string? subtitleText)
        {
            P.Shape subtitleShape = new P.Shape();

            // Non-visual properties
            P.NonVisualShapeProperties nvShapeProps = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 3, Name = "Subtitle" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = PlaceholderValues.SubTitle, Index = 1 })
            );
            subtitleShape.Append(nvShapeProps);

            // Shape properties - positioned below title
            P.ShapeProperties shapeProps = new P.ShapeProperties();
            D.Transform2D transform = new D.Transform2D(
                new D.Offset { X = 914400L, Y = 2057400L },  // Below title
                new D.Extents { Cx = 10363200L, Cy = 1371600L }
            );
            shapeProps.Append(transform);
            subtitleShape.Append(shapeProps);

            // Text body
            D.Paragraph para = new D.Paragraph();
            D.ParagraphProperties paraProps = new D.ParagraphProperties
            {
                Alignment = D.TextAlignmentTypeValues.Center
            };
            para.Append(paraProps);

            if (!string.IsNullOrEmpty(subtitleText))
            {
                D.Run run = new D.Run();
                D.RunProperties runProps = new D.RunProperties
                {
                    Language = "en-US",
                    FontSize = 2400,  // 24pt
                    Dirty = false
                };
                D.SolidFill solidFill = new D.SolidFill(
                    new D.SchemeColor { Val = D.SchemeColorValues.Text1 }
                );
                runProps.Append(solidFill);
                run.Append(runProps);
                run.Append(new D.Text(subtitleText));
                para.Append(run);
            }

            para.Append(new D.EndParagraphRunProperties { Language = "en-US" });

            P.TextBody textBody = new P.TextBody(
                new D.BodyProperties { Anchor = D.TextAnchoringTypeValues.Center },
                new D.ListStyle(),
                para
            );
            subtitleShape.Append(textBody);

            return subtitleShape;
        }

        private void AddFootnoteToShapeTree(P.ShapeTree shapeTree, string footnoteText)
        {
            P.Shape footnoteShape = new P.Shape();

            // Non-visual properties
            uint shapeId = (uint)(shapeTree.Elements<P.Shape>().Count() + shapeTree.Elements<P.Picture>().Count() + 2);
            P.NonVisualShapeProperties nvShapeProps = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Footnote {shapeId}" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            );
            footnoteShape.Append(nvShapeProps);

            // Shape properties - positioned at bottom left
            P.ShapeProperties shapeProps = new P.ShapeProperties();
            D.Transform2D transform = new D.Transform2D(
                new D.Offset { X = 457200L, Y = 6400000L },  // Bottom left
                new D.Extents { Cx = 3048000L, Cy = 457200L }  // Small size for footnote
            );
            shapeProps.Append(transform);
            shapeProps.Append(new D.PresetGeometry { Preset = D.ShapeTypeValues.Rectangle });
            shapeProps.Append(new D.NoFill());
            footnoteShape.Append(shapeProps);

            // Text body
            D.Paragraph para = new D.Paragraph();
            D.ParagraphProperties paraProps = new D.ParagraphProperties
            {
                Alignment = D.TextAlignmentTypeValues.Left
            };
            para.Append(paraProps);

            D.Run run = new D.Run();
            D.RunProperties runProps = new D.RunProperties
            {
                Language = "en-US",
                FontSize = 1000,  // 10pt - small for footnote
                Dirty = false
            };
            D.SolidFill solidFill = new D.SolidFill(
                new D.SchemeColor { Val = D.SchemeColorValues.Text1 }
            );
            runProps.Append(solidFill);
            run.Append(runProps);
            run.Append(new D.Text(footnoteText));
            para.Append(run);
            para.Append(new D.EndParagraphRunProperties { Language = "en-US" });

            P.TextBody textBody = new P.TextBody(
                new D.BodyProperties { Wrap = D.TextWrappingValues.Square },
                new D.ListStyle(),
                para
            );
            footnoteShape.Append(textBody);

            shapeTree.Append(footnoteShape);
        }

        private void AddSlideToPresentation(SlidePart slidePart)
        {
            P.Presentation presentation = _presentationPart.Presentation;
            
            if (presentation.SlideIdList == null)
            {
                presentation.SlideIdList = new SlideIdList();
            }

            SlideId slideId = new SlideId
            {
                Id = _slideIdCounter++,
                RelationshipId = _presentationPart.GetIdOfPart(slidePart)
            };

            presentation.SlideIdList.Append(slideId);
        }

        #endregion
    }
}
