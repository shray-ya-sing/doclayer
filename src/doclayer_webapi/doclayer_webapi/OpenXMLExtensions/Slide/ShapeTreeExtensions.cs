using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Office2010.Drawing;

namespace OpenXMLExtensions
{
    public static class ShapeTreeExtensions
    {
        // gets all shapes in shape tree
        public static List<P.Shape> GetShapes(this P.ShapeTree shapeTree)
        {
            return shapeTree.Elements<P.Shape>().ToList();
        }
        public static void AddTextbox(this P.ShapeTree shapeTree, string text, int hpos, int vpos)
        {
            int height = 1;
            int width = 2;
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Textbox " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties() { TextBox = true});
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);

            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Rectangle);
            shapeProperties.AddChild(new D.NoFill());

            shape.AddChild(shapeProperties);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties(new D.NoAutoFit()) { RightToLeftColumns = true , Wrap = TextWrappingValues.Square});
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText(text);

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        /// <summary>
        /// Adds a rectangle to the slide
        /// </summary>
        /// <param name="shapeTree"></param>
        /// <param name="hpos"></param>
        /// <param name="vpos"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        public static void AddRectangle(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Rectangle " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);

            
                        
            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Rectangle);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

           P.ShapeStyle style =  new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);
            
            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true , Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddRectangle(this P.ShapeTree shapeTree, decimal hpos, decimal vpos, decimal height, decimal width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Rectangle " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Rectangle);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }
        public static void AddCircle(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Oval " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Ellipse);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddHarveyBall(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Partial Circle " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Pie);
            
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);

        }

        public static void AddLine(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.ConnectionShape shape = new P.ConnectionShape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualConnectionShapeProperties nonVisualConnectionShapeProperties = new P.NonVisualConnectionShapeProperties();

            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Straight Connector " + shapeNumber;
            nonVisualDrawingProperties.Id = id;

            nonVisualConnectionShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualConnectionShapeProperties.AddChild(new P.NonVisualConnectorShapeDrawingProperties());
            nonVisualConnectionShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualConnectionShapeProperties);

            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Line);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            shapeTree.AppendChild(shape);
        }


        public static void AddRightArrow(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Arrow: Right " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.RightArrow);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddLeftArrow(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Arrow: Left " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.LeftArrow);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddUpArrow(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Arrow: Up " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.UpArrow);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddDownArrow(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Arrow: Down " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.DownArrow);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddChevron(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Arrow: Chevron " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Chevron);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddPentagonArrow(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Arrow: Pentagon " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.HomePlate);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }
        public static void AddRoundedRectangle(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Rectangle: Rounded Corners " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.RoundRectangle);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddTriangle(this P.ShapeTree shapeTree, int hpos, int vpos, int height, int width)
        {
            P.Shape shape = new P.Shape();
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Isosceles Triangle " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            nonVisualShapeProperties.AddChild(nonVisualDrawingProperties);
            nonVisualShapeProperties.AddChild(new P.NonVisualShapeDrawingProperties());
            nonVisualShapeProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            shape.AddChild(nonVisualShapeProperties);



            // Default Shape Properties
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);

            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Triangle);
            shapeProperties.SetSchemeFill(1);
            shapeProperties.SetOutlineSchemeFill(1);

            shape.AddChild(shapeProperties);

            // Default shape Style

            P.ShapeStyle style = new P.ShapeStyle();
            style.SetDefaultReferences();

            shape.AddChild(style);

            // Default TextBody
            P.TextBody textBody = new P.TextBody();
            textBody.AddChild(new D.BodyProperties() { RightToLeftColumns = true, Anchor = D.TextAnchoringTypeValues.Center });
            textBody.AddChild(new D.ListStyle());

            D.Paragraph paragraph = new D.Paragraph();
            D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
            paragraph.AddChild(paragraphProperties);
            paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
            paragraph.SetText("placeholder text....");

            textBody.AddChild(paragraph);
            shape.AddChild(textBody);

            shapeTree.AppendChild(shape);
        }

        public static void AddTable(this P.ShapeTree shapeTree, int numRows, int numCols)
        {

            D.Table table = new D.Table();
            // Add Table Properties
            D.TableProperties tableProperties = new TableProperties() { FirstRow = true, BandRow = true};


            D.TableStyleId tableStyleId = new D.TableStyleId("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
            tableProperties.AddChild(tableStyleId);
            table.AddChild(tableProperties);

            // Add Table grid

            int colWidth = 1 * 914400;
            int rowHeight = 1 * 914400;

            D.TableGrid tableGrid = new D.TableGrid();

            List< D.GridColumn> cols = new List< D.GridColumn>();
            for (int i=0; i < numCols;  i++)
            {
                D.GridColumn col = new D.GridColumn(new D.ExtensionList()) { Width = colWidth };
                cols.Add(col);
            }

            foreach (D.GridColumn gridColumn in cols)
            {
                tableGrid.AppendChild(gridColumn);
            }

            table.AddChild(tableGrid);

            // Add Table rows

            List<D.TableRow> rows = new List<D.TableRow>();
            for (int i = 0; i < numRows; i++)
            {
                D.TableRow row = new D.TableRow() { Height = rowHeight};
                for (int k = 0; k < numCols; k++)
                {
                    D.TableCell tableCell = new D.TableCell();  
                    D.Paragraph paragraph = new D.Paragraph();
                    paragraph.AddChild(new EndParagraphRunProperties() { Language = "en-US" });
                    D.ListStyle listStyle = new D.ListStyle();
                    D.BodyProperties bodyProperties = new D.BodyProperties();
                    D.TextBody textBody = new D.TextBody();
                    textBody.AddChild(bodyProperties);
                    textBody.AddChild(listStyle);
                    textBody.AddChild(paragraph);
                    tableCell.AddChild(textBody);
                    tableCell.AddChild(new TableCellProperties());

                    row.AppendChild<TableCell>(tableCell);
                }
                    
                rows.Add(row);
            }

            foreach (D.TableRow row in rows)
            {
                table.AppendChild(row);
            }


            P.GraphicFrame graphicFrame = new P.GraphicFrame();

            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            P.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Table " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
            P.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties();
            nonVisualGraphicFrameDrawingProperties.AddChild(new D.GraphicFrameLocks() { NoGrouping = true });

            nonVisualGraphicFrameProperties.AddChild(nonVisualDrawingProperties);
            nonVisualGraphicFrameProperties.AddChild(nonVisualGraphicFrameDrawingProperties);
            nonVisualGraphicFrameProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());
            graphicFrame.AddChild(nonVisualGraphicFrameProperties);

            int hPos = 2;
            int vPos = 2;
            P.Transform transform = new P.Transform();
            transform.AddChild(new D.Offset() { X = hPos * 60000, Y = vPos * 60000 });
            transform.AddChild(new D.Extents() { Cx = colWidth*numCols, Cy = rowHeight * numRows });

            graphicFrame.AddChild(transform);

            graphicFrame.AddChild(new D.Graphic(new D.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));

            shapeTree.AppendChild(graphicFrame);

        }

        public static void AddPicture(this P.ShapeTree shapeTree, string relId, int hpos, int vpos) 
        {
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties
            
            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Picture " + shapeNumber;
            nonVisualDrawingProperties.Id = id;

            P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new();
            nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
            {
                NoChangeAspect = true
            });
            P.NonVisualPictureProperties nonVisualPictureProperties = new();
            nonVisualPictureProperties.AddChild(nonVisualDrawingProperties);
            nonVisualPictureProperties.AddChild(nonVisualPictureDrawingProperties);
            nonVisualPictureProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());

            // Blip Fill with the image data from the ImagePart

            P.BlipFill blipFill = new P.BlipFill();

            D.Stretch stretch = new D.Stretch();
            stretch.AddChild(new D.FillRectangle());
            blipFill.AddChild(stretch);

            BlipExtensionList blipExtensionList = new();
            blipExtensionList.AppendChild(new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" });
            UseLocalDpi useLocalDpi = new() { Val = false};
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtensionList.AppendChild(useLocalDpi);
            D.Blip blip = new() { Embed = relId };
            blip.AddChild(blipExtensionList);

            blipFill.AddChild(blip);


            // Add shape properties for dimensions and position

            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHeight(1);
            shapeProperties.SetWidth(1);
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Rectangle);

            // Create picture element and add to slide

            P.Picture picture = new();
            picture.AddChild(nonVisualPictureProperties);
            picture.AddChild(blipFill);
            picture.AddChild(shapeProperties);

            shapeTree.AppendChild(picture);

        }

        public static void AddPicture(this P.ShapeTree shapeTree, string relId, decimal hpos, decimal vpos)
        {
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties

            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Picture " + shapeNumber;
            nonVisualDrawingProperties.Id = id;

            P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new();
            nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
            {
                NoChangeAspect = true
            });
            P.NonVisualPictureProperties nonVisualPictureProperties = new();
            nonVisualPictureProperties.AddChild(nonVisualDrawingProperties);
            nonVisualPictureProperties.AddChild(nonVisualPictureDrawingProperties);
            nonVisualPictureProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());

            // Blip Fill with the image data from the ImagePart

            P.BlipFill blipFill = new P.BlipFill();

            D.Stretch stretch = new D.Stretch();
            stretch.AddChild(new D.FillRectangle());
            blipFill.AddChild(stretch);

            BlipExtensionList blipExtensionList = new();
            blipExtensionList.AppendChild(new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" });
            UseLocalDpi useLocalDpi = new() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtensionList.AppendChild(useLocalDpi);
            D.Blip blip = new() { Embed = relId };
            blip.AddChild(blipExtensionList);

            blipFill.AddChild(blip);


            // Add shape properties for dimensions and position

            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHeight(1);
            shapeProperties.SetWidth(1);
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Rectangle);

            // Create picture element and add to slide

            P.Picture picture = new();
            picture.AddChild(nonVisualPictureProperties);
            picture.AddChild(blipFill);
            picture.AddChild(shapeProperties);

            shapeTree.AppendChild(picture);

        }

        public static void AddPicture(this P.ShapeTree shapeTree, string relId, decimal height, decimal width, decimal hpos, decimal vpos)
        {
            //Specify the non visual drawing properties
            string shapeNumber = shapeTree.GetShapeNumber();
            UInt32Value id = shapeTree.GetShapeId();

            // Default Non Visual Shape Properties

            P.NonVisualDrawingProperties nonVisualDrawingProperties = new P.NonVisualDrawingProperties();
            nonVisualDrawingProperties.Name = "Picture " + shapeNumber;
            nonVisualDrawingProperties.Id = id;
           
            P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new();
            nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
            {
                NoChangeAspect = true
            });
            P.NonVisualPictureProperties nonVisualPictureProperties = new();
            nonVisualPictureProperties.AddChild(nonVisualDrawingProperties);
            nonVisualPictureProperties.AddChild(nonVisualPictureDrawingProperties);
            nonVisualPictureProperties.AddChild(new P.ApplicationNonVisualDrawingProperties());

            // Blip Fill with the image data from the ImagePart

            P.BlipFill blipFill = new P.BlipFill() { RotateWithShape = true };

            D.Stretch stretch = new D.Stretch();
            stretch.AddChild(new D.FillRectangle());
            blipFill.AddChild(stretch);

            BlipExtensionList blipExtensionList = new();
            blipExtensionList.AppendChild(new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" });
            UseLocalDpi useLocalDpi = new() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtensionList.AppendChild(useLocalDpi);
            D.Blip blip = new() { Embed = relId };
            blip.AddChild(blipExtensionList);

            blipFill.AddChild(blip);


            // Add shape properties for dimensions and position

            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.SetHeight(height);
            shapeProperties.SetWidth(width);
            shapeProperties.SetHorizontalPosition(hpos);
            shapeProperties.SetVerticalPosition(vpos);
            shapeProperties.SetPresetGeometry(D.ShapeTypeValues.Rectangle);

            // Create picture element and add to slide

            P.Picture picture = new();
            picture.AddChild(nonVisualPictureProperties);
            picture.AddChild(blipFill);
            picture.AddChild(shapeProperties);

            shapeTree.AppendChild(picture);

        }
        public static string GetShapeNumber(this P.ShapeTree shapeTree)
        {
            return (shapeTree.Elements<P.Shape>().Count() + shapeTree.Elements<P.GraphicFrame>().Count() + shapeTree.Elements<P.Picture>().Count() + 1).ToString();
        }

        public static UInt32Value GetShapeId(this P.ShapeTree shapeTree)
        {
            return (UInt32Value)UInt32.Parse((shapeTree.Elements<P.Shape>().Count() + shapeTree.Elements<P.NonVisualGroupShapeProperties>().Count() + shapeTree.Elements<P.GraphicFrame>().Count() + shapeTree.Elements<P.Picture>().Count() + 1).ToString());
        }

        /// <summary>
        /// Adds a page number to the bottom right corner of the slide
        /// </summary>
        /// <param name="shapeTree"></param>
        /// <param name="slideNumber"></param>
        public static void AddPageNumber(this P.ShapeTree shapeTree, string slideNumber) 
        {
            P.Shape shape2 = new P.Shape();

            P.NonVisualShapeProperties nonVisualShapeProperties2 = new P.NonVisualShapeProperties();

            P.NonVisualDrawingProperties nonVisualDrawingProperties3 = new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Slide Number Placeholder 2" };

            D.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new D.NonVisualDrawingPropertiesExtensionList();

            D.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new D.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties3.Append(nonVisualDrawingPropertiesExtensionList2);

            P.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new P.NonVisualShapeDrawingProperties();
            D.ShapeLocks shapeLocks2 = new D.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);
            P.ShapeProperties shapeProperties2 = new P.ShapeProperties();

            P.TextBody textBody2 = new P.TextBody();
            D.BodyProperties bodyProperties2 = new D.BodyProperties();
            D.ListStyle listStyle2 = new D.ListStyle();

            D.Paragraph paragraph2 = new D.Paragraph();

            D.Field field1 = new D.Field() { Id = "{A3570886-BC3E-4663-A581-D50C77742136}", Type = "slidenum" };

            D.RunProperties runProperties2 = new D.RunProperties() { Language = "en-US" };
            runProperties2.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            D.Text text2 = new D.Text();
            text2.Text = slideNumber;

            field1.Append(runProperties2);
            field1.Append(text2);
            D.EndParagraphRunProperties endParagraphRunProperties2 = new D.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph2.Append(field1);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);


        }

        /// <summary>
        /// Adds a title to the top of the slide
        /// </summary>
        /// <param name="shapeTree"></param>
        /// <param name="titleText"></param>
        public static void AddTitle(this P.ShapeTree shapeTree, string titleText)
        {
            P.Shape shape2 = new P.Shape();

            P.NonVisualShapeProperties nonVisualShapeProperties2 = new P.NonVisualShapeProperties();

            P.NonVisualDrawingProperties nonVisualDrawingProperties3 = new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Title 1" };

            D.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new D.NonVisualDrawingPropertiesExtensionList();

            D.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new D.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties3.Append(nonVisualDrawingPropertiesExtensionList2);

            P.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new P.NonVisualShapeDrawingProperties();
            D.ShapeLocks shapeLocks2 = new D.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);
            P.ShapeProperties shapeProperties2 = new P.ShapeProperties();

            P.TextBody textBody2 = new P.TextBody();
            D.BodyProperties bodyProperties2 = new D.BodyProperties();
            D.ListStyle listStyle2 = new D.ListStyle();

            D.Paragraph paragraph2 = new D.Paragraph();

            D.Field field1 = new D.Field() { Id = "{A3570886-BC3E-4663-A581-D50C77742136}", Type = "slidenum" };

            D.RunProperties runProperties2 = new D.RunProperties() { Language = "en-US" };

            D.Text text2 = new D.Text();
            text2.Text = titleText;

            field1.Append(runProperties2);
            field1.Append(text2);
            D.EndParagraphRunProperties endParagraphRunProperties2 = new D.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

            paragraph2.Append(field1);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);


        }


        /// <summary>
        /// Adds a footnote to the bottom of the slide
        /// </summary>
        /// <param name="shapeTree"></param>
        /// <param name="footNoteText"></param>
        public static void AddFootnote(this P.ShapeTree shapeTree, string footNoteText = "Source:")
        {
            P.Shape shape2 = new P.Shape();

            P.NonVisualShapeProperties nonVisualShapeProperties2 = new P.NonVisualShapeProperties();

            P.NonVisualDrawingProperties nonVisualDrawingProperties3 = new P.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Footer Placeholder 1" };

            D.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new D.NonVisualDrawingPropertiesExtensionList();

            D.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new D.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties3.Append(nonVisualDrawingPropertiesExtensionList2);

            P.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new P.NonVisualShapeDrawingProperties();
            D.ShapeLocks shapeLocks2 = new D.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);
            P.ShapeProperties shapeProperties2 = new P.ShapeProperties();

            P.TextBody textBody2 = new P.TextBody();
            D.BodyProperties bodyProperties2 = new D.BodyProperties();
            D.ListStyle listStyle2 = new D.ListStyle();

            D.Paragraph paragraph2 = new D.Paragraph();

            D.RunProperties runProperties2 = new D.RunProperties() { Language = "en-US" };

            D.Text text2 = new D.Text();
            text2.Text = footNoteText;

            D.EndParagraphRunProperties endParagraphRunProperties2 = new D.EndParagraphRunProperties() { Language = "en-US", Dirty = false };


            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);


        }
    }
}
