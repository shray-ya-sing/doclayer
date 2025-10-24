using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace OpenXMLExtensions
{
    public static class ShapeExtensions
    {
        public static string GetName(this P.Shape shape)
        {
         P.NonVisualShapeProperties nvShapeProps = shape.GetFirstChild<P.NonVisualShapeProperties>();
         P.NonVisualDrawingProperties nvDrawingProps = nvShapeProps.GetFirstChild<P.NonVisualDrawingProperties>();

            return nvDrawingProps.Name.ToString();
        }

        public static UInt32Value GetId(this P.Shape shape)
        {
            P.NonVisualShapeProperties nvShapeProps = shape.GetFirstChild<P.NonVisualShapeProperties>();
            P.NonVisualDrawingProperties nvShapeDrawingProps = nvShapeProps.GetFirstChild<P.NonVisualDrawingProperties>();

            return nvShapeDrawingProps.Id;
        }

        public static string GetElementClass(this P.Shape shape)
        {
            string name = shape.GetName();
            int end = name.LastIndexOf(' ');
            string elementClass = name.Substring(0, end);
            return elementClass;           
        }

        /// <summary>
        /// TryGet the paragraph element in shape containing the text. Returns true if found, false otherwise.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="text"></param>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public static bool TryGetParagraphContaining(this P.Shape shape, string text, out D.Paragraph paragraph) 
        {
            if (shape.TextBody != null)
            {
                foreach (D.Paragraph para in shape.TextBody.Elements<D.Paragraph>())
                {
                    if (para.GetText().Contains(text, StringComparison.OrdinalIgnoreCase));
                    {
                        paragraph = para;
                        return true;
                    }
                }
            }
            paragraph = null;
            return false;
        }
        public static string GetText(this P.Shape shape)
        {
            string text = "";
            if (shape.TextBody != null)
            {
                foreach (D.Paragraph paragraph in shape.TextBody.Elements<D.Paragraph>())
                {
                    foreach (D.Run run in paragraph.Elements<D.Run>())
                    {
                        if (run.GetFirstChild<D.Text>() is not null)
                        {
                            text+= run.GetFirstChild<D.Text>().InnerText.ToString();
                                
                        }

                    }
                }
            }

            return text;
        }

        public static void SetText(this P.Shape shape, string text)
        {
            if (shape.GetFirstChild<P.TextBody>() is not null)
            {
                if (shape.GetFirstChild<P.TextBody>().Elements<D.Paragraph>() is not null)
                {
                    for (int i = shape.GetFirstChild<P.TextBody>().Elements<D.Paragraph>().Count() - 1; i > 0; i--)
                    {
                        // Remove all paras after first
                        shape.GetFirstChild<P.TextBody>().Elements<D.Paragraph>().ToList()[i].Remove();
                    }

                    D.Paragraph para = shape.GetFirstChild<P.TextBody>().Elements<D.Paragraph>().First();
                    para.SetText(text); 
                }               
                
                else
                {
                    D.Paragraph paragraph = new D.Paragraph();
                    D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
                    paragraph.AddChild(paragraphProperties);
                    paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
                    paragraph.SetText(text);

                    shape.GetFirstChild<P.TextBody>().AddChild(paragraph);
                }
            }
            else
            {
                // Default TextBody
                P.TextBody textBody = new P.TextBody();
                textBody.AddChild(new D.BodyProperties(new D.NoAutoFit()) { RightToLeftColumns = true, Wrap = TextWrappingValues.Square });
                textBody.AddChild(new D.ListStyle());

                D.Paragraph paragraph = new D.Paragraph();
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
                paragraph.AddChild(paragraphProperties);
                paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
                paragraph.SetText(text);

                textBody.AddChild(paragraph);
                shape.AddChild(textBody);
            }
            
        }

        public static void AddText(this P.Shape shape, string text)
        {
            if (shape.GetFirstChild<P.TextBody>() is not null)
            {
                D.Paragraph paragraph = new D.Paragraph();
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
                paragraph.AddChild(paragraphProperties);
                paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
                paragraph.SetText(text);

                shape.GetFirstChild<P.TextBody>().AppendChild(paragraph);

            }
            else
            {
                // Default TextBody
                P.TextBody textBody = new P.TextBody();
                textBody.AddChild(new D.BodyProperties(new D.NoAutoFit()) { RightToLeftColumns = true, Wrap = TextWrappingValues.Square });
                textBody.AddChild(new D.ListStyle());

                D.Paragraph paragraph = new D.Paragraph();
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Left };
                paragraph.AddChild(paragraphProperties);
                paragraph.AddChild(new D.EndParagraphRunProperties() { Language = "en-US" });
                paragraph.SetText(text);

                textBody.AddChild(paragraph);
                shape.AddChild(textBody);
            }

        }

        /// <summary>
        /// Sets fill by hex code
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="rgbColorHex"></param>
        public static void SetHexFill(this P.Shape shape, string rgbColorHex)
        {
            P.ShapeProperties props = shape.ShapeProperties;
            if (props.GetFirstChild<D.SolidFill>() is not null)
            {
                D.SolidFill solidFill = props.GetFirstChild<D.SolidFill>();
                D.SolidFill newSolidFill = new D.SolidFill(
                   new RgbColorModelHex() { Val = rgbColorHex }
                   );

                props.ReplaceChild<SolidFill>(newSolidFill, solidFill);


            }
            else
            {
               D.SolidFill solidFill = new D.SolidFill(
                   new RgbColorModelHex() { Val = rgbColorHex}
                   );
                props.AddChild(solidFill);
            }
            
        }

        /// <summary>
        /// Sets fill by scheme color
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="accentNum"></param>

        public static void SetSchemeFill(this P.Shape shape, int accentNum)
        {
            D.SchemeColorValues schemeColorVal;

            switch (accentNum)
            {
                case 1:
                    schemeColorVal = D.SchemeColorValues.Accent1;
                    break;

                case 2:
                    schemeColorVal = D.SchemeColorValues.Accent2;
                    break;

                case 3:
                    schemeColorVal = D.SchemeColorValues.Accent3;
                    break;

                case 4:
                    schemeColorVal = D.SchemeColorValues.Accent4;
                    break;

            }


            P.ShapeProperties props = shape.ShapeProperties;
            if (props.GetFirstChild<D.SolidFill>() is not null)
            {
                D.SolidFill solidFill = props.GetFirstChild<D.SolidFill>();
                D.SolidFill newSolidFill = new D.SolidFill(
                   new SchemeColor() { Val =  schemeColorVal }
                   );

                props.ReplaceChild<SolidFill>(newSolidFill, solidFill);
            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(
                   new SchemeColor() { Val = schemeColorVal }
                   );
                props.AddChild(solidFill);
            }

        }

        /// <summary>
        /// Sets Outline color by Hexcode
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="rgbColorHex"></param>
        public static void SetOutlineHexFill(this P.Shape shape, string rgbColorHex)
        {
            P.ShapeProperties props = shape.ShapeProperties;
            if (props.GetFirstChild<D.Outline>() is not null)
            {
                D.Outline outline = props.GetFirstChild<D.Outline>();
                D.SolidFill solidFill = outline.GetFirstChild<D.SolidFill>();
                D.SolidFill newSolidFill = new D.SolidFill(
                   new RgbColorModelHex() { Val = rgbColorHex }
                   );

                outline.ReplaceChild<SolidFill>(newSolidFill, solidFill);
            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(
                    new RgbColorModelHex() { Val = rgbColorHex }
                    );
                D.Outline outline = new D.Outline(solidFill);
                props.AddChild(outline);
                
            }

        }

        /// <summary>
        /// Sets outline fill by scheme color
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="accentNum"></param>
        public static void SetOutlineSchemeFill(this P.Shape shape, int accentNum)
        {
            //IMP: Method assumes shape props not null

            D.SchemeColorValues schemeColorVal;

            switch (accentNum)
            {
                case 1:
                    schemeColorVal = D.SchemeColorValues.Accent1;
                    break;

                case 2:
                    schemeColorVal = D.SchemeColorValues.Accent2;
                    break;

                case 3:
                    schemeColorVal = D.SchemeColorValues.Accent3;
                    break;

                case 4:
                    schemeColorVal = D.SchemeColorValues.Accent4;
                    break;

            }


            P.ShapeProperties props = shape.ShapeProperties;
            if (props.GetFirstChild<D.Outline>() != null)
            {
                D.Outline outline = props.GetFirstChild<D.Outline>();
                if (outline.GetFirstChild<D.SolidFill>() != null)
                {
                    D.SolidFill solidFill = outline.GetFirstChild<D.SolidFill>();
                    D.SolidFill newSolidFill = new D.SolidFill( new SchemeColor() { Val = schemeColorVal });
                    outline.ReplaceChild<SolidFill>(newSolidFill, solidFill);
                }
                else
                {
                    D.SolidFill newSolidFill = new D.SolidFill(new SchemeColor() { Val = schemeColorVal });
                    outline.AddChild(newSolidFill);
                }                
            }
            else
            {                 
                D.SolidFill solidFill = new D.SolidFill(new SchemeColor() { Val = schemeColorVal });
                D.Outline outline = new D.Outline();
                outline.AddChild(solidFill);
                props.AddChild(outline);
            }

        }
        //TODO: Fix int conversion, values are messed up
        public static void SetFillTransparency(this P.Shape shape, int transparency)
        {
            int opacity = (100 - transparency) * 1000;
            
            D.SolidFill solidFill = shape.ShapeProperties.GetFirstChild<D.SolidFill>();
            if (solidFill != null)
            {
                solidFill.Append(new D.Alpha() { Val = opacity });
            }
            else
            {
                shape.ShapeProperties.Append(new D.SolidFill(
                    new D.SchemeColor(){ Val = SchemeColorValues.Accent1},
                    new D.Alpha() { Val = opacity }
                    )
                   );
            }

        }
        
        public static void SetOutlineWidth(this P.Shape shape, int width)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetOutlineWidth(width);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetOutlineWidth(width);
                shape.AddChild(shapeProps);                
            }
        }

        // TODO: error proof
        public static void SetOutlineDash(this P.Shape shape)
        {
            D.Outline outline = shape.ShapeProperties.GetFirstChild<D.Outline>();
            D.PresetDash dash = new D.PresetDash() { Val = PresetLineDashValues.Dash };
            outline.AddChild(dash);
        }

        // TODO: Fix
        // TODO: error proof
        public static void SetOutlineDot(this P.Shape shape)
        {
            D.Outline outline = shape.ShapeProperties.GetFirstChild<D.Outline>();
            D.PresetDash dash = new D.PresetDash() { Val = PresetLineDashValues.Dot };
            outline.Append(dash);

        }
        /// <summary>
        /// Sets shape width
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetWidth(this P.Shape shape, int inches)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetWidth(inches);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetWidth(inches);
                shape.AddChild(shapeProps);
            }
        }

        public static void SetWidth(this P.Shape shape, Int64Value pts)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetWidth(pts);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetWidth(pts);
                shape.AddChild(shapeProps);
            }
        }

        public static Int64Value GetWidth(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                Int64Value width = shape.ShapeProperties.GetWidth();
                return width;
            }
            else
            {
                throw new NullReferenceException();
            }
        }
             

        /// <summary>
        /// Sets shape height
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetHeight(this P.Shape shape, int inches)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetHeight(inches);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetHeight(inches);
                shape.AddChild(shapeProps);
            }
        }

        public static void SetHeight(this P.Shape shape, Int64Value pts)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetHeight(pts);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetHeight(pts);
                shape.AddChild(shapeProps);
            }
        }

        public static Int64Value GetHeight(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                Int64Value height = shape.ShapeProperties.GetHeight();
                return height;
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        /// <summary>
        /// Sets horizontal position
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetHorizontalPosition(this P.Shape shape, int position)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetHorizontalPosition(position);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetHorizontalPosition(position);
                shape.AddChild(shapeProps);
            }
        }

        /// <summary>
        /// Overload for the SetHorizontalPosition method, accepting type Int64Value as position parameter
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="pts"></param>

        public static void SetHorizontalPosition(this P.Shape shape, Int64Value pts)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetHorizontalPosition(pts);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetHorizontalPosition(pts);
                shape.AddChild(shapeProps);
            }
        }
        /// <summary>
        /// Gets the horizontal position in inches
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        public static Int64Value GetHorizontalPosition(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                Int64Value hPos = shape.ShapeProperties.GetHorizontalPosition();
                return hPos;
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void SetHorizontalPositionRelativeTo(this P.Shape shape, P.Shape anchorShape, int relativeDistance)
        {
            if (shape.ShapeProperties != null)
            {
                shape.SetHorizontalPosition(pts: anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetWidth() + relativeDistance * 914400);
            }
        }

        /// <summary>
        /// Sets vertical position
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetVerticalPosition(this P.Shape shape, int inches)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetVerticalPosition(inches);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetVerticalPosition(inches);
                shape.AddChild(shapeProps);
            }
        }

        public static Int64Value GetVerticalPosition(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                Int64Value vPos = shape.ShapeProperties.GetVerticalPosition();
                return vPos;
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void SetVerticalPosition(this P.Shape shape, Int64Value pts)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetVerticalPosition(pts);
            }
            else
            {
                P.ShapeProperties shapeProps = new P.ShapeProperties();
                shapeProps.SetVerticalPosition(pts);
                shape.AddChild(shapeProps);
            }
        }


        public static void SetVerticalPositionRelativeTo(this P.Shape shape, P.Shape anchorShape, int relativeDistance)
        {
            if (shape.ShapeProperties != null)
            {
                shape.SetVerticalPosition(pts: anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() + relativeDistance * 914400);
            }
        }

       

        public static void AlignLeft(this P.Shape shape, P.Shape anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            shape.SetHorizontalPosition(anchorShape.GetHorizontalPosition());
        }


        public static void AlignTop(this P.Shape shape, P.Shape anchorShape)
        {
            // Anchor shpae SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            shape.SetVerticalPosition(anchorShape.GetVerticalPosition());
        }

        public static void AlignBottom(this P.Shape shape, P.Shape anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            Int64 vPos = (Int64)anchorShape.GetVerticalPosition();
            Int64 height = (Int64)anchorShape.GetHeight();
            Int64 bottomPos = vPos + height;
            shape.SetVerticalPosition(bottomPos);            
        }

        public static void AlignRight(this P.Shape shape, P.Shape anchorShape)
        {            
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            Int64 hPos = (Int64)anchorShape.GetHorizontalPosition();
            Int64 width = (Int64)anchorShape.GetWidth();
            Int64 rightPos = hPos + width;
            shape.SetHorizontalPosition(rightPos);
        }

        public static void CopyHeight(this P.Shape shape, P.Shape anchorShape)
        {
            shape.SetHeight(anchorShape.GetHeight());
        }

        public static void CopyWidth(this P.Shape shape, P.Shape anchorShape)
        {
            shape.SetWidth(anchorShape.GetWidth()); 
        }

        public static void CopyVerticalPosition(this P.Shape shape, P.Shape anchorShape)
        {
            shape.SetVerticalPosition(anchorShape.GetVerticalPosition());
        }

        public static void CopyHorizontalPosition(this P.Shape shape, P.Shape anchorShape)
        {
            shape.SetHorizontalPosition(anchorShape.GetVerticalPosition());
        }

        public static void CopyPosition(this P.Shape shape, P.Shape anchorShape)
        {
            shape.CopyHorizontalPosition(anchorShape);
            shape.CopyVerticalPosition(anchorShape);

        }

        public static void CopyDimensions(this P.Shape shape, P.Shape anchorShape)
        {
            shape.CopyWidth(anchorShape);
            shape.CopyHeight(anchorShape);
        }

        public static D.SolidFill GetFill(this P.Shape shape)
        {
            if (shape.GetFirstChild<P.ShapeProperties>() != null)
            {
                P.ShapeProperties props = shape.GetFirstChild<P.ShapeProperties>();
                if (props.GetFirstChild<D.SolidFill>() != null)
                {
                    return (D.SolidFill)props.GetFirstChild<D.SolidFill>().Clone();
                }

                else
                {
                    return null;
                }
            }

            else
            {
                return null;
            }
        }

        public static D.FillReference GetFillReference(this P.Shape shape)
        {
            if (shape.GetFirstChild<P.ShapeStyle>() != null)
            {
                P.ShapeStyle style = shape.GetFirstChild<P.ShapeStyle>();
                if (style.GetFirstChild<D.FillReference>() != null)
                {
                    return (D.FillReference)style.GetFirstChild<D.FillReference>().Clone();
                }

                else
                {
                    return null;
                }
            }

            else
            {
                return null;
            }
        }
        public static D.SolidFill GetOutlineFill(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                P.ShapeProperties props = shape.ShapeProperties;
                if (props.GetFirstChild<D.Outline>() != null)
                {
                    D.Outline outline = props.GetFirstChild<D.Outline>();
                    if(outline.GetFirstChild<D.SolidFill>() != null)
                    {
                        return (D.SolidFill)outline.GetFirstChild<D.SolidFill>().Clone();
                    }
                    else
                    {
                        throw new Exception("Solid Fill element is null");
                    }
                }

                else
                {
                    throw new Exception("Outline is null");
                }
            }

            else
            {
                return null;
            }
        }

        public static D.Outline GetOutline(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                P.ShapeProperties props = shape.ShapeProperties;
                if (props.GetFirstChild<D.Outline>() != null)
                {
                    return (D.Outline)props.GetFirstChild<D.Outline>().Clone();
                }

                else
                {
                    throw new NullReferenceException();
                }
            }

            else
            {
                throw new NullReferenceException();
            }
        }

        public static int GetOutlineWidth(this P.Shape shape)
        {
            if (shape.ShapeProperties != null)
            {
                P.ShapeProperties props = shape.ShapeProperties;
                if (props.GetFirstChild<D.Outline>() != null)
                {
                    return props.GetFirstChild<D.Outline>().Width;
                }

                else
                {
                    throw new NullReferenceException();
                }
            }

            else
            {
                throw new NullReferenceException();
            }
        }

        public static void CopyOutlineWidth(this P.Shape shape, P.Shape anchorShape)
        {
            if (shape.ShapeProperties != null)
            {
                shape.ShapeProperties.SetOutlineWidth(anchorShape.GetOutlineWidth());
            }
        }

        public static void CopyOutline(this P.Shape shape, P.Shape anchorShape)
        {
            if (shape.ShapeProperties != null)
            {
                if (shape.ShapeProperties.GetFirstChild<D.Outline>() != null)
                {
                    shape.ShapeProperties.ReplaceChild<D.Outline>(anchorShape.GetOutline(), shape.ShapeProperties.GetFirstChild<D.Outline>());
                }
                else
                {
                    shape.ShapeProperties.AddChild(anchorShape.GetOutline());
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void CopyOutlineFill(this P.Shape shape, P.Shape anchorShape)
        {
            if (shape.ShapeProperties != null)
            {
                if (shape.ShapeProperties.GetFirstChild<D.Outline>() != null) 
                { 
                    D.Outline outline = shape.ShapeProperties.GetFirstChild<D.Outline>();
                    if (outline.GetFirstChild<D.SolidFill>() != null)
                    {
                        outline.ReplaceChild(anchorShape.GetOutlineFill(), outline.GetFirstChild<D.SolidFill>());
                    }
                    else
                    {
                        outline.AddChild(anchorShape.GetOutlineFill());
                    }
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void CopyShapeFill(this P.Shape shape, P.Shape anchorShape)
        {
            if (shape.ShapeProperties != null)
            {
                if (shape.ShapeProperties.GetFirstChild<D.SolidFill>() != null)
                {
                    shape.ShapeProperties.ReplaceChild<D.SolidFill>(anchorShape.GetFill(), shape.ShapeProperties.GetFirstChild<D.SolidFill>());
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }      

    }

    
}
