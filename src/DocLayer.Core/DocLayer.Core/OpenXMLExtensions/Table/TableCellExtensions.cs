using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using D = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLExtensions
{
    public static class TableCellExtensions
    {
        public static string GetText(this TableCell cell)
        {
            string text = "";
            if (cell.GetFirstChild<D.TextBody>() != null)
            {

                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (body.GetFirstChild<D.Paragraph>() is not null)
                {
                    foreach (D.Paragraph p in body.Elements<D.Paragraph>()) 
                    {
                        text += p.GetText();
                    }
                    return text;
                }
                return text;

            }
            return text;

        }
        public static void AddParagraph(this TableCell cell, string text)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                D.Paragraph para = new D.Paragraph();
                para.SetAlignLeft();
                para.SetText(text);
                para.SetEndProps();

                body.AppendChild<D.Paragraph>(para);
            }
            else throw new Exception("Textbody not found");

        }

        public static void SetText(this TableCell cell, string text)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                D.Paragraph para = new D.Paragraph();
                para.SetAlignLeft();
                para.SetText(text);
                para.SetEndProps();

                body.AddChild(para);
            }
            else throw new Exception("Textbody not found");
        }

        public static void AddParagraphAt(this TableCell cell, string text, int addPosition)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (addPosition > 0 && addPosition <= body.Elements<D.Paragraph>().Count())
                {
                    D.Paragraph para = new D.Paragraph();
                    para.SetAlignLeft();
                    para.SetText(text);
                    para.SetEndProps();

                    body.InsertAt<D.Paragraph>(para, addPosition - 1);
                }
                else throw new Exception("Position not valid");
                
            }
            else throw new Exception("Textbody not found");
        }

        public static void DeleteParagraphAt(this TableCell cell, int position)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (position > 0 && position <= body.Elements<D.Paragraph>().Count())
                {
                    body.Elements<D.Paragraph>().ElementAt<D.Paragraph>(position - 1).Remove();

                    
                    if (body.GetFirstChild<D.Paragraph>() == null)
                    {
                        // if there is no paragraph left 
                        // add an empty paragraph

                        D.Paragraph paragraph = new D.Paragraph();
                        paragraph.SetEndProps();
                        body.AddChild(paragraph);
                    }
                }
                else throw new Exception("Position not valid");
            }
            else throw new Exception("Textbody not found");
        }

        /// <summary>
        /// Deletes all paragraphs of the cell which have text and adds an empty paragraph
        /// </summary>
        /// <param name="cell"></param>
        /// <exception cref="Exception"></exception>
        public static void DeleteAllText(this TableCell cell)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (body.Descendants<D.Run>().Count() > 0)
                {
                    foreach (D.Paragraph paragraph in body.Elements<D.Paragraph>())
                    {
                        if (paragraph.GetFirstChild<D.Run>() != null)
                        {
                            // if the paragraph has text, delete
                            paragraph.Remove();
                        }
                    }

                    if (body.GetFirstChild<D.Paragraph>() == null)
                    {
                        // if there is no paragraph left 
                        // add an empty paragraph

                        D.Paragraph paragraph = new D.Paragraph();
                        paragraph.SetEndProps();
                        body.AddChild(paragraph);
                    }

                }
                else throw new Exception("Text not found");         
            }
            else throw new Exception("Textbody not found");

        }

        public static bool TryGetRunContaining(this D.TableCell cell, string text, out D.Run runContainingText)
        {
            D.Run run;
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (body.GetFirstChild<D.Paragraph>() is not null)
                {
                    foreach (D.Paragraph paragraph in body.Elements<D.Paragraph>())
                    {
                        if(paragraph.TryBreakAndGetRunContainingText(text, out run))
                        {
                            runContainingText = run;
                            return true;
                        }
                    }
                }
                runContainingText = null;
                return false;
            }
            runContainingText = null;
            return false;
        }
        public static void SetSchemeFill(this D.TableCell cell, int accentNum)
        {
            if (cell.GetFirstChild<D.TableCellProperties>() != null)
            {
                D.TableCellProperties props = cell.GetFirstChild<D.TableCellProperties>();
                if (props.GetFirstChild<D.SolidFill>() != null)
                {
                    props.GetFirstChild<D.SolidFill>().SetSchemeFill(accentNum);
                }
                else
                {
                    D.SolidFill fill = new D.SolidFill();
                    fill.SetSchemeFill(accentNum);
                    props.AddChild(fill);
                }
            }
            else throw new Exception("Table cell properties not found");
        }

        /// <summary>
        /// sets font color of all text in the cell
        /// </summary>
        public static void SetFontColorScheme(this D.TableCell cell, int accentNum)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (body.Descendants<D.Run>().Count() > 0)
                {
                    // Runs exist in the textbody

                    foreach (D.Run run in body.Descendants<D.Run>())
                    {
                        run.SetRunSchemeFill(accentNum);
                    }
                }
            }
        }

        public static void SetFontColorHex(this D.TableCell cell, string hexCode)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (body.Descendants<D.Run>().Count() > 0)
                {
                    // Runs exist in the textbody

                    foreach (D.Run run in body.Descendants<D.Run>())
                    {
                        run.SetRunHexFill(hexCode);
                    }
                }
            }
        }



        public static void SetFontSize(this D.TableCell cell, int size)
        {
            if (cell.GetFirstChild<D.TextBody>() != null)
            {
                D.TextBody body = cell.GetFirstChild<D.TextBody>();
                if (body.Descendants<D.Run>().Count() > 0)
                {
                    // Runs exist in the textbody

                    foreach (D.Run run in body.Descendants<D.Run>())
                    {
                        run.SetRunSize(size);
                    }
                }
            }
        }


        public static void SetBottomBorder(this D.TableCell cell, int width)
        {
            if (cell.GetFirstChild<D.TableCellProperties>() != null)
            {
                D.TableCellProperties props = cell.GetFirstChild<D.TableCellProperties>();
                D.BottomBorderLineProperties border = new D.BottomBorderLineProperties() { Width = width * 12700, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };
                border.AddChild(new D.SolidFill(new D.SchemeColor() { Val = SchemeColorValues.Accent1 }));
                border.AddChild(new D.PresetDash() { Val = PresetLineDashValues.Solid });
                border.AddChild(new D.Round());
                border.AddChild(new D.HeadEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });
                border.AddChild(new D.TailEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });

                props.AddChild(border);
            }
            else throw new Exception("Table cell not found");

        }

        public static void SetTopBorder(this D.TableCell cell, int width)
        {
            if (cell.GetFirstChild<D.TableCellProperties>() != null)
            {
                D.TableCellProperties props = cell.GetFirstChild<D.TableCellProperties>();
                D.TopBorderLineProperties border = new D.TopBorderLineProperties() { Width = width * 12700, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };
                border.AddChild(new D.SolidFill(new D.SchemeColor() { Val = SchemeColorValues.Accent1 }));
                border.AddChild(new D.PresetDash() { Val = PresetLineDashValues.Solid });
                border.AddChild(new D.Round());
                border.AddChild(new D.HeadEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });
                border.AddChild(new D.TailEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });

                props.AddChild(border);
            }
            else throw new Exception("Table cell not found");

        }

        public static void SetLeftBorder(this D.TableCell cell, int width)
        {
            if (cell.GetFirstChild<D.TableCellProperties>() != null)
            {
                D.TableCellProperties props = cell.GetFirstChild<D.TableCellProperties>();
                D.LeftBorderLineProperties border = new D.LeftBorderLineProperties() { Width = width * 12700, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };
                border.AddChild(new D.SolidFill(new D.SchemeColor() { Val = SchemeColorValues.Accent1 }));
                border.AddChild(new D.PresetDash() { Val = PresetLineDashValues.Solid });
                border.AddChild(new D.Round());
                border.AddChild(new D.HeadEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });
                border.AddChild(new D.TailEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });

                props.AddChild(border);
            }
            else throw new Exception("Table cell not found");

        }

        public static void SetRightBorder(this D.TableCell cell, int width)
        {
            if (cell.GetFirstChild<D.TableCellProperties>() != null)
            {
                D.TableCellProperties props = cell.GetFirstChild<D.TableCellProperties>();
                D.RightBorderLineProperties border = new D.RightBorderLineProperties() { Width = width * 12700, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };
                border.AddChild(new D.SolidFill(new D.SchemeColor() { Val = SchemeColorValues.Accent1 }));
                border.AddChild(new D.PresetDash() { Val = PresetLineDashValues.Solid });
                border.AddChild(new D.Round());
                border.AddChild(new D.HeadEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });
                border.AddChild(new D.TailEnd() { Type = LineEndValues.None, Width = LineEndWidthValues.Medium, Length = LineEndLengthValues.Medium });

                props.AddChild(border);
            }
            else throw new Exception("Table cell not found");

        }

        public static void SetBottomBorderColorScheme(this D.TableCell cell, int accentNum)
        {
            if (cell.Descendants<D.BottomBorderLineProperties>().Count() > 0)
            {
                D.BottomBorderLineProperties border = cell.TableCellProperties.BottomBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetSchemeFill(accentNum);
                }
            }
            else throw new Exception("Table cell border not found");

        }

        public static void SetBottomBorderColorHex(this D.TableCell cell, string rgbColorHex)
        {
            if (cell.Descendants<D.BottomBorderLineProperties>().Count() > 0)
            {
                D.BottomBorderLineProperties border = cell.TableCellProperties.BottomBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetHexFill(rgbColorHex);
                }
            }
            else throw new Exception("Table cell border not found");

        }

        public static void SetTopBorderColorScheme(this D.TableCell cell, int accentNum)
        {
            if (cell.Descendants<D.TopBorderLineProperties>().Count() > 0)
            {
                D.TopBorderLineProperties border = cell.TableCellProperties.TopBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetSchemeFill(accentNum);
                }
            }
            else throw new Exception("Table cell border not found");

        }

        public static void SetTopBorderColorHex(this D.TableCell cell, string rgbColorHex)
        {
            if (cell.Descendants<D.TopBorderLineProperties>().Count() > 0)
            {
                D.TopBorderLineProperties border = cell.TableCellProperties.TopBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetHexFill(rgbColorHex);
                }
            }
            else throw new Exception("Table cell border not found");
        }

        public static void SetRightBorderColorScheme(this D.TableCell cell, int accentNum)
        {
            if (cell.Descendants<D.RightBorderLineProperties>().Count() > 0)
            {
                D.RightBorderLineProperties border = cell.TableCellProperties.RightBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetSchemeFill(accentNum);
                }
            }
            else throw new Exception("Table cell border not found");

        }

        public static void SetRightBorderColorHex(this D.TableCell cell, string rgbColorHex)
        {
            if (cell.Descendants<D.RightBorderLineProperties>().Count() > 0)
            {
                D.RightBorderLineProperties border = cell.TableCellProperties.RightBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetHexFill(rgbColorHex);
                }
            }
            else throw new Exception("Table cell border not found");
        }

        public static void SetLeftBorderColorHex(this D.TableCell cell, string rgbColorHex)
        {
            if (cell.Descendants<D.LeftBorderLineProperties>().Count() > 0)
            {
                D.LeftBorderLineProperties border = cell.TableCellProperties.LeftBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetHexFill(rgbColorHex);
                }
            }
            else throw new Exception("Table cell border not found");
        }

        public static void SetLeftBorderColorScheme(this D.TableCell cell, int accentNum)
        {
            if (cell.Descendants<D.LeftBorderLineProperties>().Count() > 0)
            {
                D.LeftBorderLineProperties border = cell.TableCellProperties.LeftBorderLineProperties;
                if (border.GetFirstChild<D.SolidFill>() != null)
                {
                    border.GetFirstChild<D.SolidFill>().SetSchemeFill(accentNum);
                }
            }
            else throw new Exception("Table cell border not found");

        }
        

        
    }
}
