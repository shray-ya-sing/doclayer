using D = DocumentFormat.OpenXml.Drawing;


namespace OpenXMLExtensions
{
    public static class ParagraphExtensions
    {
        public static void SetBulletFont(this D.Paragraph paragraph)
        {
            D.ParagraphProperties props = paragraph.ParagraphProperties;
            if (props.GetFirstChild<D.BulletFont>() != null )
            {
                D.BulletFont buFont = props.GetFirstChild<D.BulletFont>();
                buFont.Typeface = "Arial";
                buFont.Panose = "020B0604020202020204";
                buFont.PitchFamily = 34;
                buFont.CharacterSet = 0;

            }

            else
            {
                D.BulletFont buFont = new D.BulletFont();
                buFont.Typeface = "Arial";
                buFont.Panose = "020B0604020202020204";
                buFont.PitchFamily = 34;
                buFont.CharacterSet = 0;
                props.AddChild(buFont);
            }
        }

        public static void SetRoundBulletCharacter(this D.Paragraph paragraph)
        {
            D.ParagraphProperties props = paragraph.ParagraphProperties;
            if (props.GetFirstChild<D.CharacterBullet>() != null)
            {
                D.CharacterBullet buChar = props.GetFirstChild<D.CharacterBullet>();
                buChar.Char = "•";

            }

            else
            {
                D.CharacterBullet buChar =new D.CharacterBullet();
                buChar.Char = "•";
                props.AddChild(buChar);

            }

        }

        public static void SetDashBulletCharacter(this D.Paragraph paragraph)
        {
            D.ParagraphProperties props = paragraph.ParagraphProperties;
            if (props.GetFirstChild<D.CharacterBullet>() != null)
            {
                D.CharacterBullet buChar = props.GetFirstChild<D.CharacterBullet>();
                buChar.Char = "─";

            }

            else
            {
                D.CharacterBullet buChar = new D.CharacterBullet();
                buChar.Char = "─";
                props.AddChild(buChar);

            }

        }

        //TODO: figure out how to convert sizes and make dynamic
        public static void SetBulletIndent(this D.Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
            {
                D.ParagraphProperties props = paragraph.ParagraphProperties;
                props.LeftMargin = 285750;
                props.Indent = -285750;

            }

            else
            {
                D.ParagraphProperties props = new D.ParagraphProperties();
                props.LeftMargin = 285750;
                props.Indent = -285750;
                paragraph.AddChild(props);  

            }
        }


        public static void SetAlignCenter(this D.Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
            {
                D.ParagraphProperties props = paragraph.ParagraphProperties;
                props.Alignment =D.TextAlignmentTypeValues.Center;

            }
            else
            {
                D.ParagraphProperties props = new D.ParagraphProperties();
                props.Alignment = D.TextAlignmentTypeValues.Center;
                paragraph.AddChild(props);
            }
            

        }

        public static void SetAlignLeft(this D.Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
            {
                D.ParagraphProperties props = paragraph.ParagraphProperties;
                props.Alignment = D.TextAlignmentTypeValues.Left;

            }
            else
            {
                D.ParagraphProperties props = new D.ParagraphProperties();
                props.Alignment = D.TextAlignmentTypeValues.Left;
                paragraph.AddChild(props);
            }

        }

        public static void SetAlignRight(this D.Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
            {
                D.ParagraphProperties props = paragraph.ParagraphProperties;
                props.Alignment = D.TextAlignmentTypeValues.Right;

            }
            else
            {
                D.ParagraphProperties props = new D.ParagraphProperties();
                props.Alignment = D.TextAlignmentTypeValues.Right;
                paragraph.AddChild(props);
            }

        }

        public static void SetEndProps(this D.Paragraph paragraph)
        {
            if (paragraph.GetFirstChild<D.EndParagraphRunProperties>() != null)
            {
                D.EndParagraphRunProperties props = paragraph.GetFirstChild<D.EndParagraphRunProperties>();
                props.Language = "en-US";

            }
            else
            {
                D.EndParagraphRunProperties props = new D.EndParagraphRunProperties();
                props.Language = "en-US";
                paragraph.AddChild(props);

            }
            

        }

        public static void SetText(this D.Paragraph paragraph, string text)
        {
            D.Run textRun = new D.Run();
            textRun.SetRunEnglish();
            textRun.SetRunSize(14);
            textRun.RunProperties.Dirty = false;
            textRun.AddText(text);

            paragraph.AddChild(textRun); // add in the correct position

        }

        public static void AddText(this D.Paragraph paragraph, string text)
        {
            D.Run textRun = new D.Run();
            textRun.SetRunEnglish();
            textRun.SetRunSize(14);
            textRun.RunProperties.Dirty = false;
            textRun.AddText(text);
            paragraph.AppendChild<D.Run>(textRun);

        }

        /// <summary>
        /// Set text color of whole paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        public static void SetParagraphSchemeFill(this D.Paragraph paragraph, int accentNum)
        {
            foreach (D.Run run in paragraph.Elements<D.Run>())
            {
                run.SetRunSchemeFill(accentNum);
            }
        }



        /// <summary>
        /// Sets bullet to paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        public static void SetBullet(this D.Paragraph paragraph)
        {
            paragraph.SetBulletIndent();
            paragraph.SetBulletFont();
            paragraph.SetRoundBulletCharacter();

        }

        /// <summary>
        /// Sets line spacing to arg points
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="lineSpcPts"></param>
        public static void SetLineSpacing(this D.Paragraph paragraph, int lineSpcPts)
        {
            if (paragraph.ParagraphProperties != null)
            {
                paragraph.ParagraphProperties.SetLineSpacing(lineSpcPts);
            }
            else
            {
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties();
                paragraph.AddChild(paragraphProperties);
                paragraphProperties.SetLineSpacing(lineSpcPts);
            }
            
        }

        /// <summary>
        /// Sets space before paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="pts"></param>
        public static void SetSpaceBefore(this D.Paragraph paragraph, int pts)
        {
            if (paragraph.ParagraphProperties != null)
            {
                paragraph.ParagraphProperties.SetSpaceBefore(pts);
            }
            else
            {
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties();
                paragraph.AddChild(paragraphProperties);
                paragraphProperties.SetSpaceBefore(pts);
            }

        }

        /// <summary>
        /// Sets space after pargraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="pts"></param>

        public static void SetSpaceAfter(this D.Paragraph paragraph, int pts)
        {
            if (paragraph.ParagraphProperties != null)
            {
                paragraph.ParagraphProperties.SetSpaceAfter(pts);
            }
            else
            {
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties();
                paragraph.AddChild(paragraphProperties);
                paragraphProperties.SetSpaceAfter(pts);
            }

        }



        /// <summary>
        /// Set bullet color by hex code
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="hex"></param>
        public static void SetBulletColorHex(this D.Paragraph paragraph, string hex)
        {
            if (paragraph.ParagraphProperties != null)
            {
                paragraph.ParagraphProperties.SetBulletColorHex(hex);
            }
            else
            {
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties();
                paragraph.AddChild(paragraphProperties);
                paragraphProperties.SetBulletColorHex(hex);
            }

        }

        /// <summary>
        /// Sets bullet color by scheme color
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="accentNum"></param>

        public static void SetBulletColorScheme(this D.Paragraph paragraph, int accentNum)
        {
            if (paragraph.ParagraphProperties != null)
            {
                paragraph.ParagraphProperties.SetBulletColorScheme(accentNum);
            }
            else
            {
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties();
                paragraph.AddChild(paragraphProperties);
                paragraphProperties.SetBulletColorScheme(accentNum);
            }
        }

        /// <summary>
        /// Sets line spacing to single spaced
        /// </summary>
        /// <param name="paragraph"></param>

        public static void SetSingleSpacing(this D.Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties != null)
            {
                paragraph.ParagraphProperties.SetSingleSpacing();
            }
            else
            {
                D.ParagraphProperties paragraphProperties = new D.ParagraphProperties();
                paragraph.AddChild(paragraphProperties);
                paragraph.ParagraphProperties.SetSingleSpacing();
            }
        }

        public static bool TryGetRunContaining(this D.Paragraph paragraph, string text, out D.Run runContainingText)
        {
            // returns the first instance of run containing text, regardless if multiple runs contain the same text

            if (text != "" && text != " " && text != null)
            {
                if (paragraph.GetFirstChild<D.Run>() != null)
                {
                    foreach (D.Run run in paragraph.Elements<D.Run>())
                    {
                        foreach (D.Text runText in run.Elements<D.Text>())
                        {
                            if (runText.InnerText.Contains(text, StringComparison.OrdinalIgnoreCase))
                            {
                                runContainingText = run;
                                return true;
                            }
                        }                                               
                    }

                    runContainingText = null;
                    return false; //no run in the paragraph has the text
      
                }
                runContainingText = null;
                return false; //paragraph does not have any runs and thus no run in the paragraph has the text
            }
            else throw new Exception("Text cannot be null or empty");
        }

        public static bool TryBreakAndGetRunContainingText(this D.Paragraph paragraph, string text, out D.Run returnedRun)
        {
            if (text != "" && text != " " && text is not null)
            {
                D.Run run;
                if (paragraph.TryGetRunContaining(text, out run)) // check if the paragraph contains the run
                {
                    // if the paragraph contains the run

                    int start = run.InnerText.ToString().IndexOf(text, StringComparison.OrdinalIgnoreCase);
                    int length = text.Length;
                    string origText = run.InnerText.ToString();


                    string textBefore = origText.Substring(0, start);
                    string targetText = origText.Substring(start, length);
                    string textAfter = origText.Substring(start + length);

                    if (textBefore != "" && textAfter != "") // i.e. the run needs to be split into 3
                    {
                        D.Text newRunTextBefore = new D.Text(textBefore);
                        D.Run newRunBefore = (D.Run)run.Clone(); // clone to preserve formatting 
                        newRunBefore.GetFirstChild<D.Text>().Remove(); // A single run should never have more than 1 text element
                        newRunBefore.AddChild(newRunTextBefore);
                        run.InsertBeforeSelf(newRunBefore);

                        D.Text targetRunText = new D.Text(targetText);
                        D.Run targetRun = (D.Run)run.Clone(); // clone to preserve formatting 
                        targetRun.GetFirstChild<D.Text>().Remove(); // A single run should never have more than 1 text element
                        targetRun.AddChild(targetRunText);
                        run.InsertBeforeSelf(targetRun);


                        D.Text newRunTextAfter = new D.Text(textAfter);
                        D.Run newRunAfter = (D.Run)run.Clone(); // clone to preserve formatting 
                        newRunAfter.GetFirstChild<D.Text>().Remove(); // A single run should never have more than 1 text element
                        newRunAfter.AddChild(newRunTextAfter);
                        run.InsertBeforeSelf(newRunAfter);

                        run.Remove();
                        returnedRun = targetRun;
                        return true;
                    }

                    else
                    {
                        // return run as is without changes
                        returnedRun = run;
                        return true;
                    }
                }
                else 
                { 
                    returnedRun = null;
                    return false;
                }
            }
            else throw new Exception("Text cannot be null or empty");
        }

        public static void BoldSelectedText(this D.Paragraph paragraph, string text)
        {
            D.Run run;
            if (paragraph.TryBreakAndGetRunContainingText(text, out run))
            {
                run.SetRunBold();
            }
        }

        public static void ItalicizeSelectedText(this D.Paragraph paragraph, string text)
        {
            D.Run run;
            if (paragraph.TryBreakAndGetRunContainingText(text, out run))
            {
                run.SetRunItalic();
            }
        }

        public static void SetSelectedTextColorScheme(this D.Paragraph paragraph, string text, int accentNum)
        {
            D.Run run;
            if (paragraph.TryBreakAndGetRunContainingText(text, out run))
            {
                run.SetRunSchemeFill(accentNum);
            }
        }

        public static void SetSelectedTextColorHex(this D.Paragraph paragraph, string text, string hex)
        {
            D.Run run;
            if (paragraph.TryBreakAndGetRunContainingText(text, out run))
            {
                run.SetRunHexFill(hex);
            }
        }

        public static string GetText(this D.Paragraph paragraph)
        {
            string text = "";
            if (paragraph.GetFirstChild<D.Run>() is not null)
            {
               foreach( D.Run run in paragraph.Elements<D.Run>())
                {
                    text += run.GetInnerText();
                }

               return text;
            }

            return text; // Unsafe: may be problematic to return empty string
        }


    }
}
