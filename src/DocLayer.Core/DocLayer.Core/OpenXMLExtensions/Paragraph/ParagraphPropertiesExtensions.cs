using D = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLExtensions
{
    public static class ParagraphPropertiesExtensions
    {
        /// <summary>
        /// Sets line spacing
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <param name="lineSpacingPts"></param>
        public static void SetLineSpacing(this D.ParagraphProperties paragraphProperties, int lineSpacingPts)
        {
            if (paragraphProperties.GetFirstChild<D.LineSpacing>() != null)
            {
                D.LineSpacing lineSpacing = paragraphProperties.GetFirstChild<D.LineSpacing>();
                D.SpacingPoints currentSpcPts = lineSpacing.GetFirstChild<D.SpacingPoints>();
                D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts*100 };
                lineSpacing.ReplaceChild<D.SpacingPoints>(newSpcPts, currentSpcPts);
            }

            else
            {
                D.LineSpacing lineSpacing = new D.LineSpacing();
                D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                lineSpacing.AddChild(newSpcPts);
                paragraphProperties.AddChild(lineSpacing);

            }
        }

        /// <summary>
        /// Sets space after
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <param name="lineSpacingPts"></param>

        public static void SetSpaceAfter(this D.ParagraphProperties paragraphProperties, int lineSpacingPts)
        {
            if (paragraphProperties.GetFirstChild<D.SpaceAfter>() != null)
            {
                D.SpaceAfter spaceAfter = paragraphProperties.GetFirstChild<D.SpaceAfter>();
                if (spaceAfter.GetFirstChild<D.SpacingPoints>() != null)
                {
                    D.SpacingPoints currentSpcPts = spaceAfter.GetFirstChild<D.SpacingPoints>();
                    D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                    spaceAfter.ReplaceChild<D.SpacingPoints>(newSpcPts, currentSpcPts);
                }
                else
                {
                    D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                    spaceAfter.AddChild(newSpcPts);

                }
            }

            else
            {
                D.SpaceAfter spaceAfter = new D.SpaceAfter();
                D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                spaceAfter.AddChild(newSpcPts);
                paragraphProperties.AddChild(spaceAfter);
            }
        }
        /// <summary>
        /// Sets space before line
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <param name="lineSpacingPts"></param>
        public static void SetSpaceBefore(this D.ParagraphProperties paragraphProperties, int lineSpacingPts)
        {
            if (paragraphProperties.GetFirstChild<D.SpaceBefore>() != null)
            {
                D.SpaceBefore spaceBefore = paragraphProperties.GetFirstChild<D.SpaceBefore>();
                if (spaceBefore.GetFirstChild<D.SpacingPoints>() != null)
                {
                    D.SpacingPoints currentSpcPts = spaceBefore.GetFirstChild<D.SpacingPoints>();
                    D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                    spaceBefore.ReplaceChild<D.SpacingPoints>(newSpcPts, currentSpcPts);
                }
                else
                {
                    D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                    spaceBefore.AddChild(newSpcPts);

                }
            }

            else
            {
                D.SpaceBefore spaceBefore = new D.SpaceBefore();
                D.SpacingPoints newSpcPts = new D.SpacingPoints() { Val = lineSpacingPts * 100 };
                spaceBefore.AddChild(newSpcPts);
                paragraphProperties.AddChild(spaceBefore);
            }
        }

        /// <summary>
        /// Sets bullet Color
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <param name="hex"></param>

        public static void SetBulletColorHex(this D.ParagraphProperties paragraphProperties, string hex)
        {
            if (paragraphProperties.GetFirstChild<D.BulletColor>() != null)
            {
                D.BulletColor buColor = paragraphProperties.GetFirstChild<D.BulletColor>();
                if (buColor.GetFirstChild<D.RgbColorModelHex>() != null)
                {
                    buColor.GetFirstChild<D.RgbColorModelHex>().Val = hex;
                }
                else if (buColor.GetFirstChild<D.SchemeColor>() != null)
                {
                    D.SchemeColor currentColor = buColor.GetFirstChild<D.SchemeColor>();
                    buColor.RemoveChild(currentColor);

                    D.RgbColorModelHex rgbColor = new D.RgbColorModelHex() { Val = hex };
                    buColor.AddChild(rgbColor);
                }
                else
                {
                    D.RgbColorModelHex rgbColor = new D.RgbColorModelHex() { Val = hex };
                    buColor.AddChild(rgbColor);
                }
            }

            else
            {
                D.BulletColor buColor = new D.BulletColor();
                D.RgbColorModelHex rgbColor = new D.RgbColorModelHex() { Val = hex };
                buColor.AddChild(rgbColor);
                paragraphProperties.AddChild(buColor);
            }

        }

        /// <summary>
        /// Sets bullet color by scheme color
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <param name="accentNum"></param>

        public static void SetBulletColorScheme(this D.ParagraphProperties paragraphProperties, int accentNum)
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

            if (paragraphProperties.GetFirstChild<D.BulletColor>() != null)
            {
                D.BulletColor buColor = paragraphProperties.GetFirstChild<D.BulletColor>();
                
                if (buColor.GetFirstChild<D.SchemeColor>() != null)
                {
                    D.SchemeColor currentColor = buColor.GetFirstChild<D.SchemeColor>();
                    currentColor.Val = schemeColorVal;
                }
                else
                {
                    if (buColor.GetFirstChild<D.RgbColorModelHex>() != null)
                    {
                        D.RgbColorModelHex currentRgbColor = buColor.GetFirstChild<D.RgbColorModelHex>();
                        buColor.RemoveChild(currentRgbColor);
                    }

                    D.SchemeColor schemeColor = new D.SchemeColor() { Val = schemeColorVal };
                    buColor.AddChild(schemeColor);
                }
            }

            else
            {
                D.BulletColor buColor = new D.BulletColor();
                D.SchemeColor schemeColor = new D.SchemeColor() { Val = schemeColorVal };
                buColor.AddChild(schemeColor);
                paragraphProperties.AddChild(buColor);
            }

        }

        /// <summary>
        /// Sets line spacing to single spaced
        /// </summary>
        /// <param name="paragraphProperties"></param>

        public static void SetSingleSpacing(this D.ParagraphProperties paragraphProperties)
        {
            if (paragraphProperties.GetFirstChild<D.LineSpacing>() != null)
            {
                D.LineSpacing lineSpacing = paragraphProperties.GetFirstChild<D.LineSpacing>();
                if (lineSpacing.GetFirstChild<D.SpacingPercent>() != null)
                {
                    lineSpacing.AddChild(new D.SpacingPercent() { Val = 100000 });
                }

                else
                {
                    if (lineSpacing.GetFirstChild<D.SpacingPoints>() != null)
                    {
                        lineSpacing.GetFirstChild<D.SpacingPoints>().Remove();
                    }

                    lineSpacing.AddChild(new D.SpacingPercent() { Val = 100000 });

                }
            }

            D.LineSpacing newLineSpacing = new D.LineSpacing();
            newLineSpacing.AddChild(new D.SpacingPercent() { Val = 100000 });
            paragraphProperties.AddChild(newLineSpacing);
        }


    }
}
