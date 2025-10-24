
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;

namespace OpenXMLExtensions
{
    public static class ThemeExtensions
    {
        /// <summary>
        /// Sets accents colors 1 to 4 for the presentation
        /// </summary>
        /// <param name="theme"></param>
        /// <param name="colorsHexcodes"></param>
        public static void SetAccentColors(this D.Theme theme, List<string> colorsHexcodes)
        {
            if (theme.GetFirstChild<D.ThemeElements>() != null)
            {
                D.ThemeElements themeElements = theme.GetFirstChild<D.ThemeElements>();

                if (themeElements.GetFirstChild<D.ColorScheme>() != null)
                {
                    D.ColorScheme colorScheme = themeElements.GetFirstChild<D.ColorScheme>();
                    if (colorScheme.GetFirstChild<D.Accent1Color>() != null)
                    {
                        if (colorScheme.GetFirstChild<D.Accent1Color>().GetFirstChild<D.RgbColorModelHex>() != null)
                        {
                            colorScheme.GetFirstChild<D.Accent1Color>().GetFirstChild<D.RgbColorModelHex>().Val = colorsHexcodes[0];
                        }                            
                    }

                    if (colorScheme.GetFirstChild<D.Accent2Color>() != null)
                    {
                        if (colorScheme.GetFirstChild<D.Accent2Color>().GetFirstChild<D.RgbColorModelHex>() != null)
                        {
                            colorScheme.GetFirstChild<D.Accent2Color>().GetFirstChild<D.RgbColorModelHex>().Val = colorsHexcodes[1];
                        }
                    }

                    if (colorScheme.GetFirstChild<D.Accent3Color>() != null)
                    {
                        if (colorScheme.GetFirstChild<D.Accent3Color>().GetFirstChild<D.RgbColorModelHex>() != null)
                        {
                            colorScheme.GetFirstChild<D.Accent3Color>().GetFirstChild<D.RgbColorModelHex>().Val = colorsHexcodes[2];
                        }
                    }

                    if (colorScheme.GetFirstChild<D.Accent4Color>() != null)
                    {
                        if (colorScheme.GetFirstChild<D.Accent4Color>().GetFirstChild<D.RgbColorModelHex>() != null)
                        {
                            colorScheme.GetFirstChild<D.Accent4Color>().GetFirstChild<D.RgbColorModelHex>().Val = colorsHexcodes[3];
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Adds a set of custom colors to the presentation, replacing the existing, if any
        /// </summary>
        /// <param name="theme"></param>
        /// <param name="colorsHexCodes"></param>
        public static void SetCustomColors(this D.Theme theme, Dictionary<string,string> colorsHexCodes)
        {
            if (theme.GetFirstChild<D.CustomColorList>() != null)
            {
                D.CustomColorList currentCustomColorList = theme.GetFirstChild<D.CustomColorList>();
                D.CustomColorList customColorList = new D.CustomColorList();
                foreach (string colorName in colorsHexCodes.Keys)
                {
                    customColorList.AppendChild(new D.CustomColor() { Name = colorName });
                    customColorList.Elements<D.CustomColor>().Last().AddChild(new D.RgbColorModelHex() { Val = colorsHexCodes[colorName] });
                }
                theme.ReplaceChild<D.CustomColorList>(customColorList, currentCustomColorList);
            }
            else
            {
                D.CustomColorList customColorList = new D.CustomColorList();
                foreach(string colorName in colorsHexCodes.Keys)
                {
                    customColorList.AppendChild(new D.CustomColor() { Name = colorName });
                    customColorList.Elements<D.CustomColor>().Last().AddChild(new D.RgbColorModelHex() { Val = colorsHexCodes[colorName] });
                }
                theme.AddChild(customColorList);
            }
        }

        /// <summary>
        /// Adds custom colors to the existing custom color list, or as new list if none exists
        /// </summary>
        /// <param name="theme"></param>
        /// <param name="colorsHexCodes"></param>

        public static void AddCustomColors(this D.Theme theme, Dictionary<string, string> colorsHexCodes)
        {
            if (theme.GetFirstChild<D.CustomColorList>() != null)
            {
                D.CustomColorList customColorList = theme.GetFirstChild<D.CustomColorList>();
                if (customColorList.GetFirstChild<D.CustomColor>() != null) 
                {
                    foreach (string colorName in colorsHexCodes.Keys)
                    {
                        customColorList.AppendChild(new D.CustomColor() { Name = colorName });
                        customColorList.Elements<D.CustomColor>().Last().AddChild(new D.RgbColorModelHex() { Val = colorsHexCodes[colorName] });
                    }
                }
                else
                {
                    foreach (string colorName in colorsHexCodes.Keys)
                    {
                        customColorList.AppendChild(new D.CustomColor() { Name = colorName });
                        customColorList.Elements<D.CustomColor>().Last().AddChild(new D.RgbColorModelHex() { Val = colorsHexCodes[colorName] });
                    }
                }             
            }
            else
            {
                D.CustomColorList customColorList = new D.CustomColorList();
                foreach (string colorName in colorsHexCodes.Keys)
                {
                    customColorList.AppendChild(new D.CustomColor() { Name = colorName });
                    customColorList.Elements<D.CustomColor>().Last().AddChild(new D.RgbColorModelHex() { Val = colorsHexCodes[colorName] });
                }
            }
        }


        /// <summary>
        /// Sets presentation font to Arial
        /// </summary>
        /// <param name="theme"></param>
        public static void SetFontSchemeArial(this D.Theme theme)
        {
            D.FontScheme arialFontScheme = new D.FontScheme() { Name = "Arial" };
            D.MajorFont arialMajorFont = new D.MajorFont();
            D.MinorFont arialMinorFont = new D.MinorFont();
            arialMajorFont.AddChild(new D.LatinFont() { Typeface = "Arial" });
            arialMajorFont.AddChild(new D.EastAsianFont() { Typeface = "" });
            arialMajorFont.AddChild(new D.ComplexScriptFont() { Typeface = "" });

            arialMinorFont.AddChild(new D.LatinFont() { Typeface = "Arial" });
            arialMinorFont.AddChild(new D.EastAsianFont() { Typeface = "" });
            arialMinorFont.AddChild(new D.ComplexScriptFont() { Typeface = "" });

            arialFontScheme.AddChild(arialMajorFont);
            arialFontScheme.AddChild(arialMinorFont);

            if (theme.GetFirstChild<D.ThemeElements>() != null)
            {
                D.ThemeElements themeElements = theme.GetFirstChild<D.ThemeElements>();              

                if (themeElements.GetFirstChild<D.FontScheme>() != null)
                {
                    D.FontScheme fontScheme = themeElements.GetFirstChild<D.FontScheme>();
                    themeElements.ReplaceChild<D.FontScheme>(arialFontScheme,fontScheme);
                }
                else
                {
                    themeElements.AddChild(arialFontScheme);
                }
            }                  
        }
        
        // TODO: Incorporate check against a list of valid font typefaces
        public static void SetCustomFontScheme(this D.Theme theme, string typefaceName)
        {           
            D.FontScheme customFontScheme = new D.FontScheme() { Name = typefaceName };

            D.MajorFont majorFont = new D.MajorFont();
            D.MinorFont minorFont = new D.MinorFont();
            majorFont.AddChild(new D.LatinFont() { Typeface = typefaceName });
            majorFont.AddChild(new D.EastAsianFont() { Typeface = "" });
            majorFont.AddChild(new D.ComplexScriptFont() { Typeface = "" });

            minorFont.AddChild(new D.LatinFont() { Typeface = typefaceName });
            minorFont.AddChild(new D.EastAsianFont() { Typeface = "" });
            minorFont.AddChild(new D.ComplexScriptFont() { Typeface = "" });

            customFontScheme.AddChild(majorFont);
            customFontScheme.AddChild(minorFont);

            if(theme.GetFirstChild<D.ThemeElements>() != null)
            {
                D.ThemeElements themeElements = theme.GetFirstChild<D.ThemeElements>();

                if (themeElements.GetFirstChild<D.FontScheme>() != null)
                {
                    D.FontScheme fontScheme = themeElements.GetFirstChild<D.FontScheme>();
                    themeElements.ReplaceChild<D.FontScheme>(customFontScheme, fontScheme);
                }
                else
                {
                    themeElements.AddChild(customFontScheme);
                }
            }
        }
    }
}
