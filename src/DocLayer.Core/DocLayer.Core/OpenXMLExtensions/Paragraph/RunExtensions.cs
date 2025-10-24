using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;

namespace OpenXMLExtensions
{
    public static class RunExtensions
    {
        /// <summary>
        /// Sets run text font size
        /// </summary>
        /// <param name="run"></param>
        public static void SetRunSize(this D.Run run, int size)
        {
            if (run.RunProperties != null)
            {
                D.RunProperties props = run.RunProperties;
                props.FontSize = size * 100;
                props.Dirty = true;
            }
            else
            {
                D.RunProperties props = new D.RunProperties();
                props.FontSize = size * 100;
                run.AddChild(props);

            }
            
        }

        /// <summary>
        /// Sets run text language to english
        /// </summary>
        /// <param name="run"></param>
        public static void SetRunEnglish(this D.Run run)
        {            
            if (run.RunProperties != null)
            {
                D.RunProperties props = run.RunProperties;
                props.Language = "en-US";
                props.Dirty = true;
            }
            else
            {
                D.RunProperties props = new D.RunProperties();
                props.Language = "en-US";
                run.AddChild(props);
            }

        }

        /// <summary>
        /// Sets run text bold
        /// </summary>
        /// <param name="run"></param>
        public static void SetRunBold(this D.Run run)
        {
            if (run.RunProperties != null)
            {
                D.RunProperties props = run.RunProperties;
                props.Bold = true;
                props.Dirty = true;                
            }
            else
            {
                D.RunProperties props = new D.RunProperties();
                props.Bold = true;
                run.AddChild(props);
            }

        }

        /// <summary>
        /// Sets run text italic
        /// </summary>
        /// <param name="run"></param>
        public static void SetRunItalic(this D.Run run)
        {
            if (run.RunProperties != null)
            {
                D.RunProperties props = run.RunProperties;
                props.Italic = true;
                props.Dirty = true;
            }
            else
            {
                D.RunProperties props = new D.RunProperties();
                props.Italic = true;
                run.AddChild(props);
            }

        }
        /// <summary>
        /// Sets the color of run text according to color scheme / theme
        /// </summary>
        /// <param name="run"></param>
        /// <param name="accentNum"></param>
        public static void SetRunSchemeFill(this D.Run run, int accentNum)
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

            if (run.RunProperties != null)
            {
                D.RunProperties props = run.RunProperties;
                if (props.GetFirstChild<D.SolidFill>() is not null)
                {
                    D.SolidFill solidFill = props.GetFirstChild<D.SolidFill>();
                    D.SolidFill newSolidFill = new D.SolidFill(
                       new SchemeColor() { Val = schemeColorVal }
                       );

                    props.ReplaceChild<SolidFill>(newSolidFill, solidFill);
                    props.Dirty = true;
                }
                else
                {
                    D.SolidFill solidFill = new D.SolidFill(
                       new SchemeColor() { Val = schemeColorVal }
                       );
                    props.AddChild(solidFill);
                    props.Dirty = true;
                }
            }

            else
            {
                D.RunProperties props = new D.RunProperties();
                D.SolidFill solidFill = new D.SolidFill(
                       new SchemeColor() { Val = schemeColorVal }
                       );

                // Add in correct spot
                props.AddChild(solidFill);
                run.AddChild(props);

            }
        }

        public static void SetRunHexFill(this D.Run run, string rgbColorHex)
        {
            if (run.RunProperties != null)
            {
                D.RunProperties props = run.RunProperties;
                if (props.GetFirstChild<D.SolidFill>() != null)
                {
                    D.SolidFill solidFill = props.GetFirstChild<D.SolidFill>();
                    D.SolidFill newSolidFill = new D.SolidFill(
                       new RgbColorModelHex() { Val = rgbColorHex }
                       );

                    props.ReplaceChild<SolidFill>(newSolidFill, solidFill);
                    props.Dirty = true;
                }
                else
                {
                    D.SolidFill newSolidFill = new D.SolidFill(
                       new RgbColorModelHex() { Val = rgbColorHex }
                       );
                    props.AddChild(newSolidFill);
                    props.Dirty = true;
                }
            }

            else
            {
                D.RunProperties props = new D.RunProperties();
                D.SolidFill solidFill = new D.SolidFill(
               new RgbColorModelHex() { Val = rgbColorHex }
               );
                props.AddChild(solidFill);
                run.AddChild(props);
            }
        }


        /// <summary>
        /// Sets text string to run
        /// </summary>
        /// <param name="run"></param>
        public static void AddText(this D.Run run, string text)
        {
            if (run.GetFirstChild<D.Text>() != null)
            {
                // There is already text in the run: replace with new 
                foreach(D.Text t in run.Elements<D.Text>())
                {
                    t.Remove();
                }
                
                D.Text runText = new D.Text(text);
                run.AddChild(runText);
            }
            else
            {
                D.Text runText = new D.Text(text);
                run.AddChild(runText);
            }
            
        }

        public static D.Text GetText(this D.Run run)
        {
            if (run.GetFirstChild<D.Text>() != null)
            {
                return run.GetFirstChild<D.Text>();
            }
            else throw new Exception("No text found");
        }

        public static string GetInnerText(this D.Run run)
        {
            string text = "";
            if (run.GetFirstChild <D.Text>() != null)
            {
                foreach (D.Text t in run.Elements<D.Text>())
                {
                    text += t.InnerText.ToString();
                }

                return text;
            }

            return text;
        }
    }
}
