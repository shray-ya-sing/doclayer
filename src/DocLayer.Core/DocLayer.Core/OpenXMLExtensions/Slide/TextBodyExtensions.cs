using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLExtensions
{
    public static class TextBodyExtensions
    {
        /// <summary>
        /// Sets the basic RightToLeft columns and Text Anchoring properties
        /// </summary>
        /// <param name="textBody"></param>
        
        public static void SetBasicBodyProperties(this P.TextBody textBody)
        {
            if (textBody.BodyProperties != null)
            {
                D.BodyProperties props = textBody.BodyProperties;
                props.RightToLeftColumns = true;
                props.Anchor = D.TextAnchoringTypeValues.Center;

            }

            else
            {
                D.BodyProperties props = new D.BodyProperties();
                props.RightToLeftColumns = true;
                props.Anchor = D.TextAnchoringTypeValues.Center;
                textBody.AddChild(props);
            }
            

        }

        /// <summary>
        /// Creates run with text, and paragraph and adds to textbody
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="text"></param>
        public static void AddParagraph(this P.TextBody textBody, string text)
        {
            D.Paragraph para = new D.Paragraph();
            para.SetAlignCenter();
            para.SetText(text);
            para.SetEndProps();

            textBody.AppendChild<D.Paragraph>(para);
        }

        /// <summary>
        /// Set shape autofit
        /// </summary>
        /// <param name="textBody"></param>
        public static void SetShapeAutofit(this P.TextBody textBody)
        {
            if (textBody.BodyProperties != null)
            {
                textBody.BodyProperties.SetShapeAutofit();
            }
            else
            {
                return; // don't do anything
            }
        }
        /// <summary>
        /// Sets no autofit
        /// </summary>
        /// <param name="textBody"></param>
        public static void SetNoAutofit(this P.TextBody textBody)
        {
            if (textBody.BodyProperties != null)
            {
                textBody.BodyProperties.SetNoAutofit();
            }
            else
            {
                return; // don't do anything
            }
        }
    }
}
