using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;


namespace OpenXMLExtensions
{
    public static class PresentationDocExtensions
    {
        /// <summary>
        /// Gets the Slide object at the specified slide number from the pptx package
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="slideNumber"></param>
        /// <returns></returns>
        public static Slide GetSlide(this PresentationDocument presentationDocument, int slideNumber)
        {
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;
            string relid = presentation.GetSlideRelId(slideNumber);        
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relid);
            Slide slide = slidePart.Slide;
            return slide;
        }

        public static Slide GetLastSlide(this PresentationDocument presentationDocument)
        {
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;
            SlideId slideId = presentation.SlideIdList.Elements<SlideId>().Last();
            string relid = slideId.RelationshipId;
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relid);
            Slide slide = slidePart.Slide;
            return slide;
        }


        /// <summary>
        /// Gets the SlidePart object at the specified slide number from the pptx package
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="slideNumber"></param>
        /// <returns></returns>

        public static SlidePart GetSlidePart(this PresentationDocument presentationDocument, int slideNumber)
        {
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;
            string relid = presentation.GetSlideRelId(slideNumber);
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relid);
            return slidePart;
        }


        public static int GetSlideCount(this PresentationDocument presentationDocument)
        {
            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart? presentationPart = presentationDocument.PresentationPart;

            // Get the slide count from the SlideParts.
            if (presentationPart is not null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }

            // Return the slide count to the previous method.
            return slidesCount;
        }
        public static void WritePresDocInfo(this PresentationDocument presDoc)
        {
            WritePresPackageProps(presDoc);
            WritePresentationInfo(presDoc);
            WriteSlideMasterInfo(presDoc);
            //      GetSlidesInfo(presDoc);

        }

        public static string GetThemeInfo(this PresentationDocument presDoc)
        {
            ThemePart themePart = presDoc.PresentationPart.SlideMasterParts.First().ThemePart;
            Theme theme = themePart.Theme;
            ThemeElements e = theme.ThemeElements;
            string info = e.InnerXml;
            return info;
        }

        public static void PasteThemeInfo(this PresentationDocument presDoc, string info)
        {
            ThemeElements newtheme = new();
            newtheme.InnerXml = info;
            ThemePart themePart = presDoc.PresentationPart.SlideMasterParts.First().ThemePart;
            Theme theme = themePart.Theme;
            ThemeElements old = theme.ThemeElements;
            theme.ReplaceChild<ThemeElements>(newtheme, old);
        }
        internal static void WritePresPackageProps(PresentationDocument presDoc)
        {
            Console.WriteLine(presDoc.PackageProperties);
        }

        public static string GetPresentationInfo(PresentationDocument presDoc)
        {
            string info = "";
            Presentation pres = presDoc.PresentationPart.Presentation;
            info += "Presentation Properties:";
            foreach (OpenXmlElement child in pres.ChildElements)
            {
                info+= ($"\r\n Element Name : {child.XName}");
            }

            info += $"\r\n Number of Slides: {pres.SlideIdList.Elements<SlideId>().Count()}";
            info+= $"\r\n Slide Size: x: {pres.SlideSize.Type.Value.ToString}";
            return info;
        }

        internal static void WritePresentationInfo(PresentationDocument presDoc)
        {
            Presentation pres = presDoc.PresentationPart.Presentation;
            Console.WriteLine("----------\r\n Presentation Properties:");

            Console.WriteLine("Child Elements:");
            foreach (OpenXmlElement child in pres.ChildElements)
            {
                Console.WriteLine($"Element Name : {child.XName}");
            }

            Console.WriteLine($"\r\n Number of Slides: {pres.SlideIdList.Elements<SlideId>().Count()}" +
                $"\r\n Slide Size: x: {pres.SlideSize.Cx},y: {pres.SlideSize.Cy}");

        }

        public static string GetSlideMasterInfo(PresentationDocument presDoc)
        {
            string info = "";
            ThemePart themePart = presDoc.PresentationPart.SlideMasterParts.First().ThemePart;

            Theme theme = themePart.Theme;

            info += $"------\r\n Theme: \r\n Theme Id: {theme.ThemeId} " +
                $"\r\n Theme Name: {theme.Name} ";


            return info;
        }
        internal static void WriteSlideMasterInfo(PresentationDocument presDoc)
        {
            ThemePart themePart = presDoc.PresentationPart.SlideMasterParts.First().ThemePart;

            Theme theme = themePart.Theme;

            Console.WriteLine($"------\r\n Theme: \r\n Theme Id: {theme.ThemeId} " +
                $"\r\n Theme Name: {theme.Name} ");

            foreach (OpenXmlElement child in theme.ThemeElements)
            {
                Console.WriteLine($"Theme Elements: " +
                    $"\r\n Color Scheme: {theme.ThemeElements.ColorScheme.Name}" +
                    $"\r\n Font Scheme: {theme.ThemeElements.FontScheme.Name}" +
                    $"\r\n Format Scheme: {theme.ThemeElements.FormatScheme.Name}");
            }

            Console.WriteLine($"-------------Font Scheme: {theme.ThemeElements.FontScheme.InnerXml} \r\n ---------");
            Console.WriteLine($"-------------Format Scheme: {theme.ThemeElements.FormatScheme.InnerXml} \r\n ---------");


        }
    }
}
