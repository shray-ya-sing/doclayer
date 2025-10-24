using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using Syncfusion.Presentation;

namespace OpenXMLExtensions
{
    public static class PresentationExtensions
    {
        /// <summary>
        /// Sets slide to standard 4x3
        /// </summary>
        /// <param name="presentation"></param>
        public static void SetSlideSizeStandard(this P.Presentation presentation)
        {
            if(presentation.GetFirstChild<P.SlideSize>() != null)
            {
                P.SlideSize oldSlideSize = presentation.GetFirstChild<P.SlideSize>();
                P.SlideSize slideSize = new P.SlideSize();
                slideSize.Cx = 9144000;
                slideSize.Cy = 6858000;
                slideSize.Type = SlideSizeValues.Screen4x3; ;
                slideSize.Type = P.SlideSizeValues.Screen4x3;
                presentation.ReplaceChild<P.SlideSize>(slideSize, oldSlideSize);                
            }
        }

        /// <summary>
        /// Sets slide size to 16x9
        /// </summary>
        /// <param name="presentation"></param>
        public static void SetSlideSizeWidescreen(this P.Presentation presentation)
        {
            if (presentation.GetFirstChild<P.SlideSize>() != null)
            {
                P.SlideSize oldSlideSize = presentation.GetFirstChild<P.SlideSize>();
                P.SlideSize slideSize = new P.SlideSize();
                slideSize.Cx = 12192000;
                slideSize.Cy = 6858000;
                slideSize.Type = SlideSizeValues.Screen16x9;
                presentation.ReplaceChild<P.SlideSize>(slideSize, oldSlideSize);
            }

        }

        /// <summary>
        /// Get slide size in EMU
        /// </summary>
        /// <param name="presentation"></param>
        public static (Int32Value?, Int32Value?) GetSlideSize(this P.Presentation presentation)
        {
            if (presentation.GetFirstChild<P.SlideSize>() != null)
            {
                P.SlideSize slideSize = presentation.GetFirstChild<P.SlideSize>();
                if (slideSize.Cx is not null && slideSize.Cy is not null)
                {
                    return (slideSize.Cx, slideSize.Cy);
                }
                else throw new Exception("ERROR: No slide size found");
            }
            else throw new Exception("ERROR: No slide size found");
        }

        /// <summary>
        /// Returns number of slides in the presentation
        /// </summary>
        /// <param name="presentation"></param>
        /// <returns></returns>
        public static int CountSlides(this P.Presentation presentation)
        {
            if (presentation.GetFirstChild<SlideIdList>() != null)
            {
                int count = presentation.SlideIdList.Count();
                return count;
            }
            else
            {
                return 0;
            }
        }
        
        /// <summary>
        /// Returns the slide relationship id
        /// </summary>
        /// <param name="presentation"></param>
        /// <param name="slideNumber"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static string GetSlideRelId(this P.Presentation presentation, int slideNumber)
        {
            int slideCount = presentation.CountSlides();

            string relid;
            
            if (slideNumber <= slideCount && presentation.GetFirstChild<SlideIdList>() != null)
            {
                SlideId slideId = presentation.SlideIdList.Elements<SlideId>().ElementAt(slideNumber - 1);
                relid = slideId.RelationshipId;
                return relid;
                
            }
            else
            {
                throw new Exception("Slide number is not valid");
            }          
        }
    }
}
