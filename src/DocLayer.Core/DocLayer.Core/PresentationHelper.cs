using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OpenXMLExtensions;

namespace DocLayer.Core
{
    /// <summary>
    /// Helper methods for creating and managing PowerPoint presentations
    /// </summary>
    public static class PresentationHelper
    {
        /// <summary>
        /// Creates a new PowerPoint presentation file with basic structure
        /// </summary>
        /// <param name="filepath">Path where the presentation will be created</param>
        /// <param name="widescreen">If true, uses 16:9 format; otherwise uses 4:3</param>
        /// <returns>PresentationDocument instance ready for use</returns>
        public static PresentationDocument CreatePresentation(string filepath, bool widescreen = true)
        {
            PresentationDocument presentationDoc = PresentationHelperMethods.CreatePresentation(filepath);
            if (widescreen) {
                if (presentationDoc.PresentationPart!.Presentation is not null){
                    presentationDoc.PresentationPart!.Presentation.SetSlideSizeWidescreen();
                }
            }
            return presentationDoc;
        }

    }
}
