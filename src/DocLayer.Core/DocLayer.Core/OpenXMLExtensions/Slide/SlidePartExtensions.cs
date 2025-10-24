using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using InternalUtilities.Files;

namespace OpenXMLExtensions
{
    public static class SlidePartExtensions
    {

        /// <summary>
        /// Adds an image to the image part and returns the relId of the image
        /// </summary>
        /// <param name="presDoc"></param>
        /// <returns> RelId of the imagePart </returns>
        public static string AddImagePartFromStream(this SlidePart slidePart, Stream imageStream)
        {
            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);

            using (imageStream)
            {
                imagePart.FeedData(imageStream);
            }

            string relId = slidePart.GetIdOfPart(imagePart);
            return relId;
        }

        // TODO: add support for other extensions besides png
        public static async Task<string> AddImagePartFromUri(this SlidePart slidePart, Uri uri, string fileName)
        {
            
            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
            string path = await FileUtilities.DownloadImageAsync(fileName, uri);
           
            using (Stream imageStream = File.OpenRead(path))
            {
                imagePart.FeedData(imageStream);
            }

            string relId = slidePart.GetIdOfPart(imagePart);
            return relId;
        }

        // TODO: add support for other extensions besides png
        public static async Task<string> AddImagePartFromLocalPath(this SlidePart slidePart, string filePath)
        {

            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);

            using (Stream imageStream = File.OpenRead(filePath))
            {
                imagePart.FeedData(imageStream);
            }

            string relId = slidePart.GetIdOfPart(imagePart);
            return relId;
        }
    }
}

