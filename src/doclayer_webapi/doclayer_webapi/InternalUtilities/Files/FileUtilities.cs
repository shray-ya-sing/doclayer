namespace InternalUtilities.Files
{
    public static class FileUtilities
    {
        /// <summary>
        /// Creates a temporary filepath in local memory
        /// </summary>
        /// <param name="extension"> Should not include the "." prefix </param>
        /// <returns></returns>
        public static string CreateTempFilepath(string extension) 
        {
            // Create Local disk filepath
            var trustedFileName = Path.GetTempFileName();
            var trustedFilePath = Path.ChangeExtension(trustedFileName, extension);
            return trustedFilePath;
        }

        /// <summary>
        /// Downloads image from a web url to local disk
        /// </summary>
        /// <param name="directoryPath"></param>
        /// <param name="fileName"></param>
        /// <param name="uri"></param>
        /// <returns> Filepath of the downloaded image </returns>
        public static async Task<string> DownloadImageAsync(string fileName, Uri imageUri)
        {
            using var httpClient = new HttpClient();

            // Get the file extension
            var uriWithoutQuery = imageUri.GetLeftPart(UriPartial.Path);
            var fileExtension = Path.GetExtension(uriWithoutQuery);

            // Create file path and ensure directory exists
            string path = FileUtilities.CreateTempFilepath(fileExtension);

            try
            {
                // Download the image and write to the file
                var imageBytes = await httpClient.GetByteArrayAsync(imageUri);
                await File.WriteAllBytesAsync(path, imageBytes);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            httpClient.Dispose();
            return path;
        }

        /// <summary>
        /// Gets image stream from a web url
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="imageUri"></param>
        /// <returns> Image byte stream </returns>
        public static async Task<Stream> GetImageStreamAsync(string fileName, Uri imageUri)
        {
            try
            {
                string filePath = await DownloadImageAsync(fileName, imageUri);
                return File.OpenRead(filePath);
            }
            catch (Exception ex)
            {                
                throw new Exception(ex.Message);
            }

        }

        public static byte[] ReadFully(Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
