using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using InternalUtilities.Files;
using Syncfusion;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Interactive;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO.Packaging;

namespace InternalUtilities.Syncfusion
{
    public static class SyncfusionHelperMethods
    {
        public static void CopySlideToPresentation(string destinationFilePath, string sourceFilePath, int sourceSlideNumber)
        {
            var slide = CloneSlideFromPresentation(sourceFilePath, sourceSlideNumber);
            AddClonedSlide(destinationFilePath, slide);
        }
        public static ISlide CloneSlideFromPresentation( string sourceFilePath, int sourceSlideNumber)
        {
            
            using (FileStream fileStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //sourceSlideNumber starts from 0 index
                    return pptxDoc.Slides[sourceSlideNumber].Clone();
                }
            }
                      
        }

        public static void AddClonedSlide(string destinationFilePath, ISlide clonedSlide) 
        {
            if (clonedSlide != null)
            {
                using (FileStream fileStream = new FileStream(destinationFilePath, FileMode.Open, FileAccess.Read))
                {
                    //Open the existing PowerPoint presentation.
                    using (IPresentation pptxDoc = Presentation.Open(fileStream))
                    {
                        //sourceSlideNumber starts from 0 index
                        pptxDoc.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme);
                    }
                }

            }
            
        }

        /// <summary>
        /// Exports a slide to an image
        /// </summary>
        /// <param name="path"></param>
        /// <param name="slideNumber"></param>
        /// <returns>String local path to the exported slide image</returns>
        public static string ExportSlideToImage(string path, int slideNumber)
        {
            string imgFileName = Path.GetTempFileName();
            string imgFilePath = Path.ChangeExtension(imgFileName, ".jpg");
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();

                    //Convert PowerPoint slide to image as stream.
                    using (Stream stream = pptxDoc.Slides[slideNumber - 1].ConvertToImage(ExportImageFormat.Jpeg))
                    {
                        //Reset the stream position
                        stream.Position = 0;
                        //Create the output image file stream
                        using (FileStream fileStreamOutput = File.Create(imgFilePath))
                        {
                            //Copy the converted image stream into created output stream
                            stream.CopyTo(fileStreamOutput);
                        }
                    }

                    return imgFilePath;
                }
            }

        }

        /// <summary>
        /// Exports pptx file to images
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns>List of local filepaths of the images</returns>
        public static List<string> ExportPptToImages(string filepath) 
        {
            List<string> imgFilePaths = new List<string>();
            using (FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint to image as stream.
                    Stream[] images = pptxDoc.RenderAsImages(ExportImageFormat.Jpeg);
                    int index = 0;
                    //Saves the images to file system
                    foreach (Stream stream in images)
                    {
                        index++;
                        string imgFileName = Path.GetTempFileName();
                        string imgFilePath = Path.ChangeExtension(imgFileName, ".jpg");
                        imgFilePaths.Add(imgFilePath);
                        //Create the output image file stream
                        using (FileStream fileStreamOutput = File.Create(imgFilePath))
                        {
                            //Copy the converted image stream into created output stream
                            stream.CopyTo(fileStreamOutput);
                        }
                    }                    
                }
            }
            return imgFilePaths;
        }

        /// <summary>
        /// Exports a pptx file to a pdf file
        /// Returns the path of the pdf file created
        /// </summary>
        /// <param name="pptfilepath"></param>
        /// <returns></returns>
        public static string ExportPptToPdf(string pptfilepath) 
        {
            string pdffilepath = FileUtilities.CreateTempFilepath("pdf");
            //Load the PowerPoint presentation into stream.
            using (FileStream fileStreamInput = new FileStream(pptfilepath, FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation with loaded stream.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
                {
                    //Convert PowerPoint into PDF document. 
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                    {
                        //Save the PDF file to file system. 
                        using (FileStream outputStream = new FileStream(pdffilepath, FileMode.Create, FileAccess.ReadWrite))
                        {
                            pdfDocument.Save(outputStream);
                        }
                    }
                }
            }

            return pdffilepath;
        }
    }
}
