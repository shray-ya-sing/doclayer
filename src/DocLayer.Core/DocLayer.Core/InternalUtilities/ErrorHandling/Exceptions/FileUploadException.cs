namespace InternalUtilities.ErrorHandling;

public class FileUploadException : Exception
{
    public FileUploadException() : base("An error occurred while uploading the file.")
    {
    }

    public FileUploadException(string message) : base(message)
    {
    }

    public FileUploadException(string message, Exception innerException) : base(message, innerException)
    {
    }
}