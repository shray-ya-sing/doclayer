namespace InternalUtilities.ErrorHandling;

public class PresentationProcessingException : Exception
{
    public PresentationProcessingException() : base("An error occurred while processing the presentation.")
    {
    }

    public PresentationProcessingException(string message) : base(message)
    {
    }

    public PresentationProcessingException(string message, Exception innerException) : base(message, innerException)
    {
    }
}