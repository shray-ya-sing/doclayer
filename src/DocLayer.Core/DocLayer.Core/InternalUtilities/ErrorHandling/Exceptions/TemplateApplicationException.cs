namespace InternalUtilities.ErrorHandling;

public class TemplateApplicationException : Exception
{
    public TemplateApplicationException() : base("An error occurred while applying the template to the presentation.")
    {
    }

    public TemplateApplicationException(string message) : base(message)
    {
    }

    public TemplateApplicationException(string message, Exception innerException) : base(message, innerException)
    {
    }
}