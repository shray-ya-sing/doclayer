namespace InternalUtilities.ErrorHandling;

public class ThemeApplicationException : Exception
{
    public ThemeApplicationException() : base("An error occurred while applying the theme to the presentation.")
    {
    }

    public ThemeApplicationException(string message) : base(message)
    {
    }

    public ThemeApplicationException(string message, Exception innerException) : base(message, innerException)
    {
    }
}