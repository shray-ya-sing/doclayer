using System;

namespace InternalUtilities.ErrorHandling;

/// <summary>
/// Base class for all exceptions that arise from the user entering an invalid or malicious input into a UI component that interacts with the server.
/// </summary>
public class UserInputException : Exception
{
    /// <summary>
    /// A user friendly message to display to the user on the UI
    /// </summary>
    private readonly string UserFriendlyMessage = "That input is invalid. Please try again.";
    public UserInputException() : base("User input is invalid.")
    {
    }

    public UserInputException(string message) : base(message)
    {
    }

    public UserInputException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
