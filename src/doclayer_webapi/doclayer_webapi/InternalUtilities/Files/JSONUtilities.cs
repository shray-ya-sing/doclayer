using System.Text.RegularExpressions;

namespace InternalUtilities.Files;

public static class JSONUtilities
{
    /// <summary>
    /// Extracts a JSON object from a string.
    /// </summary>
    /// <param name="message"></param>
    /// <param name="sectionName"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public static string ExtractJsonObject(string message, string sectionName)
    {
        var pattern = $@"\""{sectionName}\"":\s*\{{(?:[^{{}}]|(?<open>\{{)|(?<-open>\}}))+(?(open)(?!))\}}";
        var match = Regex.Match(message, pattern);

        if (match.Success)
        {
            // Include the enclosing braces
            int startIndex = match.Index;
            int length = match.Length;
            return message.Substring(startIndex, length);
        }
        else
        {
            throw new InvalidOperationException($"No JSON object found for section '{sectionName}' in the message.");
        }
    }

    /// <summary>
    /// Extracts JSON array from a string response from the PerplexityAPI.
    /// </summary>
    /// <param name="message"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    public static string ExtractJsonArray(string message)
    {
        var pattern = @"\[\s*\{[\s\S]*?\}\s*\]";
        var match = Regex.Match(message, pattern);

        if (match.Success)
        {
            return match.Value;
        }
        else
        {
            throw new InvalidOperationException("No JSON array found in the message.");
        }
    }
}
