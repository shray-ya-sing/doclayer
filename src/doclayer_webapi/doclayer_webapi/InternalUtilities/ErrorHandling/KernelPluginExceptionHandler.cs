namespace InternalUtilities.ErrorHandling;

public static class KernelPluginExceptionHandler
{
    public class OperationResult<T>
    {
        /// <summary>
        /// Indicates whether the function call was succesful
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Message { get; set; }
        public T Data { get; set; }
        public Exception Error { get; set; }
    }

    public static async Task<T> HandleTaskExceptionAsync<T>(Func<Task<T>> func, ILogger logger, string pluginName, string functionName)
    {
        try
        {
            return await func();
        }
        catch (Exception ex)
        {
            logger.LogError("Error in plugin {PluginName}, function {FunctionName}: {Message}\nStack Trace: {StackTrace}", pluginName, functionName, ex.Message, ex.StackTrace);
            throw new Exception($"An error occurred in {pluginName}.{functionName}: {ex.Message}");
        }
    }

    public static T HandleException<T>(Func<T> func, ILogger logger, string pluginName, string functionName)
    {
        try
        {
            return func();
        }
        catch (Exception ex)
        {
            logger.LogError("Error in plugin {PluginName}, function {FunctionName}: {Message}\nStack Trace: {StackTrace}", pluginName, functionName, ex.Message, ex.StackTrace);
            throw new Exception($"An error occurred in {pluginName}.{functionName}: {ex.Message}");
        }
    }

    public static void HandleException(Action action, ILogger logger, string pluginName, string functionName)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            logger.LogError("Error in plugin {PluginName}, function {FunctionName}: {Message}\nStack Trace: {StackTrace}", pluginName, functionName, ex.Message, ex.StackTrace);
            throw new Exception($"An error occurred in {pluginName}.{functionName}: {ex.Message}");
        }
    }

    public static async Task<OperationResult<T>> HandleExceptionAsync<T>(Func<Task<T>> func, ILogger logger, string pluginName, string functionName)
    {
        try
        {
            var result = await func();
            return new OperationResult<T>
            {
                IsSuccess = true,
                Message = "Operation completed successfully.",
                Data = result
            };
        }
        catch (Exception ex)
        {
            logger.LogError("Error in plugin {PluginName}, function {FunctionName}: {Message}\nStack Trace: {StackTrace}", pluginName, functionName, ex.Message, ex.StackTrace);
            return new OperationResult<T>
            {
                IsSuccess = false,
                Message = $"An error occurred in {pluginName}.{functionName}: {ex.Message}",
                Error = ex
            };
        }
    }

    public static OperationResult<T> HandleOperationResultException<T>(Func<T> func, ILogger logger, string pluginName, string functionName)
    {
        try
        {
            var result = func();
            return new OperationResult<T>
            {
                IsSuccess = true,
                Message = "Operation completed successfully.",
                Data = result
            };
        }
        catch (Exception ex)
        {
            logger.LogError("Error in plugin {PluginName}, function {FunctionName}: {Message}\nStack Trace: {StackTrace}", pluginName, functionName, ex.Message, ex.StackTrace);
            return new OperationResult<T>
            {
                IsSuccess = false,
                Message = $"An error occurred in {pluginName}.{functionName}: {ex.Message}",
                Error = ex
            };
        }
    }

    public static OperationResult<bool> HandleOperationResultException(Action action, ILogger logger, string pluginName, string functionName)
    {
        try
        {
            action();
            return new OperationResult<bool>
            {
                IsSuccess = true,
                Message = "Operation completed successfully.",
                Data = true
            };
        }
        catch (Exception ex)
        {
            logger.LogError("Error in plugin {PluginName}, function {FunctionName}: {Message}\nStack Trace: {StackTrace}", pluginName, functionName, ex.Message, ex.StackTrace);
            return new OperationResult<bool>
            {
                IsSuccess = false,
                Message = $"An error occurred in {pluginName}.{functionName}: {ex.Message}",
                Error = ex,
                Data = false
            };
        }
    }
}
