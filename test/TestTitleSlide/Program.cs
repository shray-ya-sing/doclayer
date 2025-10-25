using DocLayer.Core.Examples;

Console.WriteLine("Testing Convenience Methods for Agent Integration");
Console.WriteLine("=========================================\n");

try
{
    TestConvenienceMethods.Run();
    
    Console.WriteLine("\n" + "=".PadRight(50, '='));
    Console.WriteLine("✓ All tests completed successfully!");
}
catch (Exception ex)
{
    Console.WriteLine($"\n✗ Error: {ex.Message}");
    Console.WriteLine($"Stack trace: {ex.StackTrace}");
    return 1;
}

return 0;
