using DocLayer.Core.Examples;

Console.WriteLine("Testing Slide Editing Features");
Console.WriteLine("=========================================\n");

try
{
    TestSlideEditing.Run();
    
    Console.WriteLine("\n" + "=".PadRight(50, '='));
    Console.WriteLine("✓ Tests completed successfully!");
}
catch (Exception ex)
{
    Console.WriteLine($"\n✗ Error: {ex.Message}");
    Console.WriteLine($"Stack trace: {ex.StackTrace}");
    return 1;
}

return 0;
