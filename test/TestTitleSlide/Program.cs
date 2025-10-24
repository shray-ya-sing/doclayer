using DocLayer.Core.Examples;

Console.WriteLine("Testing DocLayer.Core");
Console.WriteLine("=========================================\n");

try
{
    // Test 1: Title Layout Slide
    Console.WriteLine("[Test 1] Title Layout Slide");
    Console.WriteLine(new string('-', 40));
    TestTitleLayoutSlide.Run();
    Console.WriteLine();

    // Test 2: Custom Theme
    Console.WriteLine("[Test 2] Custom Theme");
    Console.WriteLine(new string('-', 40));
    TestTheme.Run();
    Console.WriteLine();

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
