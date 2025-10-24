using DocLayer.Core.Examples;

Console.WriteLine("Testing DocLayer.Core - Title Layout Slide");
Console.WriteLine("=========================================\n");

try
{
    TestTitleLayoutSlide.Run();
}
catch (Exception ex)
{
    Console.WriteLine($"✗ Error: {ex.Message}");
    Console.WriteLine($"Stack trace: {ex.StackTrace}");
    return 1;
}

return 0;
