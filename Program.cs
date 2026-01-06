// © Ricki Angel 2026 | TechAngelX
// All rights reserved.
// Program.cs - the e C# program entry point that initialises and starts the Avalonia UI application.

using Avalonia;
using System;

namespace ADMerger;

class Program
{
    [STAThread]
    public static void Main(string[] args) => BuildAvaloniaApp()
        .StartWithClassicDesktopLifetime(args);

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
