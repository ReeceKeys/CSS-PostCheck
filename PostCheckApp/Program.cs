using System;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        // Command to open the Sound Settings in Windows
        string soundSettingsCommand = "ms-settings:sound";

        // Start the process to open the Sound settings
        Process.Start(new ProcessStartInfo
        {
            FileName = soundSettingsCommand,
            UseShellExecute = true
        });

        Console.WriteLine("Opening Sound Settings...");
    }
}
