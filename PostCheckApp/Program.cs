using System;
using System.Diagnostics;
using System.Management;
using NAudio.CoreAudioApi;

class Program
{
    static void Main()
    {
        // Open Display Settings
        Console.WriteLine(new string('-', 40) + '\n');
        Console.WriteLine("Starting script...\n");
        Console.WriteLine(new string('-', 40));
        System.Threading.Thread.Sleep(3000);

        OpenWindowsSetting("ms-settings:display");
        Console.WriteLine("Opened DISPLAY settings.");
        Console.WriteLine(new string('-', 40));
        System.Threading.Thread.Sleep(2000);


        // Open Sound Settings
        OpenWindowsSetting("ms-settings:sound");
        Console.WriteLine("Opened SOUND settings.");
        Console.WriteLine(new string('-', 40));
        System.Threading.Thread.Sleep(2000);

        // Retrieve and display audio devices
        Console.WriteLine("Sound DEVICES list:\n");
        GetAudioDevices();
        Console.WriteLine(new string('-', 40));
        System.Threading.Thread.Sleep(2000);

        // Get the currently selected speaker
        string defaultDeviceName = GetDefaultAudioOutputDeviceName();
        Console.WriteLine("Default Audio Output Device: " + defaultDeviceName);
        System.Threading.Thread.Sleep(2000);

        // Open zoom in the default browser
        Process.Start(new ProcessStartInfo
        {
            FileName = "https://zoom.us/signin#/login",
            UseShellExecute = true
        });
        Console.WriteLine(new string('-', 40));
        Console.WriteLine("Opened ZOOM login.");
        Console.WriteLine(new string('-', 40));
    }

    static void GetAudioDevices()
    {
        // Query WMI for audio output devices
        using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SoundDevice"))
        {
            foreach (ManagementObject device in searcher.Get())
            {
                Console.WriteLine(" - " + device["Name"]);
            }
        }
    }

    static string GetDefaultAudioOutputDeviceName()
    {
        var enumerator = new MMDeviceEnumerator();
        var defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

        // Return the name of the default audio output device
        return defaultDevice.FriendlyName;
    }

    static void OpenWindowsSetting(string settingCommand)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = settingCommand,
            UseShellExecute = true
        });
    }
}
