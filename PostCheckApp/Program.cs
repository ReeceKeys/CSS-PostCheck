using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;
using NAudio.CoreAudioApi;
using System.Runtime.InteropServices;


class Program
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    static void Main()
    {
        var inputSimulator = new InputSimulator();
        Console.Write("Username: ");
        string? username = Console.ReadLine()?.Trim();

        Console.Write("Password: ");
        string? password = ReadPassword();
        Console.Write("\n");        

        if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
        {
            Console.WriteLine("Username or password cannot be empty. Exiting...");
            return;
        }

        Console.WriteLine(new string('-', 40));
        Console.WriteLine("\nStarting script...\n");
        Console.WriteLine("[DO NOT TOUCH KEYBOARD/MOUSE]\n");
        Console.WriteLine(new string('-', 40));

        // Open system settings
        OpenWindowsSetting("ms-settings:display", "DISPLAY");
        OpenWindowsSetting("ms-settings:sound", "SOUND");

        // Retrieve and display audio devices
        DisplayAudioDevices();

        // Start and focus Zoom
        StartAndFocusZoom();
        Thread.Sleep(3000);
        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        inputSimulator.Keyboard.KeyDown(VirtualKeyCode.RETURN);
        Thread.Sleep(3000);
        // Simulate login actions
        SimulateLogin(username, password);

        Console.WriteLine("Script completed successfully.");
    }

    static string ReadPassword()
    {
        string password = "";
        ConsoleKeyInfo key;

        do
        {
            key = Console.ReadKey(true); // 'true' prevents character from showing

            if (key.Key == ConsoleKey.Backspace && password.Length > 0)
            {
                // Handle backspace (remove last character)
                password = password.Substring(0, password.Length - 1);
                Console.Write("\b \b"); // Erase the last '*' from console
            }
            else if (!char.IsControl(key.KeyChar))
            {
                password += key.KeyChar;
                Console.Write("*"); // Mask character
            }
        } while (key.Key != ConsoleKey.Enter);

        return password;
    }

    static void OpenWindowsSetting(string settingCommand, string settingName)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = settingCommand,
                UseShellExecute = true
            });
            Console.WriteLine($"Opened {settingName} settings.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to open {settingName} settings: {ex.Message}");
        }
        Console.WriteLine(new string('-', 40));
        Thread.Sleep(2000);
    }

    static void DisplayAudioDevices()
    {
        Console.WriteLine("Sound DEVICES list:\n");

        using (var soundsearcher = new ManagementObjectSearcher("SELECT * FROM Win32_SoundDevice"))
        {
            foreach (ManagementObject device in soundsearcher.Get())
            {
                Console.WriteLine(" - " + device["Name"]);
            }
        }

        Console.WriteLine("Microphone DEVICES list:\n");

        using (var micsearcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Description LIKE '%Microphone%' OR Name LIKE '%Microphone%'"))
        {
            foreach (ManagementObject device in micsearcher.Get())
            {
                Console.WriteLine(" - " + device["Name"]);
            }
        }
        Console.WriteLine(new string('-', 40));
        var enumerator = new MMDeviceEnumerator();
        var defaultSpeaker = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        var defaultMicrophone = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);

        if (defaultSpeaker == null)
            Console.WriteLine("No default audio output device found.");
        else 
            Console.WriteLine($"Default Audio Output Device: {defaultSpeaker.FriendlyName}");
        if (defaultMicrophone == null)
            Console.WriteLine("No default microphone device found.");
        else 
            Console.WriteLine($"Default Microphone Device: {defaultMicrophone.FriendlyName}");
        Console.WriteLine(new string('-', 40));
    }

    static void StartAndFocusZoom()
    {
        try
        {
            Process process = Process.Start(@"C:\Program Files\Zoom\bin\Zoom.exe");
            if (process == null)
            {
                Console.WriteLine("Failed to start Zoom.");
                return;
            }

            Console.WriteLine("Zoom application started. Waiting for 'Zoom Meetings' window...");

            IntPtr zoomHWnd = IntPtr.Zero;
            int attempts = 0;

            // Wait up to 10 seconds for Zoom to open (retry every 500ms)
            while (attempts < 20)
            {
                zoomHWnd = FindWindow(null, "Zoom Meetings"); // Try finding the window
                if (zoomHWnd != IntPtr.Zero)
                    break;

                Thread.Sleep(500); // Wait 500ms before retrying
                attempts++;
            }

            if (zoomHWnd != IntPtr.Zero)
            {
                SetForegroundWindow(zoomHWnd);
                Console.WriteLine("Zoom Meetings window brought to the foreground.");
            }
            else
            {
                Console.WriteLine("Unable to find the 'Zoom Meetings' window after waiting.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error starting Zoom: {ex.Message}");
        }
    }

    static IntPtr FindWindowByTitle(string title)
    {
        IntPtr foundHWnd = IntPtr.Zero;

        EnumWindows((hWnd, lParam) =>
        {
            if (IsWindowVisible(hWnd))
            {
                StringBuilder sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, sb.Capacity);

                if (sb.ToString().Contains(title, StringComparison.OrdinalIgnoreCase))
                {
                    foundHWnd = hWnd;
                    return false; // Stop enumeration
                }
            }
            return true; // Continue enumeration
        }, IntPtr.Zero);

        return foundHWnd;
    }

    static void SimulateLogin(string username, string password)
    {
        Console.WriteLine("Logging into Zoom with these credentials...");
        var inputSimulator = new InputSimulator();
        Thread.Sleep(3000);

        // Simulate entering username
        foreach (char c in username)
        {
            inputSimulator.Keyboard.TextEntry(c);
            Thread.Sleep(100);
        }

        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.TAB);

        // Simulate entering password
        foreach (char c in password)
        {
            inputSimulator.Keyboard.TextEntry(c);
            Thread.Sleep(100);
        }

        inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
        Console.WriteLine("Zoom login process simulated.");
    }
}
