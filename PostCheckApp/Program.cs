using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using System.Globalization;
using System.Management;
using System.Diagnostics.PerformanceData;
using System.Security.Policy;
using System.Linq;
using System.ServiceProcess;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Net;



public class MainForm : Form
{
    private string configFilePath = "config.ini";
    
    public MainForm()
    {
        LoadConfig();
        InitializeComponent();
    }

    private void LoadConfig()
{
    if (!File.Exists(configFilePath))
    {
        File.WriteAllText(configFilePath, "# config.ini\n"
            + "ExpectedAudioDevice=ExtronScalerD (HD Audio Driver for Display Audio)\n"
            + "ExpectedMicrophone=COLLABORATE Versa USB Input (COLLABORATE Versa USB)\n"
            + "FileToCheck=C:\\timestamp.txt\n"
            + "FileToCheck2=C:\\timestamp_Instruct.txt\n"
            + "ExpectedZoomPath=C:\\Program Files\\Zoom\\bin\\Zoom.exe\n"
            //+ "ExpectedTeamsPath=C:\\Program Files (x86)\\Teams Installer\\Teams.exe\n"
            //+ "ExpectedTeamsPath2=%localappdata%\\Microsoft\\Teams\\Current\n"
            + "ExpectedVisualizerPath=C:\\Program Files (x86)\\IPEVO\\Visualizer\\Visualizer.exe\n"
            + "ExpectedCitrixPath=C:\\Program Files (x86)\\Citrix\\ICA Client\n"
            + "ZoomInstallPath=\\\\cm\\source\\Zoom\\Zoom CFR\\InstallsZoomClientHDEnabled.bat\n"
            //+ "TeamsInstallPath=\\\\cm\\source\\Microsoft Teams\\Install Teams.cmd\n"
            + "VisualizerInstallPath=\\\\cm\\source\\IPEVO Presenter\\Visualizer_win7-11_v3.6.4.1.msi\n"
            + "OpenCustom=\n");
    }
}

    private string GetConfigValue(string key)
    {
        foreach (var line in File.ReadAllLines(configFilePath))
        {
            if (!line.StartsWith("#") && line.Contains("="))
            {
                var parts = line.Split('=');
                if (parts[0].Trim() == key)
                {
                    return parts[1].Trim();
                }
            }
        }
        return string.Empty;
    }

    private Dictionary<string, ListViewItem> testItems;
    private TextBox txtLog;    
    private Button btnCheckAVSettings, btnCheckFile, btnCheckWindowsUpdate;
    private Button btnCheckZoom, btnOpenZoom, btnOpenCustom, btnUSBDetect;
    private Button btnInstallZoom, /*btnInstallTeams,*/ btnInstallVisualizer, btnCheckVisualizer, btnRunAllTests;
    private Button btnEnableAudioDevices, btnChangeDisplaySetting, btnGPupdate, btnGroupPolicy, btnOpenApp;
    private bool test_passed = false;
    private string panel_color = "#FFFFFF";
    private string button_color = "#F0FFF0";
    private string main_bg_color = "#FDFFBA";
    public string fail_log;
    public bool update_done;
    


    private void InitializeComponent()
    {
        this.Text = "System Diagnostics";
        this.Size = new System.Drawing.Size(420, 800);
        this.MinimumSize = new System.Drawing.Size(420, 800);

        testItems = new Dictionary<string, ListViewItem>();

        // Main content panel
        Panel mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = ColorTranslator.FromHtml(main_bg_color) };

        // Test categories
        GroupBox grpSystemTests = new GroupBox { Text = "System Tests", Dock = DockStyle.Top, Height = 90, BackColor= ColorTranslator.FromHtml(panel_color) };
        GroupBox grpAVTests = new GroupBox { Text = "Audio/Visual Tests", Dock = DockStyle.Top, Height = 90, BackColor= ColorTranslator.FromHtml(panel_color) };
        GroupBox grpZoomTeamsTests = new GroupBox { Text = "Zoom/Teams Tests", Dock = DockStyle.Top, Height = 90, BackColor= ColorTranslator.FromHtml(panel_color) };
        GroupBox grpInstallTests = new GroupBox { Text = "Install Programs", Dock = DockStyle.Top, Height = 90, BackColor= ColorTranslator.FromHtml(panel_color) };
        GroupBox grpInstallTests2 = new GroupBox { Text = "Installations Tests", Dock = DockStyle.Top, Height = 90, BackColor= ColorTranslator.FromHtml(panel_color) };
        GroupBox grpSweepTest = new GroupBox { Text = "Sweep", Dock = DockStyle.Top, Height = 90, BackColor= ColorTranslator.FromHtml(panel_color) };

        FlowLayoutPanel sysPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
        FlowLayoutPanel avPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
        FlowLayoutPanel zoomTeamsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
        FlowLayoutPanel installPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
        FlowLayoutPanel installTestPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
        FlowLayoutPanel runSweepPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };

        // Buttons
        btnCheckWindowsUpdate = new Button { Text = "Check Windows Updates", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnCheckFile = new Button { Text = "Check File Timestamp", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnCheckAVSettings = new Button { Text = "AV Settings Test", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnEnableAudioDevices = new Button { Text = "Enable/Disable Sound Settings", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnGPupdate = new Button { Text = "Run gpupdate /force", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnGroupPolicy = new Button { Text = "Group Policies", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnOpenZoom = new Button { Text = "Open Zoom", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnOpenCustom = new Button { Text = "Open Custom", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };

        btnInstallZoom = new Button { Text = "Install Zoom", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        //btnInstallTeams = new Button { Text = "Install Microsoft Teams", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnInstallVisualizer = new Button { Text = "Install Visualizer", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnCheckVisualizer = new Button { Text = "Check if Visualizer is Installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        //btnCheckTeams = new Button { Text = "Check if Microsoft Teams is installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnCheckZoom = new Button { Text = "Check if Zoom is Installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnChangeDisplaySetting = new Button { Text = "Duplicate Display Settings", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnUSBDetect = new Button { Text = "Detect USB Devices", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnRunAllTests = new Button { Text = "Run Sweep", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
        btnChangeDisplaySetting.Click += (s, e) => ChangeDisplay();
        btnOpenApp = new Button { Text = "Open App", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };

        // Event Handlers
        btnCheckWindowsUpdate.Click += (s, e) => RunTest(CheckWindowsUpdates, btnCheckWindowsUpdate);
        btnCheckFile.Click += (s, e) => RunTest(CheckFileTimestamp, btnCheckFile);
        btnCheckAVSettings.Click += (s, e) => RunTest(CheckAVSettings, btnCheckAVSettings);
        btnEnableAudioDevices.Click += (s, e) => RunTest(OpenSoundSettings, btnEnableAudioDevices);
        btnGPupdate.Click += (s, e) => RunTest(GPUpdate, btnGPupdate);
        btnGroupPolicy.Click += (s, e) => RunTest(GPResult, btnGroupPolicy);
        btnUSBDetect.Click += (s, e) => RunTest(DetectUSBDevices, btnUSBDetect);
        btnOpenApp.Click += (s, e) => showApp();

        //btnCheckTeams.Click += (s, e) => RunTest(CheckMicrosoftTeams, btnCheckTeams);
        btnCheckZoom.Click += (s, e) => RunTest(CheckZoom, btnCheckZoom);
        btnOpenZoom.Click += (s, e) => RunTest(OpenZoom, btnOpenZoom);
        btnOpenCustom.Click += (s, e) => RunTest(OpenCustom, btnOpenCustom);

        btnInstallZoom.Click += (s, e) => RunTest(InstallZoom, btnInstallZoom);
        //btnInstallTeams.Click += (s, e) => RunTest(InstallTeams, btnInstallTeams);
        btnInstallVisualizer.Click += (s, e) => RunTest(InstallVisualizer, btnInstallVisualizer);
        btnCheckVisualizer.Click += (s, e) => RunTest(CheckVisualizer, btnCheckVisualizer);
        btnRunAllTests.Click += (s, e) => RunTest(RunAllTests, btnRunAllTests);

        // Add buttons to respective panels
        sysPanel.Controls.AddRange(new Control[] { btnCheckWindowsUpdate, btnCheckFile, btnGPupdate, btnGroupPolicy });
        avPanel.Controls.AddRange(new Control[] { btnCheckAVSettings, btnEnableAudioDevices, btnChangeDisplaySetting, btnUSBDetect  });
        zoomTeamsPanel.Controls.AddRange(new Control[] { btnOpenApp, btnOpenCustom });
        installPanel.Controls.AddRange(new Control[] { btnInstallZoom, /*btnInstallTeams,*/ btnInstallVisualizer });
        installTestPanel.Controls.AddRange(new Control[] { btnCheckZoom, /*btnCheckTeams,*/ btnCheckVisualizer });
        runSweepPanel.Controls.AddRange(new Control[] { btnRunAllTests });

        // Add panels to group boxes
        grpSystemTests.Controls.Add(sysPanel);
        grpAVTests.Controls.Add(avPanel);
        grpZoomTeamsTests.Controls.Add(zoomTeamsPanel);
        grpInstallTests.Controls.Add(installPanel);
        grpInstallTests2.Controls.Add(installTestPanel);
        grpSweepTest.Controls.Add(runSweepPanel);


        // Log and progress bar
        txtLog = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Font = new Font("Consolas", 10) };
        txtLog.BackColor = Color.White;

        // Add components to main panel
        mainPanel.Controls.Add(txtLog);
        mainPanel.Controls.Add(grpSweepTest);
        mainPanel.Controls.Add(grpInstallTests2);
        mainPanel.Controls.Add(grpInstallTests);
        mainPanel.Controls.Add(grpZoomTeamsTests);
        mainPanel.Controls.Add(grpAVTests);
        mainPanel.Controls.Add(grpSystemTests);

        // Add components to form
        this.Controls.Add(mainPanel);
        //Log("Welcome to the CSUSM CSS Sweep Script\nDeveloped by Reece Harris & Calvary Fisher\n\n");
    }



    private void RunTest(Action testMethod, Button btn)
    {
        btn.Enabled = false;
        test_passed = false;

        testMethod();
        
        Application.DoEvents(); // Keeps UI responsive during long operations

        if (!test_passed) 
            btn.Enabled = true;
    }

    
    
    private ManualResetEvent updateDoneEvent = new ManualResetEvent(false);

    // Assume updateDoneEvent is declared at class level:
// private ManualResetEvent updateDoneEvent = new ManualResetEvent(false);

    private void RunAllTests()
    {
        fail_log = "[FAIL LOG]\n\n";
        Log("Running all tests...\n");
        
        updateDoneEvent.Reset(); // Reset the event for this test
        RunTest(CheckWindowsUpdates, btnCheckWindowsUpdate);
        updateDoneEvent.WaitOne(); // Wait until CheckWindowsUpdates signals completion

        // AV Test - Execute CheckAVSettings and CheckFileTimestamp sequentially
        updateDoneEvent.Reset();
        RunTest(CheckAVSettings, btnCheckAVSettings);
        updateDoneEvent.WaitOne();

        updateDoneEvent.Reset();
        RunTest(CheckFileTimestamp, btnCheckFile);
        updateDoneEvent.WaitOne();
        
        // Install Test - Run each test sequentially       
        /* 
        updateDoneEvent.Reset();
        RunTest(CheckMicrosoftTeams, btnCheckTeams);
        updateDoneEvent.WaitOne();
        */
        updateDoneEvent.Reset();
        RunTest(CheckZoom, btnCheckZoom);
        updateDoneEvent.WaitOne();

        updateDoneEvent.Reset();
        RunTest(CheckVisualizer, btnCheckVisualizer);
        updateDoneEvent.WaitOne();

        updateDoneEvent.Reset();
        RunTest(OpenZoom, btnCheckZoom);
        updateDoneEvent.WaitOne();
        updateDoneEvent.Reset();
        // Report the final result
        if (fail_log != "[FAIL LOG]\n\n")
        {
            Log("\n\n\n\n\n");
            Log(fail_log);
            test_passed = false;
        }
        else
        {
            Log("All tests passed!\n");
            test_passed = true;
        }
        CheckFileTimestamp();
        Log("[TEST COMPLETE]");
    }

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(ref Guid clsid, IntPtr pUnkOuter, uint dwClsContext, ref Guid riid, out IPolicyConfig ppv);

    private Guid CLSID_PolicyConfig = new Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9");
    private Guid IID_IPolicyConfig = new Guid("f8679f50-850a-41cf-9c72-430f290290c8");

    [ComImport]
    [Guid("f8679f50-850a-41cf-9c72-430f290290c8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPolicyConfig
    {
        void SetDefaultEndpoint(string deviceId, Role role);
    }
    public void OpenSoundSettings()
    {
        Log("Opening Windows Sound settings...");

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "mmsys.cpl",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Log("Failed to open Sound settings: " + ex.Message);
        }
        updateDoneEvent.Set();
    }


    private void CheckAVSettings()
    {
        Log("[Checking AV settings...]\n");
        var enumerator = new MMDeviceEnumerator();
        var defaultPlayback = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        var defaultRecording = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia);
        
        string expectedPlayback = GetConfigValue("ExpectedAudioDevice");
        string expectedRecording = GetConfigValue("ExpectedMicrophone");
        
        bool playbackMatch = defaultPlayback.FriendlyName.Equals(expectedPlayback, StringComparison.OrdinalIgnoreCase);
        bool recordingMatch = defaultRecording.FriendlyName.Equals(expectedRecording, StringComparison.OrdinalIgnoreCase);
        if (playbackMatch) {
            Log("Playback devices match! ✔");
        }
        else {
            Log("Playback devices do not match! ✖");
            fail_log += "Playback devices do not match! ✖\n\n";
        }
        if (recordingMatch) {
            Log("Recording devices match! ✔");
        }
        else {
            Log("Recording devices do not match! ✖");
            fail_log += "Recording devices do not match! ✖\n\n";
        }
        Log($"Default Playback Device: {defaultPlayback.FriendlyName}");
        Log($"Default Recording Device: {defaultRecording.FriendlyName}");
        Log("\n");
        Log($"Expected Playback Device: {expectedPlayback}");
        Log($"Expected Recording Device: {expectedRecording}");
        Log("\n");
        DetectCameraDevice();
        Log("\n\n");
        updateDoneEvent.Set();
    }
private void ChangeDisplay()
{
    // Display the options for display settings
    DialogResult result = MessageBox.Show(
        "Duplicate Display Settings?", 
        "Duplicate Display Settings", 
        MessageBoxButtons.YesNoCancel, 
        MessageBoxIcon.Question);

    // Yes = Extend Display, No = Duplicate Display, Cancel = do nothing
    if (result == DialogResult.Yes)
    {
        // Duplicate Display
        SetDisplayMode("Duplicate");
    }
    else if (result == DialogResult.No)
    {
        // Extend Display
        SetDisplayMode("Extend");
    }
    else
    {
        // Cancel: Do nothing
        Log("Display setting change was canceled.");
    }
    updateDoneEvent.Set();
}

private void SetDisplayMode(string mode)
{
    try
    {
        string modeCommand = mode == "Extend" ? "/extend" : "/clone";  // /extend for extend, /clone for duplicate

        // Use DisplaySwitch.exe to apply the display settings
        Process.Start(new ProcessStartInfo
        {
            FileName = "DisplaySwitch.exe", // Built-in Windows tool for display configurations
            Arguments = modeCommand,  // Arguments to set either extend or duplicate
            UseShellExecute = true
        });

        Log($"Display set to {mode} mode.");
    }
    catch (Exception ex)
    {
        Log($"Failed to change display settings: {ex.Message}");
    }
    
}

private void CheckFileTimestamp()
{
    Log("[Timestamp Check]\n");
    string filePath = GetConfigValue("FileToCheck");
    string filePath2 = GetConfigValue("FileToCheck2");
    string content = null;

    if (File.Exists(filePath))
    {
        content = File.ReadAllText(filePath);
    }
    else if (File.Exists(filePath2))
    {
        content = File.ReadAllText(filePath2);
    }

    if (content != null)
    {
        Log("Timestamp file contents:\n" + content + "\n✔\n");
    }
    else
    {
        Log("Timestamp file not found! ✖\n");
        fail_log += "Timestamp file not found! ✖\n";
    }

    updateDoneEvent.Set();
    Log("\n\n");
}
    const int VK_TAB = 0x09;
    const int VK_RETURN = 0x0D;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern short GetAsyncKeyState(int vKey);
    private const int SW_MINIMIZE = 6;

    private void CheckWindowsUpdates()
    {
        Log("[Windows Update Check]\n");

        try
        {
            // Close existing Settings windows
            var processes = Process.GetProcessesByName("SystemSettings");
            foreach (var process in processes)
            {
                Console.WriteLine("Closing Settings app...");
                process.Kill();
                process.WaitForExit();
            }

            // Open Windows Update settings
            Process.Start("explorer.exe", "ms-settings:windowsupdate");
            Console.WriteLine("Windows Update settings opened successfully! ✔");

            Thread.Sleep(5000);  // Wait for Settings to open

            // Simulate keystrokes
            SimulateKeyPress(VK_TAB);
            Thread.Sleep(300);
            SimulateKeyPress(VK_TAB);
            Thread.Sleep(300);
            SimulateKeyPress(VK_TAB);
            Thread.Sleep(300);

            Version osVersion = Environment.OSVersion.Version;
            if (osVersion.Major == 10 && osVersion.Build < 22000)
            {
                SimulateKeyPress(VK_TAB);
                Thread.Sleep(300);
            }

            SimulateKeyPress(VK_RETURN);  // Press Enter to check for updates
            Log("Triggered 'Check for updates' successfully! ✔");

            // Give time for the button press to register before minimizing
            Thread.Sleep(2000);

            update_done = true;
            updateDoneEvent.Set();
            processes = Process.GetProcessesByName("SystemSettings");
            foreach (var process in processes)
            {
                Console.WriteLine("Closing Settings app...");
                process.Kill();
                process.WaitForExit();
            }
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
            updateDoneEvent.Set();
        }
        finally
        {
            Log("\n\n");
        }
    }
    

    // Method to simulate key press
    private static void SimulateKeyPress(int keyCode)
    {
        keybd_event((byte)keyCode, 0, 0, 0);  // Key down
        keybd_event((byte)keyCode, 0, 2, 0);  // Key up
    }

    /*
    private void CheckMicrosoftTeams()
    {
        Log("[Microsoft Teams Install Check]\n");
        string expectedPath = GetConfigValue("ExpectedTeamsPath");
        string expectedPath2 = GetConfigValue("ExpectedTeamsPath2");
        string expectedPath3 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Teams");
        bool isInstalled = File.Exists(expectedPath);
        bool isInstalled2 = File.Exists(expectedPath2);
        bool isInstalled3 = Directory.Exists(expectedPath3);
        if (isInstalled || isInstalled2 || isInstalled3)
        {
            Log($"Microsoft Teams is installed: ✔");
            if (isInstalled)
                Log($"Path: {expectedPath}");
            if (isInstalled2)
                Log($"Path: {expectedPath2}");
            if (isInstalled3)
                Log($"Path: {expectedPath3}");
            test_passed = true;
        }
        else {
            Log("Microsoft Teams is not installed: ✖");
            DialogResult result = MessageBox.Show("Teams is not installed, would you like to install?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {       
                Log("[User clicked Yes on Teams install]");
                InstallTeams();
            }
                else
            {
                Log("[User clicked No on Teams install]");
                fail_log += "Microsoft Teams is not installed: ✖\n";
            }
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }
    */

    private void CheckZoom()
    {
        Log("[Zoom Install Check]\n");
        string expectedPath = GetConfigValue("ExpectedZoomPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled)
        {
            Log($"Zoom is installed: ✔");
            Log($"Path: {expectedPath}");
            test_passed = true;
        }
        else {
            Log("Zoom is not installed: ✖");
            DialogResult result = MessageBox.Show("Zoom is not installed, would you like to install?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {       
                Log("[User clicked Yes on Zoom install]");
                InstallZoom();
            }
            else
            {
                Log("[User clicked No on Zoom install]");
                fail_log += "Zoom is not installed: ✖\n";
            }
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }

    private void DetectCameraDevice()
{
    List<string> cameraNames = new List<string> {"AVer TR313V2", "AVer TR313"};
    Log("[Checking for AVER camera connection]\n");
    try
    {
        // Query Win32_PnPEntity to get camera devices
        ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE PNPClass = 'Camera'");

        foreach (ManagementObject device in searcher.Get())
        {
            string deviceName = device["Name"]?.ToString() ?? "Unknown Device";
            // Check if the detected device matches any name in the provided list
            foreach (string name in cameraNames)
            {
                if (deviceName.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Log($"Camera detected: {deviceName} ✔");
                    return;
                }
            }
        }
        Log("No matching camera device found! ✖");
        fail_log += "No matching camera device found! ✖\n";
        updateDoneEvent.Set();
    }
    catch (Exception e)
    {
        Log($"Error detecting camera: {e.Message}");
        updateDoneEvent.Set();
    }
}

    private void InstallZoom() {
        string zoomDownloadUrl = "https://zoom.us/client/latest/ZoomInstallerFull.msi"; // Official Zoom installer URL
        string tempPath = Path.Combine(Path.GetTempPath(), "ZoomInstaller.msi");

        try
        {
            Log("Downloading latest Zoom installer...");
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(zoomDownloadUrl, tempPath);
            }
            Log("Zoom download completed ✔");

            // Run the installer
            ProcessStartInfo startInfo = new ProcessStartInfo(tempPath)
            {
                UseShellExecute = true,
                Verb = "runas" // Run as administrator
            };

            Process.Start(startInfo);
            Log("Zoom installation started ✔");
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message} ✖");
            fail_log += "Could not install Zoom ✖\n";
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }
    /*
    private void InstallTeams() {
        string expectedPath = GetConfigValue("TeamsInstallPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            ProcessStartInfo startInfo = new ProcessStartInfo(expectedPath) {
                UseShellExecute = true,
                Verb = "runas" // This specifies to run the process as an administrator
            };
            try {
                Process.Start(startInfo);
                Log("Initialized Teams install: ✔");
            } catch (System.ComponentModel.Win32Exception) {
                Log("The user declined the elevation request ✖");
            }
        } else {
            Log("Could not connect to the network folder: ✖");
            fail_log += "Could not install Teams ✖\n\n";
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }
    */

    private void InstallVisualizer() {
        string expectedPath = GetConfigValue("VisualizerInstallPath");
        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = "msiexec",
            Arguments = $"/i \"{expectedPath}\"" // Use /i for installation
        };

        try
        {
            Process.Start(startInfo);
            Log("Initialized Visualizer install: ✔");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}");
            fail_log += "Could not install Visualizer ✖\n\n";
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }

    private void OpenZoom()
    {
        Log("[Attempting to open Zoom...]\n");
        string expectedPath = GetConfigValue("ExpectedZoomPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            Process.Start(expectedPath);
            Log("Zoom opened successfully! ✔");
        }
        else if (!isInstalled) {
            Log("Zoom is not installed! ✖");
            DialogResult result = MessageBox.Show("Do you want to proceed?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {       
                Log("[User clicked Yes on Zoom install]");
                InstallZoom();
            }
                else
            {
                Log("[User clicked No on Zoom install]");
                fail_log += "[User clicked No on Zoom install]\n";
            }
            }
        else {
            Log("Failed to open Zoom! ✖");
            fail_log += "Failed to open Zoom! ✖\n";
        }
        updateDoneEvent.Set();
    }   

    private void OpenCustom()
    {
        Log("Attempting to open custom path...");
        string expectedPath = GetConfigValue("ExpectedZoomPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            Process.Start(expectedPath);
            Log("Custom program opened successfully! ✔");
        }
        else if (!isInstalled) {
            Log("Custom program path is not correct! ✖");
            fail_log += "Custom program path is not correct! ✖";
    
        }
        else {
            Log("Could not start custom program! ✖");
            fail_log += "Could not start custom program! ✖";
        }
        updateDoneEvent.Set();

    }


    private void DetectUSBDevices()
    {
        try
        {
            // Query Win32_PnPEntity to get USB devices
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_PnPEntity WHERE PNPClass = 'USB'");

            bool foundDevices = false;

            foreach (ManagementObject usbDevice in searcher.Get())
            {
                string friendlyName = usbDevice["Name"]?.ToString() ?? "Unknown USB Device";
                string deviceID = usbDevice["DeviceID"]?.ToString() ?? "N/A";
                string manufacturer = usbDevice["Manufacturer"]?.ToString() ?? "Unknown Manufacturer";
                string pnpDeviceID = usbDevice["PNPDeviceID"]?.ToString() ?? "N/A";
                string status = usbDevice["Status"]?.ToString() ?? "Unknown";

                Log($"USB Device: {friendlyName}");
                Log($"  - Device ID: {deviceID}");
                Log($"  - Manufacturer: {manufacturer}");
                Log($"  - PNP Device ID: {pnpDeviceID}");
                Log($"  - Status: {status}");
                Log("\n");

                foundDevices = true;
                test_passed = true;
            }

            if (!foundDevices)
            {
                Log("No USB devices detected.");
            }
        }
        catch (Exception e)
        {
            Log($"Error getting USB devices: {e.Message}");
        }
        updateDoneEvent.Set();
    }

    private void CheckVisualizer()
    {
        Log("[Visualizer Install Check]\n");
        string expectedPath = GetConfigValue("ExpectedVisualizerPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled)
        {
            Log($"Visualizer is installed: ✔");
            Log($"Path: {expectedPath}");
            test_passed = true;
        }
        else {
            Log("Visualizer is not installed: ✖");
            DialogResult result = MessageBox.Show("Visualizer is not installed, would you like to install?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {       
                Log("[User clicked Yes on Visualizer install]");
                InstallVisualizer();
            }
                else
            {
                Log("[User clicked No on Visualizer install]");
                fail_log += "Visualizer is not installed ✖\n";
            }
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }
    private void CheckCitrix()
    {
        string expectedPath = GetConfigValue("ExpectedCitrixPath");
        bool isInstalled = Path.Exists(expectedPath);
        if (isInstalled)
        {
            Log($"Citrix is installed: ✔");
            Log($"Path: {expectedPath}");
            test_passed = true;
        }
        else {
            Log("Citrix is not installed: ✖");
            fail_log += "Citrix is not installed: ✖\n";
        
        }
        updateDoneEvent.Set();
        Log("\n\n");
    }

    private void GPUpdate()
    {
        Log("Starting GPUpdate...");

        // Create a new process start info
        ProcessStartInfo processStartInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = "/c gpupdate /force",
            UseShellExecute = true
        };

        try
        {
            // Start the process
            using (Process process = Process.Start(processStartInfo))
            {
                // Wait for the process to exit
                process.WaitForExit();
            }

            // Set test_passed to true after the process exits
            test_passed = true;
            Log("GPUpdate completed successfully.");
        }
        catch (Exception ex)
        {
            Log($"GPUpdate failed: {ex.Message}");
        }
        updateDoneEvent.Set();
    }

    private void showApp() {
        ShowAppSelectionDialog();
        if (citrix_open == true) {
            CheckCitrix();
        }
    }

    private void ShowAppSelectionDialog()
    {
        Form appDialog = new Form
        {
            Text = "Select an App to Open",
            Size = new Size(300, 300),
            StartPosition = FormStartPosition.CenterParent
        };
        

        FlowLayoutPanel panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };

        // Buttons for different apps/websites
        Button btnZoom = new Button { Text = "Zoom", Width = 120 };
        //Button btnTeams = new Button { Text = "Teams", Width = 120 };
        Button btnVisualizer = new Button { Text = "Visualizer", Width = 120 };
        Button btnCougarApps = new Button { Text = "Cougar Apps", Width = 120 };

        // Assign click events with installation checks
        btnZoom.Click += (s, e) => CheckAndOpenApp("Zoom", GetConfigValue("ExpectedZoomPath"), InstallZoom);
        //btnTeams.Click += (s, e) => CheckAndOpenApp("Microsoft Teams", GetConfigValue("ExpectedTeamsPath"), InstallTeams);
        btnVisualizer.Click += (s, e) => CheckAndOpenApp("Visualizer", GetConfigValue("ExpectedVisualizerPath"), InstallVisualizer);
        btnCougarApps.Click += (s, e) => OpenWebsite("https://cougarapps.csusm.edu/Citrix/CougarAppsProdWeb/");

        // Add buttons to panel
        panel.Controls.AddRange(new Control[] { btnZoom, btnVisualizer, btnCougarApps });

        appDialog.Controls.Add(panel);
        appDialog.ShowDialog();
    }

    private void CheckAndOpenApp(string appName, string appPath, Action installMethod)
{
    if (File.Exists(appPath))
    {
        Log($"{appName} is installed: ✔");
        Log($"Path: {appPath}");
        OpenApplication(appPath);
        if (appPath == "https://cougarapps.csusm.edu/Citrix/CougarAppsProdWeb/")
            CheckCitrix();
    }
    else
    {
        Log($"{appName} is not installed: ✖");
        DialogResult result = MessageBox.Show($"{appName} is not installed. Do you want to install it?", "Confirmation", MessageBoxButtons.YesNo);
        if (result == DialogResult.Yes)
        {
            Log($"[User clicked Yes on {appName} install]");
            installMethod.Invoke();
        }
        else
        {
            Log($"[User clicked No on {appName} install]");
        }
    }
    updateDoneEvent.Set();
}



private void OpenApplication(string appPath)
{
    try
    {
        System.Diagnostics.Process.Start(appPath);
        updateDoneEvent.Set();
    }
    catch (Exception ex)
    {
        MessageBox.Show("Error opening application: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        updateDoneEvent.Set();
    }
    
}

public bool citrix_open;

private void OpenWebsite(string url)
{
    citrix_open = false;
    try
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
        updateDoneEvent.Set();
        citrix_open = true;

    }
    catch (Exception ex)
    {
        MessageBox.Show("Error opening website: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        updateDoneEvent.Set();
    }
}


    private void GPResult()
    {
        // Create a process to run the gpresult command
        Process gpResultProcess = new Process();
        gpResultProcess.StartInfo.FileName = "cmd.exe";
        gpResultProcess.StartInfo.Arguments = "/c gpresult /Scope User /v";
        gpResultProcess.StartInfo.RedirectStandardOutput = true;
        gpResultProcess.StartInfo.UseShellExecute = false;
        gpResultProcess.StartInfo.CreateNoWindow = true;

        // Start the gpresult process and read the output
        gpResultProcess.Start();
        string gpResultOutput = gpResultProcess.StandardOutput.ReadToEnd();
        gpResultProcess.WaitForExit();

        // Extract the "Administrative Templates" section
        string adminTemplatesSection = ExtractAdministrativeTemplates(gpResultOutput);

        // Get the groups of the current computer
        string computerGroups = GetComputerGroups();

        // Combine the outputs
        string combinedOutput = adminTemplatesSection + "\n\nComputer Groups:\n" + computerGroups;

        // Create a temporary file with the combined output
        string tempFilePath = Path.GetTempFileName();
        File.WriteAllText(tempFilePath, combinedOutput);

        // Open the temporary file in Notepad
        Process.Start("notepad.exe", tempFilePath);

        // Delete the temporary file after a delay to ensure Notepad has time to open it
        System.Threading.Thread.Sleep(5000); // Wait for 5 seconds
        File.Delete(tempFilePath);
        updateDoneEvent.Set();
    }

    private string ExtractAdministrativeTemplates(string output)
    {
        // Use a regular expression to extract the "Administrative Templates" section
        Regex regex = new Regex(@"Administrative Templates[\s\S]*?(?=\r?\n\r?\n|$)", RegexOptions.IgnoreCase);
        Match match = regex.Match(output);
        return match.Success ? match.Value : "Administrative Templates section not found.";
    }

    private string GetComputerGroups()
    {
        // Create a process to run the net group command
        Process netGroupProcess = new Process();
        netGroupProcess.StartInfo.FileName = "cmd.exe";
        netGroupProcess.StartInfo.Arguments = "/c net localgroup";
        netGroupProcess.StartInfo.RedirectStandardOutput = true;
        netGroupProcess.StartInfo.UseShellExecute = false;
        netGroupProcess.StartInfo.CreateNoWindow = true;

        // Start the net group process and read the output
        netGroupProcess.Start();
        string netGroupOutput = netGroupProcess.StandardOutput.ReadToEnd();
        netGroupProcess.WaitForExit();

        return netGroupOutput;
    }


        private void Log(string message)
        {
            Invoke(new Action(() => {
                string formattedMessage = message.Replace("\n", Environment.NewLine);
                txtLog.AppendText(formattedMessage + Environment.NewLine);
            }));
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        public class AudioDeviceManager
    {
        private readonly MMDeviceEnumerator deviceEnumerator;

        public AudioDeviceManager()
        {
            deviceEnumerator = new MMDeviceEnumerator();
        }

    }
    
}
