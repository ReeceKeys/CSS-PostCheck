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
using NAudio.CoreAudioApi;
using System.Linq;
using System.ServiceProcess;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;



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
            + "ExpectedTeamsPath=C:\\Program Files (x86)\\Teams Installer\\Teams.exe\n"
            + "ExpectedVisualizerPath=C:\\Program Files (x86)\\IPEVO\\Visualizer\\Visualizer.exe\n"
            + "ZoomInstallPath=\\\\cm\\source\\Zoom\\Zoom CFR\\InstallsZoomClientHDEnabled.bat\n"
            + "TeamsInstallPath=https://go.microsoft.com/fwlink/?linkid=2281613&clcid=0x409&culture=en-us&country=us\n"
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
    private ProgressBar progressBar;
    
    private Button btnCheckAVSettings, btnCheckFile, btnCheckWindowsUpdate;
    private Button btnCheckTeams, btnCheckZoom, btnOpenZoom, btnOpenCustom;
    private Button btnInstallZoom, btnInstallTeams, btnInstallVisualizer, btnCheckVisualizer;
    private Button btnEnableAudioDevices, btnChangeDisplaySetting;
    private bool test_passed = false;
    private string panel_color = "#B7E892";
    private string button_color = "#F0FFF0";



    private void InitializeComponent()
{
    this.Text = "System Diagnostics";
    this.Size = new System.Drawing.Size(600, 600);
    this.MinimumSize = new System.Drawing.Size(600, 800);

    testItems = new Dictionary<string, ListViewItem>();

    // Main content panel
    Panel mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };

    // Test categories
    GroupBox grpSystemTests = new GroupBox { Text = "System Tests", Dock = DockStyle.Top, Height = 100 };
    GroupBox grpAVTests = new GroupBox { Text = "Audio/Visual Tests", Dock = DockStyle.Top, Height = 100 };
    GroupBox grpZoomTeamsTests = new GroupBox { Text = "Zoom/Teams Tests", Dock = DockStyle.Top, Height = 130 };
    GroupBox grpInstallTests = new GroupBox { Text = "Installations", Dock = DockStyle.Top, Height = 130 };

    FlowLayoutPanel sysPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel avPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel zoomTeamsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel installPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };

    // Buttons
    btnCheckWindowsUpdate = new Button { Text = "Check Windows Updates", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckFile = new Button { Text = "Check File Timestamp", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckAVSettings = new Button { Text = "AV Settings Test", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnEnableAudioDevices = new Button { Text = "Open Sound Settings", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };

    btnOpenZoom = new Button { Text = "Open Zoom", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnOpenCustom = new Button { Text = "Open Custom", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };

    btnInstallZoom = new Button { Text = "Install Zoom", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnInstallTeams = new Button { Text = "Install Microsoft Teams", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnInstallVisualizer = new Button { Text = "Install Visualizer", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckVisualizer = new Button { Text = "Check if Visualizer is Installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckTeams = new Button { Text = "Check if Microsoft Teams is installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckZoom = new Button { Text = "Check if Zoom is Installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnChangeDisplaySetting = new Button { Text = "Change Display Setting", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnChangeDisplaySetting.Click += (s, e) => ChangeDisplay();



    // Event Handlers
    btnCheckWindowsUpdate.Click += (s, e) => RunTest(CheckWindowsUpdates, btnCheckWindowsUpdate);
    btnCheckFile.Click += (s, e) => RunTest(CheckFileTimestamp, btnCheckFile);
    btnCheckAVSettings.Click += (s, e) => RunTest(CheckAVSettings, btnCheckAVSettings);
    btnEnableAudioDevices.Click += (s, e) => RunTest(OpenSoundSettings, btnEnableAudioDevices);

    btnCheckTeams.Click += (s, e) => RunTest(CheckMicrosoftTeams, btnCheckTeams);
    btnCheckZoom.Click += (s, e) => RunTest(CheckZoom, btnCheckZoom);
    btnOpenZoom.Click += (s, e) => RunTest(OpenZoom, btnOpenZoom);
    btnOpenCustom.Click += (s, e) => RunTest(OpenCustom, btnOpenCustom);


    btnInstallZoom.Click += (s, e) => RunTest(InstallZoom, btnInstallZoom);
    btnInstallTeams.Click += (s, e) => RunTest(InstallTeams, btnInstallTeams);
    btnInstallVisualizer.Click += (s, e) => RunTest(InstallVisualizer, btnInstallVisualizer);
    btnCheckVisualizer.Click += (s, e) => RunTest(CheckVisualizer, btnCheckVisualizer);

    // Add buttons to respective panels
    sysPanel.Controls.AddRange(new Control[] { btnCheckWindowsUpdate, btnCheckFile });
    avPanel.Controls.AddRange(new Control[] { btnCheckAVSettings, btnEnableAudioDevices, btnChangeDisplaySetting });
    zoomTeamsPanel.Controls.AddRange(new Control[] { btnOpenZoom, btnOpenCustom });
    installPanel.Controls.AddRange(new Control[] { btnInstallZoom, btnInstallTeams, btnInstallVisualizer, btnCheckZoom, btnCheckTeams, btnCheckVisualizer });

    // Add panels to group boxes
    grpSystemTests.Controls.Add(sysPanel);
    grpAVTests.Controls.Add(avPanel);
    grpZoomTeamsTests.Controls.Add(zoomTeamsPanel);
    grpInstallTests.Controls.Add(installPanel);

    // Log and progress bar
    txtLog = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Font = new Font("Consolas", 10) };
    txtLog.BackColor = Color.White;  
    progressBar = new ProgressBar { Dock = DockStyle.Bottom, Height = 20, Style = ProgressBarStyle.Marquee, Visible = false };

    // Add components to main panel
    mainPanel.Controls.Add(txtLog);
    mainPanel.Controls.Add(progressBar);
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
        btn.Enabled = false;  // Disable button to prevent multiple rapid clicks
        progressBar.Visible = true;
        test_passed = false; 

        Thread testThread = new Thread(() =>
        {
            testMethod();

            // Re-enable button and hide progress bar after test completes
            Invoke(new Action(() =>
            {
                if (test_passed != true) {
                    btn.Enabled = true;
                }
                progressBar.Visible = false;
                
            }));
        });

        testThread.IsBackground = true;
        testThread.Start();
    }

    private void RunAllTests()
    {
        CheckAVSettings();
        CheckFileTimestamp();
        CheckWindowsUpdates();
        CheckMicrosoftTeams();
        CheckZoom();
        OpenZoom();
        InstallZoom();
        CheckVisualizer();
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
    public static void OpenSoundSettings()
    {
        Console.WriteLine("Opening Windows Sound settings...");

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
            Console.WriteLine("Failed to open Sound settings: " + ex.Message);
        }
    }


    private void CheckAVSettings()
    {
        Log("Checking AV settings...");
        var enumerator = new MMDeviceEnumerator();
        var defaultPlayback = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        var defaultRecording = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Multimedia);
        
        string expectedPlayback = GetConfigValue("ExpectedAudioDevice");
        string expectedRecording = GetConfigValue("ExpectedMicrophone");
        
        bool playbackMatch = defaultPlayback.FriendlyName.Equals(expectedPlayback, StringComparison.OrdinalIgnoreCase);
        bool recordingMatch = defaultRecording.FriendlyName.Equals(expectedRecording, StringComparison.OrdinalIgnoreCase);
        
        if (playbackMatch && recordingMatch) {
            test_passed = true;
        }
        Log($"Default Playback Device: {defaultPlayback.FriendlyName} {(playbackMatch ? "✔" : "✖")}");
        Log($"Default Recording Device: {defaultRecording.FriendlyName} {(recordingMatch ? "✔" : "✖")}");
        Log("\n");
        Log($"Expected Playback Device: {expectedPlayback}");
        Log($"Expected Recording Device: {expectedRecording}");
        Log("\n\n");
    }
private void ChangeDisplay()
{
    // Display the options for display settings
    DialogResult result = MessageBox.Show(
        "Duplicate Display Settings? (No - Extend)", 
        "Change Display Setting", 
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
        Console.WriteLine("Display setting change was canceled.");
    }
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
    string filePath = GetConfigValue("FileToCheck");
    string filePath2 = GetConfigValue("FileToCheck2");
    if (File.Exists(filePath))
    {
        Process.Start("notepad.exe", filePath);
        Log("Timestamp file opened successfully! ✔");
        test_passed = true;
    }
    else if (File.Exists(filePath2)) {
        Process.Start("notepad.exe", filePath2);
        Log("Timestamp file opened successfully! ✔");
        test_passed = true;
    }
    else
    {
        Log("Timestamp file not found! ✖");
    }
}
    private void CheckWindowsUpdates()
    {
        Log("Checking Windows Updates...");
        try
        {
            Process.Start("explorer.exe", "ms-settings:windowsupdate");
            Log("Windows Update settings opened successfully! ✔");
            test_passed = true;
        }
        catch (Exception ex)
        {
            Log($"Failed to open Windows Update settings ✖\nError: {ex.Message}");
        }
        Log("\n\n");
    }

    private void CheckMicrosoftTeams()
    {
        string expectedPath = GetConfigValue("ExpectedTeamsPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled)
        {
            Log($"Microsoft Teams is installed: ✔");
            Log($"Path: {expectedPath}");
            test_passed = true;
        }
        else {
            Log("Microsoft Teams is not installed: ✖");
            DialogResult result = MessageBox.Show("Do you want to proceed?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {       
                Log("[User clicked Yes on Teams install]");
                InstallTeams();
            }
                else
            {
                Log("[User clicked No on Teams install]");
            }
        }
        Log("\n\n");
    }

    private void CheckZoom()
    {
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
            DialogResult result = MessageBox.Show("Do you want to proceed?", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {       
                Log("[User clicked Yes on Zoom install]");
                InstallZoom();
            }
                else
            {
                Log("[User clicked No on Zoom install]");
            }
        }
        Log("\n\n");
    }

    private void InstallZoom() {
        string expectedPath = GetConfigValue("ZoomInstallPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            ProcessStartInfo startInfo = new ProcessStartInfo(expectedPath) {
                UseShellExecute = true,
                Verb = "runas" // This specifies to run the process as an administrator
            };
            try {
                Process.Start(startInfo);
            } catch (System.ComponentModel.Win32Exception) {
                Log("The user declined the elevation request.");
            }
        } else {
            Log("Could not connect to the network folder: ✖");
        }
        Log("\n\n");
    }

    private void InstallTeams() {
        string downloadUrl = GetConfigValue("TeamsInstallPath");
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = downloadUrl,
                UseShellExecute = true
            });
            Log("Microsoft Teams download page opened successfully! ✔");
            test_passed = true;
        }
        catch (Exception ex)
        {
            Log($"Failed to open Microsoft Teams download page: ✖\nError: {ex.Message}");
        }
        Log("\n\n");
    }
    

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
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}");
        }
        Log("\n\n");
    }

    private void OpenZoom()
    {
        Log("Attempting to open Zoom...");
        string expectedPath = GetConfigValue("ExpectedZoomPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            Process.Start(expectedPath);
            Log("Zoom opened successfully! ✔");
            test_passed = true;
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
            }
            }
        else {
            Log("Failed to open Zoom! ✖");
        }

    }

    private void OpenCustom()
    {
        Log("Attempting to open custom path...");
        string expectedPath = GetConfigValue("ExpectedZoomPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            Process.Start(expectedPath);
            Log("Program opened successfully! ✔");
        }
        else if (!isInstalled) {
            Log("Program path is not correct! ✖");
        }
        else {
            Log("Could not start file! ✖");
        }

    }

    private void SetBothAudio() {
        Log("Attempting to set both audio devices to default...");
        SetDefaultAudioDevice($"{GetConfigValue("ExpectedMicrophone")}", DataFlow.Capture);
        SetDefaultAudioDevice($"{GetConfigValue("ExpectedAudioDevice")}", DataFlow.Render);
    }

    private void SetDefaultAudioDevice(string deviceName, DataFlow dataFlow)
{
    var enumerator = new MMDeviceEnumerator();
    var devices = enumerator.EnumerateAudioEndPoints(dataFlow, DeviceState.Active);

    foreach (var device in devices)
    {
        if (device.FriendlyName.Equals(deviceName, StringComparison.OrdinalIgnoreCase))
        {
            // Unmute the device (optional)
            device.AudioEndpointVolume.Mute = false;
            Log($"Unmuted {dataFlow} device: {device.FriendlyName}");

            // Now set this device as the default
            SetAsDefaultAudioDevice(device.ID);
            Log($"Set {dataFlow} device as default: {device.FriendlyName}");
            break;
        }
    }
}

// Using the IPolicyConfig COM interface to set the default device
private void SetAsDefaultAudioDevice(string deviceId)
{
    // Create the IPolicyConfig instance
    IPolicyConfig policyConfig;
    int result = CoCreateInstance(ref CLSID_PolicyConfig, IntPtr.Zero, 0, ref IID_IPolicyConfig, out policyConfig);
    
    if (result != 0)
    {
        Log($"CoCreateInstance failed with error code {result}");
        return;
    }
    
    try
    {
        // Set this device as the default for multimedia role
        policyConfig.SetDefaultEndpoint(deviceId, Role.Multimedia);
        Log($"Successfully set device {deviceId} as the default multimedia device.");
    }
    catch (Exception ex)
    {
        Log($"Error setting default device: {ex.Message}");
    }
}


    private void EnableDevice(string deviceId)
    {
        using (var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_PnPEntity WHERE DeviceID = '{deviceId.Replace("\\", "\\\\")}'"))
        {
            foreach (ManagementObject obj in searcher.Get())
            {
                obj.InvokeMethod("Enable", null);
            }
        }
    }

    private void CheckVisualizer()
    {
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
        }
        Log("\n\n");
    }

    private void Log(string message)
    {
        Invoke(new Action(() => txtLog.AppendText(message + Environment.NewLine)));
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
