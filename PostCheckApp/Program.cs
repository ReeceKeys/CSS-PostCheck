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

public class MainForm : Form
{
    private Button btnRunAllTests, btnCheckAVSettings, btnCheckFile, btnCheckWindowsUpdate, btnCheckTeams, btnCheckZoom, btnCheckVisualizer, btnOpenZoom, btnInstallZoom, btnInstallTeams, btnInstallVisualizer, btnDisableAudioDevices, btnEnableAudioDevices;
    private TextBox txtLog;
    private ProgressBar progressBar;
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
            + "VisualizerInstallPath=\\\\cm\\source\\IPEVO Presenter\\Visualizer\\Visualizer_win7-11_v3.6.4.1.msi\n");
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

    private void InitializeComponent()
    {
        this.Text = "System Diagnostics";
        this.Size = new System.Drawing.Size(800, 600);
        this.MinimumSize = new System.Drawing.Size(600, 400);

        btnRunAllTests = new Button { Text = "Run All Tests", Dock = DockStyle.Top, Height = 40 };
        btnRunAllTests.Click += (s, e) => RunTest(RunAllTests, btnRunAllTests);
        
        btnCheckAVSettings = new Button { Text = "AV Settings Test", Dock = DockStyle.Top, Height = 40 };
        btnCheckAVSettings.Click += (s, e) => RunTest(CheckAVSettings, btnCheckAVSettings);
        
        btnCheckFile = new Button { Text = "Check File Timestamp", Dock = DockStyle.Top, Height = 40 };
        btnCheckFile.Click += (s, e) => RunTest(CheckFileTimestamp, btnCheckFile);
        
        btnCheckWindowsUpdate = new Button { Text = "Check Windows Updates", Dock = DockStyle.Top, Height = 40 };
        btnCheckWindowsUpdate.Click += (s, e) => RunTest(CheckWindowsUpdates, btnCheckWindowsUpdate);
        
        btnCheckTeams = new Button { Text = "Check Microsoft Teams", Dock = DockStyle.Top, Height = 40 };
        btnCheckTeams.Click += (s, e) => RunTest(CheckMicrosoftTeams, btnCheckTeams);
        
        btnCheckZoom = new Button { Text = "Check if Zoom is installed", Dock = DockStyle.Top, Height = 40 };
        btnCheckZoom.Click += (s, e) => RunTest(CheckZoom, btnCheckZoom);

        btnOpenZoom = new Button { Text = "Open Zoom", Dock = DockStyle.Top, Height = 40 };
        btnOpenZoom.Click += (s, e) => RunTest(OpenZoom, btnOpenZoom);
        
        btnInstallZoom = new Button { Text = "Install Zoom", Dock = DockStyle.Top, Height = 40 };
        btnInstallZoom.Click += (s, e) => RunTest(InstallZoom, btnInstallZoom);

        btnInstallTeams = new Button { Text = "Install Teams", Dock = DockStyle.Top, Height = 40 };
        btnInstallTeams.Click += (s, e) => RunTest(InstallTeams, btnInstallTeams);

        btnInstallVisualizer = new Button { Text = "Install Visualizer", Dock = DockStyle.Top, Height = 40 };
        btnInstallVisualizer.Click += (s, e) => RunTest(InstallVisualizer, btnInstallVisualizer);

        btnDisableAudioDevices = new Button { Text = "Disable Unwanted Audio Devices [Still in Development]", Dock = DockStyle.Top, Height = 40 };
        btnDisableAudioDevices.Click += (s, e) => RunTest(DisableUnwantedAudioDevices, btnDisableAudioDevices);

        btnEnableAudioDevices = new Button { Text = "Enable All Audio Devices [Still in Development]", Dock = DockStyle.Top, Height = 40 };
        btnEnableAudioDevices.Click += (s, e) => RunTest(EnableAllAudioDevices, btnEnableAudioDevices);

        btnCheckVisualizer = new Button { Text = "Check if Visualizer is installed", Dock = DockStyle.Top, Height = 40 };
        btnCheckVisualizer.Click += (s, e) => RunTest(CheckVisualizer, btnCheckVisualizer);
        
        txtLog = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10) };
        
        progressBar = new ProgressBar { Dock = DockStyle.Bottom, Height = 20, Style = ProgressBarStyle.Marquee, Visible = false };
        
        this.Controls.Add(txtLog);
        this.Controls.Add(progressBar);
        this.Controls.Add(btnInstallZoom);
        this.Controls.Add(btnOpenZoom);
        this.Controls.Add(btnCheckZoom);
        this.Controls.Add(btnInstallTeams);
        this.Controls.Add(btnCheckTeams);
        this.Controls.Add(btnInstallVisualizer);
        this.Controls.Add(btnCheckVisualizer);
        this.Controls.Add(btnCheckWindowsUpdate);
        this.Controls.Add(btnCheckFile);
        this.Controls.Add(btnCheckAVSettings);
        this.Controls.Add(btnDisableAudioDevices);
        this.Controls.Add(btnEnableAudioDevices);

        //this.Controls.Add(btnRunAllTests);
    }

    private void RunTest(Action testMethod, Button btn)
    {
        btn.Enabled = false;
        progressBar.Visible = true;
        Thread testThread = new Thread(() =>
        {
            testMethod();
            Invoke(new Action(() => progressBar.Visible = false));
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
        
        Log($"Default Playback Device: {defaultPlayback.FriendlyName} {(playbackMatch ? "✔" : "✖")}");
        Log($"Default Recording Device: {defaultRecording.FriendlyName} {(recordingMatch ? "✔" : "✖")}");
        Log("\n");
        Log($"Expected Playback Device: {expectedPlayback}");
        Log($"Expected Recording Device: {expectedRecording}");
        Log("\n\n");
    }

 

private void CheckFileTimestamp()
{
    string filePath = GetConfigValue("FileToCheck");
    string filePath2 = GetConfigValue("FileToCheck2");
    if (File.Exists(filePath))
    {
        Process.Start("notepad.exe", filePath);
        Log("Timestamp file opened successfully! ✔");
    }
    else if (File.Exists(filePath2)) {
        Process.Start("notepad.exe", filePath2);
        Log("Timestamp file opened successfully! ✔");
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
        }
        catch (Exception ex)
        {
            Log($"Failed to open Microsoft Teams download page: ✖\nError: {ex.Message}");
        }
        Log("\n\n");
    }
    

    private void InstallVisualizer() {
        string expectedPath = GetConfigValue("VisualizerInstallPath");
        bool isInstalled = File.Exists(expectedPath);
        if (isInstalled) {
            ProcessStartInfo startInfo = new ProcessStartInfo(expectedPath) {
                UseShellExecute = true,
                Verb = "runas" // This specifies to run the process as an administrator
            };
            Process.Start(startInfo);
        } else {
            Log("Could not connect to the network folder: ✖");
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

   private void DisableUnwantedAudioDevices()
    {
        var enumerator = new MMDeviceEnumerator();
        var devices = enumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Active);

        string expectedPlayback = GetConfigValue("ExpectedAudioDevice");
        string expectedRecording = GetConfigValue("ExpectedMicrophone");

        foreach (var device in devices)
        {
            if (device.DataFlow == DataFlow.Render && !device.FriendlyName.Equals(expectedPlayback, StringComparison.OrdinalIgnoreCase))
            {
                DisableDevice(device.ID);
                Log($"Disabled playback device: {device.FriendlyName}");
            }
            else if (device.DataFlow == DataFlow.Capture && !device.FriendlyName.Equals(expectedRecording, StringComparison.OrdinalIgnoreCase))
            {
                DisableDevice(device.ID);
                Log($"Disabled recording device: {device.FriendlyName}");
            }
        }

        SetDefaultAudioDevice(expectedPlayback, DataFlow.Render);
        SetDefaultAudioDevice(expectedRecording, DataFlow.Capture);
    }

    private void EnableAllAudioDevices()
    {
        var enumerator = new MMDeviceEnumerator();
        var devices = enumerator.EnumerateAudioEndPoints(DataFlow.All, DeviceState.Disabled);

        foreach (var device in devices)
        {
            EnableDevice(device.ID);
            Log($"Enabled audio device: {device.FriendlyName}");
        }
    }

    private void SetDefaultAudioDevice(string deviceName, DataFlow dataFlow)
    {
        var enumerator = new MMDeviceEnumerator();
        var devices = enumerator.EnumerateAudioEndPoints(dataFlow, DeviceState.Active);

        foreach (var device in devices)
        {
            if (device.FriendlyName.Equals(deviceName, StringComparison.OrdinalIgnoreCase))
            {
                device.AudioEndpointVolume.Mute = false;
                Log($"Set default {dataFlow} device: {device.FriendlyName}");
                break;
            }
        }
    }

    private void DisableDevice(string deviceId)
    {
        using (var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_PnPEntity WHERE DeviceID = '{deviceId.Replace("\\", "\\\\")}'"))
        {
            foreach (ManagementObject obj in searcher.Get())
            {
                obj.InvokeMethod("Disable", null);
            }
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
}
