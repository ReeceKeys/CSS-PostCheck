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
            + "TeamsInstallPath=\\\\cm\\source\\Microsoft Teams\\Install Teams.cmd\n"
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
    private Button btnCheckTeams, btnCheckZoom, btnOpenZoom, btnOpenCustom, btnUSBDetect;
    private Button btnInstallZoom, btnInstallTeams, btnInstallVisualizer, btnCheckVisualizer;
    private Button btnEnableAudioDevices, btnChangeDisplaySetting, btnGPupdate, btnGroupPolicy, btnOpenApp;
    private bool test_passed = false;
    private string panel_color = "#FFFFFF";
    private string button_color = "#F0FFF0";
    private string main_bg_color = "#FDFFBA";



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


    FlowLayoutPanel sysPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel avPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel zoomTeamsPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel installPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };
    FlowLayoutPanel installTestPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, BackColor = ColorTranslator.FromHtml(panel_color) };

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
    btnInstallTeams = new Button { Text = "Install Microsoft Teams", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnInstallVisualizer = new Button { Text = "Install Visualizer", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckVisualizer = new Button { Text = "Check if Visualizer is Installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckTeams = new Button { Text = "Check if Microsoft Teams is installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnCheckZoom = new Button { Text = "Check if Zoom is Installed", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnChangeDisplaySetting = new Button { Text = "Duplicate Display Settings", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnUSBDetect = new Button { Text = "Detect USB Devices", Width = 180, BackColor = ColorTranslator.FromHtml(button_color) };
    btnChangeDisplaySetting.Click += (s, e) => ChangeDisplay();
    btnOpenApp = new Button 
{ 
    Text = "Open App", 
    Width = 180, 
    BackColor = ColorTranslator.FromHtml(button_color) 
};



    // Event Handlers
    btnCheckWindowsUpdate.Click += (s, e) => RunTest(CheckWindowsUpdates, btnCheckWindowsUpdate);
    btnCheckFile.Click += (s, e) => RunTest(CheckFileTimestamp, btnCheckFile);
    btnCheckAVSettings.Click += (s, e) => RunTest(CheckAVSettings, btnCheckAVSettings);
    btnEnableAudioDevices.Click += (s, e) => RunTest(OpenSoundSettings, btnEnableAudioDevices);
    btnGPupdate.Click += (s, e) => RunTest(GPUpdate, btnGPupdate);
    btnGroupPolicy.Click += (s, e) => RunTest(GPResult, btnGroupPolicy);
    btnUSBDetect.Click += (s, e) => RunTest(DetectUSBDevices, btnUSBDetect);
    btnOpenApp.Click += (s, e) => ShowAppSelectionDialog();

    btnCheckTeams.Click += (s, e) => RunTest(CheckMicrosoftTeams, btnCheckTeams);
    btnCheckZoom.Click += (s, e) => RunTest(CheckZoom, btnCheckZoom);
    btnOpenZoom.Click += (s, e) => RunTest(OpenZoom, btnOpenZoom);
    btnOpenCustom.Click += (s, e) => RunTest(OpenCustom, btnOpenCustom);


    btnInstallZoom.Click += (s, e) => RunTest(InstallZoom, btnInstallZoom);
    btnInstallTeams.Click += (s, e) => RunTest(InstallTeams, btnInstallTeams);
    btnInstallVisualizer.Click += (s, e) => RunTest(InstallVisualizer, btnInstallVisualizer);
    btnCheckVisualizer.Click += (s, e) => RunTest(CheckVisualizer, btnCheckVisualizer);

    // Add buttons to respective panels
    sysPanel.Controls.AddRange(new Control[] { btnCheckWindowsUpdate, btnCheckFile, btnGPupdate, btnGroupPolicy });
    avPanel.Controls.AddRange(new Control[] { btnCheckAVSettings, btnEnableAudioDevices, btnChangeDisplaySetting, btnUSBDetect  });
    zoomTeamsPanel.Controls.AddRange(new Control[] { btnOpenApp, btnOpenCustom });
    installPanel.Controls.AddRange(new Control[] { btnInstallZoom, btnInstallTeams, btnInstallVisualizer });
    installTestPanel.Controls.AddRange(new Control[] { btnCheckZoom, btnCheckTeams, btnCheckVisualizer });


    // Add panels to group boxes
    grpSystemTests.Controls.Add(sysPanel);
    grpAVTests.Controls.Add(avPanel);
    grpZoomTeamsTests.Controls.Add(zoomTeamsPanel);
    grpInstallTests.Controls.Add(installPanel);
    grpInstallTests2.Controls.Add(installTestPanel);

    // Log and progress bar
    txtLog = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Font = new Font("Consolas", 10) };
    txtLog.BackColor = Color.White;  
    progressBar = new ProgressBar { Dock = DockStyle.Bottom, Height = 20, Style = ProgressBarStyle.Marquee, Visible = false };

    // Add components to main panel
    mainPanel.Controls.Add(txtLog);
    mainPanel.Controls.Add(progressBar);
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
        btn.Enabled = false;  // Disable button to prevent multiple rapid clicks
        progressBar.Visible = true;
        test_passed = false; 

        Thread testThread = new Thread(() =>
        {
            testMethod();

            // Re-enable button and hide progress bar after test completes (if it doesn't pass the test)
            Invoke(new Action(() =>
            {
                if (!test_passed) 
                    btn.Enabled = true;
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
        DetectCameraDevice();
        Log("\n\n");
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

    private void DetectCameraDevice()
{
    List<string> cameraNames = new List<string> {"AVer TR313V2", "AVer TR313"};
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

        Log("No matching camera device found. ✖");
    }
    catch (Exception e)
    {
        Log($"Error detecting camera: {e.Message}");
    }
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
                Log("Initialized Zoom install: ✔");
            } catch (System.ComponentModel.Win32Exception) {
                Log("The user declined the elevation request.");
            }
        } else {
            Log("Could not connect to the network folder: ✖");
        }
        Log("\n\n");
    }

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
                Log("The user declined the elevation request.");
            }
        } else {
            Log("Could not connect to the network folder: ✖");
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
            Log("Initialized Zoom install: ✔");
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
        Button btnTeams = new Button { Text = "Teams", Width = 120 };
        Button btnVisualizer = new Button { Text = "Visualizer", Width = 120 };
        Button btnCougarApps = new Button { Text = "Cougar Apps", Width = 120 };

        // Assign click events with installation checks
        btnZoom.Click += (s, e) => CheckAndOpenApp("Zoom", GetConfigValue("ExpectedZoomPath"), InstallZoom);
        btnTeams.Click += (s, e) => CheckAndOpenApp("Microsoft Teams", GetConfigValue("ExpectedTeamsPath"), InstallTeams);
        btnVisualizer.Click += (s, e) => CheckAndOpenApp("Visualizer", GetConfigValue("ExpectedVisualizerPath"), InstallVisualizer);
        btnCougarApps.Click += (s, e) => OpenWebsite("https://cougarapps.csusm.edu/Citrix/CougarAppsWeb/");

        // Add buttons to panel
        panel.Controls.AddRange(new Control[] { btnZoom, btnTeams, btnVisualizer, btnCougarApps });

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
}



private void OpenApplication(string appPath)
{
    try
    {
        System.Diagnostics.Process.Start(appPath);
    }
    catch (Exception ex)
    {
        MessageBox.Show("Error opening application: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

private void OpenWebsite(string url)
{
    try
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }
    catch (Exception ex)
    {
        MessageBox.Show("Error opening website: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
