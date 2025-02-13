# CSUSM Classroom Systems Support IT Sweeps Test Script

## Description

This script performs the following checks and installations:

- **ExpectedAudioDevice**: Verifies that the expected audio device (ExtronScalerD) is installed and configured.
- **ExpectedMicrophone**: Checks for the presence of the expected microphone (COLLABORATE Versa USB Input).
- **FileToCheck**: Ensures the existence of the `timestamp.txt` file at the specified path.
- **FileToCheck2**: Ensures the existence of the `timestamp_Instruct.txt` file at the specified path.
- **ExpectedZoomPath**: Verifies that Zoom is installed at the specified path.
- **ExpectedTeamsPath**: Verifies that Microsoft Teams is installed at the specified path.
- **ExpectedVisualizerPath**: Verifies that IPEVO Visualizer is installed at the specified path.
- **ZoomInstallPath**: Provides the network path to the Zoom installation script.
- **TeamsInstallPath**: Provides the URL to download the Microsoft Teams installer.
- **VisualizerInstallPath**: Provides the network path to the IPEVO Visualizer installer.

## Usage

1. **Setup**: Ensure the `config.ini` file is correctly configured with the expected values.
2. **Run the Script**: Execute the test script on the target Windows 10/11 machines.
3. **Verify Results**: Check the output to ensure all tests pass based on the configuration settings.

## Installation

To install the necessary software, follow these steps:

1. **Zoom**: Run the installation script located at `\\cm\source\Zoom\Zoom CFR\InstallsZoomClientHDEnabled.bat`.
2. **Microsoft Teams**: Download and install Teams from this link.
3. **IPEVO Visualizer**: Run the installer located at `\\cm\source\IPEVO Presenter\Visualizer\Visualizer_win7-11_v3.6.4.1.msi`.

## Tests

The script performs the following tests:

- **Audio Device Check**: Confirms the presence of the expected audio device.
- **Microphone Check**: Confirms the presence of the expected microphone.
- **File Existence Check**: Verifies the existence of specified files.
- **Software Installation Check**: Ensures that Zoom, Teams, and Visualizer are installed at the expected paths.

## Known Issues

- The functionality for enabling/disabling audio devices is not yet complete.
- Installation buttons for Visualizer and Teams are not yet complete.

## Contributing

We welcome contributions to improve this script. Please fork the repository and submit a pull request with your changes. Ensure your code adheres to our coding standards and includes appropriate tests.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

## Contact

For any issues or questions, please contact the CSUSM Classroom Systems Support team at support@csusm.edu.
