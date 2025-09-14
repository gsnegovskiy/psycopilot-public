#!/usr/bin/env pwsh
# psycopilot-installer-windows.ps1
# Windows installer for PsycoPilot - Real-time dual-channel audio transcription
# 
# Usage:
#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
#   curl -fsSL https://raw.githubusercontent.com/gsnegovskiy/psycopilot-public/master/psycopilot-installer-windows.ps1 | pwsh
#
# Options:
#   -TestAudioOnly          Only test audio setup, skip full installation
#   -InstallVirtualAudio    Install VB-Cable virtual audio device
#   -EnableWSL              Enable Windows Subsystem for Linux
#   -PythonVersion          Specify Python version (default: 3.13)
#   -InstallDir             Installation directory (default: $env:USERPROFILE\psycopilot)
#   -Force                  Force overwrite existing installation
#
# Examples:
#   .\psycopilot-installer-windows.ps1 -TestAudioOnly
#   .\psycopilot-installer-windows.ps1 -InstallVirtualAudio
#   .\psycopilot-installer-windows.ps1 -EnableWSL

param(
    [switch]$EnableWSL,
    [switch]$InstallVirtualAudio,
    [switch]$TestAudioOnly,
    [string]$PythonVersion = "3.13",
    [string]$InstallDir = "$env:USERPROFILE\psycopilot",
    [switch]$Force
)

# Configuration
$APP_VERSION = "1.0.0"
$GITHUB_REPO = "gsnegovskiy/psycopilot-public"
$MIN_WINDOWS_VERSION = "10.0.19041"  # Windows 10 2004 or later

# Colors for output
$Colors = @{
    Red = "Red"
    Green = "Green"
    Yellow = "Yellow"
    Blue = "Blue"
    Cyan = "Cyan"
    White = "White"
}

# Utility functions
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Colors[$Color]
}

function Write-Success {
    param([string]$Message)
    Write-ColorOutput "[✓] $Message" "Green"
}

function Write-Warning {
    param([string]$Message)
    Write-ColorOutput "[!] $Message" "Yellow"
}

function Write-Error {
    param([string]$Message)
    Write-ColorOutput "[✗] $Message" "Red"
}

function Write-Info {
    param([string]$Message)
    Write-ColorOutput "[i] $Message" "Cyan"
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-Banner {
    Clear-Host
    Write-ColorOutput @"
     ____                       ____  _ __      __
   / __ \____  __  _________  / __ \(_) /___  / /_
  / /_/ / __ \/ / / / ___/ _ \/ /_/ / / / __ \/ __/
 / ____/ /_/ / /_/ / /__/  __/ ____/ / / /_/ / /_
/_/   \____/\__, /\___/\___/_/   /_/_/\____/\__/
           /____/

Real-time dual-channel audio transcription for therapy sessions
"@ "Cyan"
    Write-ColorOutput "Version: $APP_VERSION" "White"
    Write-ColorOutput "Installing to: $InstallDir" "White"
    Write-ColorOutput "Repository: $GITHUB_REPO" "White"
    Write-Host ""
}

function Test-SystemRequirements {
    Write-Info "Checking system requirements..."
    
    # Check Windows version
    $osVersion = [System.Environment]::OSVersion.Version
    $osVersionString = "$($osVersion.Major).$($osVersion.Minor).$($osVersion.Build)"
    
    if ([System.Version]$osVersionString -lt [System.Version]$MIN_WINDOWS_VERSION) {
        Write-Error "Windows 10 version 2004 or later required. Found: $osVersionString"
        exit 1
    }
    Write-Success "Windows version compatible: $osVersionString"
    
    # Check architecture
    $arch = $env:PROCESSOR_ARCHITECTURE
    if ($arch -eq "AMD64") {
        Write-Success "64-bit architecture detected"
    } else {
        Write-Warning "32-bit architecture detected. Some features may not work correctly."
    }
    
    # Check available disk space (require at least 2GB)
    $drive = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = [math]::Round($drive.FreeSpace / 1GB, 2)
    
    if ($freeSpaceGB -lt 2) {
        Write-Error "Insufficient disk space. At least 2GB required, found ${freeSpaceGB}GB"
        exit 1
    }
    Write-Success "Sufficient disk space available: ${freeSpaceGB}GB"
    
    # Check if running as administrator
    if (-not (Test-Administrator)) {
        Write-Warning "Not running as administrator. Some features may require elevation."
    } else {
        Write-Success "Running with administrator privileges"
    }
}

function Install-Chocolatey {
    Write-Info "Installing Chocolatey package manager..."
    
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Success "Chocolatey already installed"
        return
    }
    
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-Success "Chocolatey installed successfully"
    } catch {
        Write-Error "Failed to install Chocolatey: $_"
        exit 1
    }
}

function Install-Python {
    Write-Info "Installing Python $PythonVersion..."
    
    try {
        # Install Python via Chocolatey
        choco install python --version=$PythonVersion -y
        Write-Success "Python $PythonVersion installed"
        
        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        
        # Verify installation
        $pythonVersion = python --version 2>&1
        Write-Success "Python version: $pythonVersion"
        
    } catch {
        Write-Error "Failed to install Python: $_"
        exit 1
    }
}

function Install-VisualCppRedistributables {
    Write-Info "Installing Visual C++ Redistributables..."
    
    try {
        choco install vcredist-all -y
        Write-Success "Visual C++ Redistributables installed"
    } catch {
        Write-Warning "Failed to install Visual C++ Redistributables: $_"
        Write-Warning "Some Python packages may not work correctly"
    }
}

function Install-Git {
    Write-Info "Installing Git..."
    
    try {
        choco install git -y
        Write-Success "Git installed"
    } catch {
        Write-Warning "Failed to install Git: $_"
    }
}

function Enable-WSL {
    Write-Info "Enabling Windows Subsystem for Linux (WSL)..."
    
    try {
        # Enable WSL feature
        dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart
        dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart
        
        Write-Success "WSL features enabled"
        Write-Warning "Please restart your computer and run: wsl --install"
        Write-Warning "Then run this installer again to continue with WSL setup"
        
    } catch {
        Write-Error "Failed to enable WSL: $_"
        exit 1
    }
}

function Install-VirtualAudio {
    Write-Info "Installing VB-Cable virtual audio device..."
    
    try {
        # Create temp directory
        $tempDir = Join-Path $env:TEMP "psycopilot-setup"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Download VB-Cable
        $vbcableUrl = "https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack43.zip"
        $vbcableZip = Join-Path $tempDir "vbcable.zip"
        
        Write-Info "Downloading VB-Cable..."
        Invoke-WebRequest -Uri $vbcableUrl -OutFile $vbcableZip
        
        # Extract and install
        Write-Info "Extracting VB-Cable..."
        Expand-Archive -Path $vbcableZip -DestinationPath $tempDir -Force
        
        $installerPath = Join-Path $tempDir "VBCABLE_Setup_x64.exe"
        if (Test-Path $installerPath) {
            Write-Info "Installing VB-Cable (this may require user interaction)..."
            Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait
            Write-Success "VB-Cable installed successfully"
        } else {
            Write-Warning "VB-Cable installer not found. Please install manually from: https://vb-audio.com/Cable/"
        }
        
        # Cleanup
        Remove-Item -Path $tempDir -Recurse -Force
        
    } catch {
        Write-Warning "Failed to install VB-Cable: $_"
        Write-Warning "Please install manually from: https://vb-audio.com/Cable/"
    }
}

function Test-AudioDevices {
    Write-Info "Testing audio device configuration..."
    
    try {
        # Check for VB-Cable
        $vbcableDevice = Get-WmiObject -Class Win32_SoundDevice | Where-Object { $_.Name -like "*VB-Cable*" }
        if ($vbcableDevice) {
            Write-Success "VB-Cable detected: $($vbcableDevice.Name)"
        } else {
            Write-Warning "VB-Cable not detected. System audio capture may not work."
        }
        
        # Check for Stereo Mix
        $stereoMixDevice = Get-WmiObject -Class Win32_SoundDevice | Where-Object { $_.Name -like "*Stereo Mix*" }
        if ($stereoMixDevice) {
            Write-Success "Stereo Mix detected: $($stereoMixDevice.Name)"
        } else {
            Write-Warning "Stereo Mix not detected. You may need to enable it manually."
        }
        
        # Check for microphone
        $micDevice = Get-WmiObject -Class Win32_SoundDevice | Where-Object { $_.Name -like "*Microphone*" -or $_.Name -like "*Mic*" }
        if ($micDevice) {
            Write-Success "Microphone detected: $($micDevice.Name)"
        } else {
            Write-Warning "No microphone detected. Please connect a microphone."
        }
        
    } catch {
        Write-Warning "Failed to test audio devices: $_"
    }
}

function Enable-StereoMix {
    Write-Info "Attempting to enable Stereo Mix..."
    
    try {
        # This requires registry modification to enable Stereo Mix
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
        
        # Get all audio render devices
        $renderDevices = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
        
        foreach ($device in $renderDevices) {
            $devicePath = Join-Path $regPath $device.PSChildName
            $propertiesPath = Join-Path $devicePath "Properties"
            
            # Check if this is a Stereo Mix device
            $deviceName = Get-ItemProperty -Path $propertiesPath -Name "{a45c254e-df1c-4efd-8020-67d146a850e0},2" -ErrorAction SilentlyContinue
            if ($deviceName -and $deviceName."{a45c254e-df1c-4efd-8020-67d146a850e0},2" -like "*Stereo Mix*") {
                # Enable the device
                $deviceStatePath = Join-Path $devicePath "DeviceState"
                Set-ItemProperty -Path $deviceStatePath -Name "DeviceState" -Value 1 -ErrorAction SilentlyContinue
                Write-Success "Stereo Mix enabled"
                return
            }
        }
        
        Write-Warning "Stereo Mix not found or could not be enabled automatically"
        Write-Warning "Please enable it manually in Windows Sound settings"
        
    } catch {
        Write-Warning "Failed to enable Stereo Mix: $_"
        Write-Warning "Please enable it manually in Windows Sound settings"
    }
}

function Show-AudioSetupInstructions {
    Write-Host ""
    Write-ColorOutput "🔊 Audio Setup Instructions:" "Cyan"
    Write-Host ""
    Write-ColorOutput "For system audio capture, you need one of the following:" "White"
    Write-ColorOutput "1. VB-Cable (recommended) - Virtual audio cable" "White"
    Write-ColorOutput "2. Stereo Mix - Built-in Windows feature" "White"
    Write-ColorOutput "3. VoiceMeeter - Advanced virtual audio mixer" "White"
    Write-Host ""
    Write-ColorOutput "To enable Stereo Mix manually:" "Yellow"
    Write-ColorOutput "1. Right-click speaker icon in system tray" "White"
    Write-ColorOutput "2. Select 'Open Sound settings'" "White"
    Write-ColorOutput "3. Click 'Sound Control Panel'" "White"
    Write-ColorOutput "4. Go to 'Recording' tab" "White"
    Write-ColorOutput "5. Right-click empty space and select 'Show Disabled Devices'" "White"
    Write-ColorOutput "6. Right-click 'Stereo Mix' and select 'Enable'" "White"
    Write-Host ""
    Write-ColorOutput "For microphone capture:" "Yellow"
    Write-ColorOutput "1. Ensure your microphone is connected and working" "White"
    Write-ColorOutput "2. Test microphone in Windows Sound settings" "White"
    Write-ColorOutput "3. Check microphone permissions in Windows Privacy settings" "White"
    Write-Host ""
}

function Test-AudioSetup {
    Write-Info "Testing audio setup..."
    
    try {
        # Test if we can access audio devices
        $audioDevices = Get-WmiObject -Class Win32_SoundDevice
        $inputDevices = $audioDevices | Where-Object { $_.Name -like "*Microphone*" -or $_.Name -like "*Mic*" -or $_.Name -like "*Stereo Mix*" -or $_.Name -like "*VB-Cable*" }
        
        if ($inputDevices.Count -gt 0) {
            Write-Success "Audio setup appears to be working"
            Write-Info "Found $($inputDevices.Count) input devices:"
            foreach ($device in $inputDevices) {
                Write-Info "  - $($device.Name)"
            }
        } else {
            Write-Warning "No suitable input devices found"
            Show-AudioSetupInstructions
        }
        
    } catch {
        Write-Warning "Failed to test audio setup: $_"
        Show-AudioSetupInstructions
    }
}

function Setup-PythonEnvironment {
    Write-Info "Setting up Python environment..."
    
    try {
        # Create installation directory
        if (Test-Path $InstallDir) {
            if ($Force) {
                Write-Warning "Removing existing installation directory..."
                Remove-Item -Path $InstallDir -Recurse -Force
            } else {
                Write-Error "Installation directory already exists: $InstallDir"
                Write-Error "Use -Force to overwrite or choose a different directory"
                exit 1
            }
        }
        
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        Set-Location $InstallDir
        
        # Create virtual environment
        Write-Info "Creating Python virtual environment..."
        python -m venv .venv
        
        # Activate virtual environment
        $activateScript = Join-Path $InstallDir ".venv\Scripts\Activate.ps1"
        & $activateScript
        
        # Upgrade pip
        Write-Info "Upgrading pip..."
        python -m pip install --upgrade pip wheel setuptools
        
        Write-Success "Python environment created"
        
    } catch {
        Write-Error "Failed to setup Python environment: $_"
        exit 1
    }
}

function Install-PythonDependencies {
    Write-Info "Installing Python dependencies..."
    
    try {
        # Create requirements.txt
        $requirements = @"
setuptools<81

# numpy per Python
numpy==2.1.1; python_version >= "3.13"
numpy==1.26.4; python_version < "3.13"

sounddevice==0.4.7
webrtcvad==2.0.10

# Whisper via whisper.cpp (much faster, GPU accelerated, no PyAV dependency)
pywhispercpp

# Insanely Fast Whisper (GPU-accelerated with Flash Attention 2)
transformers>=4.44.0
optimum>=1.20.0
accelerate>=0.30.0
torch>=2.0.0

# GigaAM via ONNX Runtime
onnxruntime==1.22.1
onnx-asr[cpu,hub]==0.6.1

# HF client
huggingface-hub==0.25.2

# Web UI dependencies
flask>=2.3.0
requests>=2.31.0
"@
        
        $requirements | Out-File -FilePath "requirements.txt" -Encoding UTF8
        
        # Install dependencies
        Write-Info "Installing core dependencies (this may take several minutes)..."
        pip install -r requirements.txt
        
        Write-Success "Python dependencies installed"
        
    } catch {
        Write-Error "Failed to install Python dependencies: $_"
        exit 1
    }
}

function Download-PsycoPilot {
    Write-Info "Downloading PsycoPilot application..."
    
    try {
        # Clone or download the repository
        if (Get-Command git -ErrorAction SilentlyContinue) {
            Write-Info "Cloning repository..."
            git clone https://github.com/$GITHUB_REPO.git .
        } else {
            Write-Info "Downloading repository as ZIP..."
            $zipUrl = "https://github.com/$GITHUB_REPO/archive/refs/heads/master.zip"
            $zipFile = Join-Path $InstallDir "psycopilot.zip"
            
            Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile
            Expand-Archive -Path $zipFile -DestinationPath $InstallDir -Force
            
            # Move contents from subdirectory
            $extractedDir = Join-Path $InstallDir "psycopilot-public-master"
            Get-ChildItem -Path $extractedDir | Move-Item -Destination $InstallDir
            Remove-Item -Path $extractedDir -Recurse -Force
            Remove-Item -Path $zipFile -Force
        }
        
        Write-Success "PsycoPilot application downloaded"
        
    } catch {
        Write-Error "Failed to download PsycoPilot: $_"
        exit 1
    }
}

function Create-StartScripts {
    Write-Info "Creating start scripts..."
    
    try {
        # Create start script for PowerShell
        $startScript = @"
#!/usr/bin/env pwsh
# Start PsycoPilot on Windows

param(
    [string]`$Engine = "whisper",
    [string]`$Language = "en",
    [string]`$Model = "base"
)

`$scriptDir = Split-Path -Parent `$MyInvocation.MyCommand.Path
Set-Location `$scriptDir

# Activate virtual environment
& ".venv\Scripts\Activate.ps1"

# Start the application
Write-Host "Starting PsycoPilot..." -ForegroundColor Green
Write-Host "Engine: `$Engine" -ForegroundColor Cyan
Write-Host "Language: `$Language" -ForegroundColor Cyan
Write-Host "Model: `$Model" -ForegroundColor Cyan
Write-Host ""

python application/audio-capture/transcribe_dual_database.py --engine `$Engine --language `$Language --model `$Model
"@
        
        $startScript | Out-File -FilePath "start-psycopilot.ps1" -Encoding UTF8
        
        # Create batch file for easy double-click
        $batchScript = @"
@echo off
echo Starting PsycoPilot...
powershell -ExecutionPolicy Bypass -File "%~dp0start-psycopilot.ps1"
pause
"@
        
        $batchScript | Out-File -FilePath "start-psycopilot.bat" -Encoding ASCII
        
        Write-Success "Start scripts created"
        
    } catch {
        Write-Error "Failed to create start scripts: $_"
    }
}

function Show-CompletionMessage {
    Write-Host ""
    Write-ColorOutput "🎉 Installation completed successfully!" "Green"
    Write-Host ""
    Write-ColorOutput "Next steps:" "Cyan"
    Write-ColorOutput "1. Ensure audio devices are properly configured" "White"
    Write-ColorOutput "2. Test audio setup by running: .\start-psycopilot.ps1" "White"
    Write-ColorOutput "3. If audio issues occur, check the troubleshooting guide below" "White"
    Write-Host ""
    Write-ColorOutput "🔧 Troubleshooting Audio Issues:" "Yellow"
    Write-ColorOutput "• If system audio capture doesn't work:" "White"
    Write-ColorOutput "  - Install VB-Cable from: https://vb-audio.com/Cable/" "White"
    Write-ColorOutput "  - Or enable Stereo Mix in Windows Sound settings" "White"
    Write-ColorOutput "• If microphone capture doesn't work:" "White"
    Write-ColorOutput "  - Check microphone permissions in Windows Privacy settings" "White"
    Write-ColorOutput "  - Test microphone in Windows Sound settings" "White"
    Write-ColorOutput "• For advanced audio setup, consider VoiceMeeter" "White"
    Write-Host ""
    Write-ColorOutput "📚 Documentation:" "Cyan"
    Write-ColorOutput "• GitHub: https://github.com/$GITHUB_REPO" "White"
    Write-ColorOutput "• Windows Audio Setup Guide: See README.md" "White"
    Write-Host ""
}

# Main installation process
function Main {
    try {
        Show-Banner
        
        # If only testing audio, skip system requirements and installation
        if ($TestAudioOnly) {
            Write-ColorOutput "🔊 Audio Testing Mode" "Cyan"
            Write-Host ""
            Test-AudioDevices
            Enable-StereoMix
            Test-AudioSetup
            Show-AudioSetupInstructions
            return
        }
        
        Test-SystemRequirements
        
        # Install system dependencies
        Install-Chocolatey
        Install-Python
        Install-VisualCppRedistributables
        Install-Git
        
        # Optional installations
        if ($EnableWSL) {
            Enable-WSL
            return  # Exit after WSL setup, user needs to restart
        }
        
        if ($InstallVirtualAudio) {
            Install-VirtualAudio
        }
        
        # Setup application
        Setup-PythonEnvironment
        Download-PsycoPilot
        Install-PythonDependencies
        Create-StartScripts
        
        # Audio setup and testing
        Write-Host ""
        Write-ColorOutput "🔊 Audio Setup and Testing" "Cyan"
        Test-AudioDevices
        Enable-StereoMix
        Test-AudioSetup
        
        Show-CompletionMessage
        
    } catch {
        Write-Error "Installation failed: $_"
        exit 1
    }
}

# Run main function
Main
