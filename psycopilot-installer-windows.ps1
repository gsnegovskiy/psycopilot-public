#!/usr/bin/env pwsh
# psycopilot-installer-windows.ps1
# Windows installer for PsycoPilot - Real-time dual-channel audio transcription
# 
# Usage:
#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
#   $env:GITHUB_TOKEN="your_token_here"
#   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/gsnegovskiy/psycopilot-public/master/psycopilot-installer-windows.ps1" | Invoke-Expression
#
# Options:
#   -TestAudioOnly          Only test audio setup, skip full installation
#   -InstallVirtualAudio    Install VB-Cable virtual audio device
#   -EnableWSL              Enable Windows Subsystem for Linux
#   -PythonVersion          Specify Python version (default: 3.13)
#   -InstallDir             Installation directory (default: $env:USERPROFILE\psycopilot)
#   -GitHubToken            GitHub personal access token for private repository access
#   -Force                  Force overwrite existing installation
#
# Examples:
#   $env:GITHUB_TOKEN="ghp_xxxxxxxxxxxx"; Invoke-WebRequest -Uri "https://raw.githubusercontent.com/gsnegovskiy/psycopilot-public/master/psycopilot-installer-windows.ps1" | Invoke-Expression
#   .\psycopilot-installer-windows.ps1 -TestAudioOnly
#   $env:GITHUB_TOKEN="ghp_xxxxxxxxxxxx"; .\psycopilot-installer-windows.ps1 -InstallVirtualAudio
#   .\psycopilot-installer-windows.ps1 -EnableWSL

param(
    [switch]$EnableWSL,
    [switch]$InstallVirtualAudio,
    [switch]$TestAudioOnly,
    [string]$PythonVersion = "3.13",
    [string]$InstallDir = "$env:USERPROFILE\psycopilot",
    [string]$GitHubToken
)

# Configuration
$APP_VERSION = "1.0.0"
$GITHUB_REPO = "gsnegovskiy/psycopilot"
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

function Wait-ForUserInput {
    param(
        [string]$Message = "Press any key to continue...",
        [string]$Color = "White"
    )
    
    try {
        Write-ColorOutput $Message $Color
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Write-ColorOutput "Press Enter to continue..." $Color
        Read-Host
    }
}

function Test-InstallationDirectory {
    Write-Info "Checking installation directory..."
    
    if (Test-Path $InstallDir) {
        Write-Info "Installation directory already exists: $InstallDir"
        Write-Info "Contents:"
        Get-ChildItem $InstallDir | ForEach-Object { Write-Info "  - $($_.Name)" }
        Write-Info "Continuing with existing directory..."
    } else {
        Write-Info "Creating installation directory: $InstallDir"
        try {
            New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
            Write-Success "Installation directory created: $InstallDir"
        } catch {
            Write-Error "Failed to create installation directory: $_"
            exit 1
        }
    }
}

function Refresh-EnvironmentPath {
    Write-Info "Refreshing environment PATH..."
    try {
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        Start-Sleep -Seconds 1
        Write-Success "Environment PATH refreshed"
    } catch {
        Write-Warning "Failed to refresh PATH: $_"
    }
}

function Test-PowerShellCompatibility {
    Write-Info "Checking PowerShell compatibility..."
    
    $psVersion = $PSVersionTable.PSVersion
    Write-Info "PowerShell version: $($psVersion.ToString())"
    
    # Check if we're running on Windows PowerShell or PowerShell Core
    if ($PSVersionTable.PSEdition -eq "Desktop") {
        Write-Success "Running on Windows PowerShell (Desktop edition) - Supported"
    } else {
        Write-Success "Running on PowerShell Core (Cross-platform edition) - Supported"
    }
    
    # Check minimum PowerShell version (5.1 for Windows PowerShell, 6.0+ for Core)
    $minVersion = if ($PSVersionTable.PSEdition -eq "Desktop") { [Version]"5.1" } else { [Version]"6.0" }
    
    if ($psVersion -lt $minVersion) {
        Write-Error "PowerShell version $minVersion or later required. Found: $($psVersion.ToString())"
        exit 1
    }
    Write-Success "PowerShell version compatible: $($psVersion.ToString())"
}

function Test-SystemRequirements {
    Write-Info "Checking system requirements..."
    
    # Check PowerShell compatibility first
    Test-PowerShellCompatibility
    
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

function Test-GitHubToken {
    Write-Info "Validating GitHub token..."
    
    # Check for token in parameter first, then environment variable
    if (-not $GitHubToken) {
        $GitHubToken = $env:GITHUB_TOKEN
    }
    
    if (-not $GitHubToken) {
        Write-Error "GitHub token is required for private repository access."
        Write-Error "Please provide a GitHub personal access token with repository access."
        Write-Error ""
        Write-Error "Option 1 - Environment variable:"
        Write-Error "  `$env:GITHUB_TOKEN='ghp_xxxxxxxxxxxx'"
        Write-Error "  .\psycopilot-installer-windows.ps1"
        Write-Error ""
        Write-Error "Option 2 - Parameter:"
        Write-Error "  .\psycopilot-installer-windows.ps1 -GitHubToken 'ghp_xxxxxxxxxxxx'"
        Write-Error ""
        Write-Error "To get a token ask Greg for it"
        exit 1
    }
    
    # Validate token format
    if (-not ($GitHubToken -match "^ghp_[A-Za-z0-9]{36}$" -or $GitHubToken -match "^github_pat_[A-Za-z0-9_]{82}$")) {
        Write-Warning "GitHub token format appears invalid. Expected format: ghp_xxxxxxxxxxxx or github_pat_xxxxxxxxxxxx"
        Write-Warning "Continuing anyway, but authentication may fail..."
    }

    # Test token by making a request to GitHub API
    try {
        $headers = @{
            "Authorization" = "token $GitHubToken"
            "Accept" = "application/vnd.github.v3+json"
        }
        
        $response = Invoke-RestMethod -Uri "https://api.github.com/user" -Headers $headers -Method Get
        Write-Success "GitHub token validated for user: $($response.login)"

    } catch {
        Write-Error "GitHub token validation failed: $_"
        Write-Error "Please check your token and ensure it has repository access permissions."
        exit 1
    }
}

function Invoke-ChocolateyCommand {
    param(
        [string]$Command,
        [string]$Arguments = ""
    )
    
    # Try to use choco command first
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Info "Using choco command: choco $Command $Arguments"
        $result = & choco $Command $Arguments 2>&1
        return $result
    }
    
    # Fallback to full path
    $chocoExe = "C:\ProgramData\chocolatey\bin\choco.exe"
    if (Test-Path $chocoExe) {
        Write-Info "Using full path to choco.exe: $chocoExe $Command $Arguments"
        $result = & $chocoExe $Command $Arguments 2>&1
        return $result
    }
    
    throw "Chocolatey not found. Please ensure Chocolatey is installed."
}

function Install-Chocolatey {
    Write-Info "Installing Chocolatey package manager..."
    
    # Check if Chocolatey is already available
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Success "Chocolatey already installed"
        $chocoVersion = & choco --version
        Write-Info "Chocolatey version: $chocoVersion"
        return
    }
    
    # Check if choco.exe exists but not in PATH
    $chocoExe = "C:\ProgramData\chocolatey\bin\choco.exe"
    if (Test-Path $chocoExe) {
        Write-Info "Found existing Chocolatey installation at: $chocoExe"
        Write-Info "Adding to PATH and creating alias..."
        $chocoPath = "C:\ProgramData\chocolatey\bin"
        $env:Path = "$chocoPath;$env:Path"
        Set-Alias -Name choco -Value $chocoExe
        Write-Success "Chocolatey available via existing installation"
        return
    }
    
    try {
        Write-Info "Downloading and installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        
        # Install Chocolatey using official documentation command
        Write-Info "Executing Chocolatey installation script..."
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Wait for installation to complete
        Write-Info "Waiting for Chocolatey installation to complete..."
        Start-Sleep -Seconds 10
        
        # Check if installation was successful
        if (Test-Path $chocoExe) {
            Write-Success "Chocolatey installation completed successfully"
        } else {
            throw "Chocolatey installation failed - choco.exe not found after installation"
        }
        
        # Add Chocolatey to PATH for this session
        $chocoPath = "C:\ProgramData\chocolatey\bin"
        Write-Info "Adding Chocolatey to PATH: $chocoPath"
        $env:Path = "$chocoPath;$env:Path"
        
        # Create alias for this session
        Set-Alias -Name choco -Value $chocoExe
        Write-Info "Created PowerShell alias for choco command"
        
        # Wait for PATH to take effect
        Start-Sleep -Seconds 3
        
        # Verify Chocolatey is working
        try {
            $chocoVersion = & choco --version
            Write-Success "Chocolatey installed and working - version: $chocoVersion"
        } catch {
            Write-Warning "Chocolatey installed but version check failed: $_"
            Write-Info "Continuing with installation - choco command should be available"
        }
        
    } catch {
        Write-Error "Failed to install Chocolatey: $_"
        Write-Error "Error details: $($_.Exception.Message)"
        Write-Error "Please try running the installer as Administrator or install Chocolatey manually"
        exit 1
    }
}

function Install-Python {
    Write-Info "Installing Python $PythonVersion..."
    
    try {
        # Check if Python is already installed
        if (Get-Command python -ErrorAction SilentlyContinue) {
            $existingVersion = python --version 2>&1
            Write-Success "Python already installed: $existingVersion"
            return
        }
        
        # Install Python via Chocolatey
        Write-Info "Installing Python $PythonVersion via Chocolatey..."
        $result = Invoke-ChocolateyCommand -Command "install" -Arguments "python --version=$PythonVersion -y"
        Write-Info "Chocolatey output: $result"
        
        if ($LASTEXITCODE -ne 0) {
            throw "Chocolatey installation failed with exit code: $LASTEXITCODE"
        }
        
        Write-Success "Python $PythonVersion installed"
        
        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        
        # Verify installation
        Start-Sleep -Seconds 2  # Give time for PATH to update
        $pythonVersion = python --version 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "Python installation verification failed: $pythonVersion"
        }
        Write-Success "Python version: $pythonVersion"
        
    } catch {
        Write-Error "Failed to install Python: $_"
        Write-Error "Last exit code: $LASTEXITCODE"
        Write-Error "Please check Chocolatey installation and try again"
        exit 1
    }
}

function Install-VisualCppRedistributables {
    Write-Info "Installing Visual C++ Redistributables..."
    
    try {
        Write-Info "Installing Visual C++ Redistributables via Chocolatey..."
        $result = Invoke-ChocolateyCommand -Command "install" -Arguments "vcredist-all -y"
        Write-Info "Chocolatey output: $result"
        
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Visual C++ Redistributables installation failed with exit code: $LASTEXITCODE"
            Write-Warning "Some Python packages may not work correctly"
        } else {
            Write-Success "Visual C++ Redistributables installed"
        }
    } catch {
        Write-Warning "Failed to install Visual C++ Redistributables: $_"
        Write-Warning "Some Python packages may not work correctly"
    }
}

function Install-Git {
    Write-Info "Installing Git..."
    
    try {
        # Check if Git is already installed
        if (Get-Command git -ErrorAction SilentlyContinue) {
            $existingVersion = git --version 2>&1
            Write-Success "Git already installed: $existingVersion"
            return
        }
        
        Write-Info "Installing Git via Chocolatey..."
        $result = Invoke-ChocolateyCommand -Command "install" -Arguments "git -y"
        Write-Info "Chocolatey output: $result"
        
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Git installation failed with exit code: $LASTEXITCODE"
            Write-Warning "Repository cloning may not work correctly"
        } else {
            Write-Success "Git installed"
        }
    } catch {
        Write-Warning "Failed to install Git: $_"
        Write-Warning "Repository cloning may not work correctly"
    }
}

function Enable-WSL {
    Write-Info "Enabling Windows Subsystem for Linux (WSL)..."
    
    try {
        # Check if running as administrator
        if (-not (Test-Administrator)) {
            Write-Warning "WSL requires administrator privileges. Skipping WSL setup."
            Write-Warning "To enable WSL manually:"
            Write-Warning "1. Right-click PowerShell and select 'Run as Administrator'"
            Write-Warning "2. Run this installer again with administrator privileges"
            Write-Warning "3. Or manually enable WSL with these commands:"
            Write-Warning "   dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart"
            Write-Warning "   dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart"
            Write-Warning "   wsl --install"
            return
        }
        
        # Check if WSL is already enabled
        $wslFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -ErrorAction SilentlyContinue
        if ($wslFeature -and $wslFeature.State -eq "Enabled") {
            Write-Success "WSL is already enabled"
            return
        }
        
        # Enable WSL feature
        Write-Info "Enabling WSL feature..."
        $result1 = dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart 2>&1
        Write-Info "DISM output: $result1"
        
        Write-Info "Enabling Virtual Machine Platform feature..."
        $result2 = dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart 2>&1
        Write-Info "DISM output: $result2"
        
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "WSL feature enablement failed with exit code: $LASTEXITCODE"
            Write-Warning "WSL may not work correctly"
        } else {
            Write-Success "WSL features enabled"
            Write-Warning "Please restart your computer and run: wsl --install"
            Write-Warning "Then run this installer again to continue with WSL setup"
        }
        
    } catch {
        Write-Warning "Failed to enable WSL: $_"
        Write-Warning "WSL may not work correctly"
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
        # Clone or download the repository with authentication
        if (Get-Command git -ErrorAction SilentlyContinue) {
            Write-Info "Cloning repository with authentication..."
            
            # Configure git to use the token for authentication
            $repoUrl = "https://$GitHubToken@github.com/$GITHUB_REPO.git"
            
            # Clone the repository
            git clone $repoUrl .
            
            # Remove token from git config for security
            git config --unset credential.helper
            git config --unset credential.https://github.com.username
            
        } else {
            Write-Info "Downloading repository as ZIP with authentication..."
            
            # Use GitHub API to download the repository as ZIP
            $zipUrl = "https://api.github.com/repos/$GITHUB_REPO/zipball/main"
            $zipFile = Join-Path $InstallDir "psycopilot.zip"
            
            $headers = @{
                "Authorization" = "token $GitHubToken"
                "Accept" = "application/vnd.github.v3+json"
            }
            
            Write-Info "Downloading from GitHub API..."
            Invoke-WebRequest -Uri $zipUrl -Headers $headers -OutFile $zipFile
            
            Write-Info "Extracting repository..."
            Expand-Archive -Path $zipFile -DestinationPath $InstallDir -Force
            
            # Move contents from subdirectory (GitHub API creates a folder with commit hash)
            $extractedDirs = Get-ChildItem -Path $InstallDir -Directory | Where-Object { $_.Name -like "gsnegovskiy-psycopilot-*" }
            if ($extractedDirs.Count -gt 0) {
                $extractedDir = $extractedDirs[0].FullName
                Get-ChildItem -Path $extractedDir | Move-Item -Destination $InstallDir
                Remove-Item -Path $extractedDir -Recurse -Force
            }
            Remove-Item -Path $zipFile -Force
        }
        
        Write-Success "PsycoPilot application downloaded"
        
    } catch {
        Write-Error "Failed to download PsycoPilot: $_"
        Write-Error "Please check your GitHub token and repository access permissions."
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
    Write-ColorOutput "💡 Installation Command:" "Cyan"
    Write-ColorOutput "• `$env:GITHUB_TOKEN='your_token'; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/gsnegovskiy/psycopilot-public/master/psycopilot-installer-windows.ps1' | Invoke-Expression" "White"
    Write-Host ""
    Write-ColorOutput "🔐 Security Note:" "Yellow"
    Write-ColorOutput "• Your GitHub token was used for authentication and has been cleared from git config" "White"
    Write-ColorOutput "• Keep your GitHub token secure and do not share it" "White"
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
        
        Write-Host ""
        Write-ColorOutput "===========================================" "Cyan"
        Write-ColorOutput "RUNNING SYSTEM CHECKS" "Cyan"
        Write-ColorOutput "===========================================" "Cyan"
        Write-Host ""
        
        Write-Info "Checking system requirements..."
        Test-SystemRequirements
        
        Write-Info "Validating GitHub token..."
        Test-GitHubToken
        
        Write-Info "Checking installation directory..."
        Test-InstallationDirectory
        
        Write-Host ""
        Write-ColorOutput "System checks completed!" "Green"
        Write-Host ""
        
        # Install system dependencies
        Write-Host ""
        Write-ColorOutput "===========================================" "Cyan"
        Write-ColorOutput "INSTALLING SYSTEM DEPENDENCIES" "Cyan"
        Write-ColorOutput "===========================================" "Cyan"
        Write-Host ""
        
        Write-Info "Starting system dependency installation..."
        Install-Chocolatey
        
        Write-Host ""
        Write-Info "Installing Python..."
        Install-Python
        Refresh-EnvironmentPath
        
        Write-Host ""
        Write-Info "Installing Visual C++ Redistributables..."
        Install-VisualCppRedistributables
        
        Write-Host ""
        Write-Info "Installing Git..."
        Install-Git
        Refresh-EnvironmentPath
        
        Write-Host ""
        Write-ColorOutput "System dependencies installation completed!" "Green"
        Write-Host ""
        
        # Enable WSL by default (unless explicitly disabled or only testing audio)
        if ($EnableWSL -or (-not $TestAudioOnly)) {
            Write-Info "Enabling Windows Subsystem for Linux (WSL)..."
            Enable-WSL
            Write-Info "WSL setup completed. Continuing with installation..."
        }
        
        if ($InstallVirtualAudio) {
            Install-VirtualAudio
        }
        
        # Setup application
        Write-Host ""
        Write-ColorOutput "===========================================" "Cyan"
        Write-ColorOutput "SETTING UP PSYCOPILOT APPLICATION" "Cyan"
        Write-ColorOutput "===========================================" "Cyan"
        Write-Host ""
        
        Write-Info "Setting up Python environment..."
        Setup-PythonEnvironment
        
        Write-Info "Downloading PsycoPilot application..."
        Download-PsycoPilot
        
        Write-Info "Installing Python dependencies..."
        Install-PythonDependencies
        
        Write-Info "Creating start scripts..."
        Create-StartScripts
        
        Write-Host ""
        Write-ColorOutput "Application setup completed!" "Green"
        Write-Host ""
        
        # Audio setup and testing
        Write-Host ""
        Write-ColorOutput "🔊 Audio Setup and Testing" "Cyan"
        Test-AudioDevices
        Enable-StereoMix
        Test-AudioSetup
        
        Show-CompletionMessage
        
        Write-Host ""
        Write-ColorOutput "Installation completed successfully!" "Green"
        Write-ColorOutput "Keeping terminal open for 30 seconds so you can read the output..." "Yellow"
        Write-ColorOutput "Press any key to exit immediately, or wait 30 seconds..." "Yellow"
        
        # Try to read a key with timeout
        try {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } catch {
            Write-ColorOutput "Press Enter to continue..." "Green"
            Read-Host
        }
        
        # If no key was pressed, wait 30 seconds
        Write-ColorOutput "Waiting 30 seconds before closing..." "Yellow"
        Start-Sleep -Seconds 30
        
    } catch {
        Write-Error "Installation failed: $_"
        Write-Error "Error details: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        Write-Host ""
        Write-ColorOutput "Installation failed!" "Red"
        Write-ColorOutput "Keeping terminal open for 30 seconds so you can read the error details..." "Yellow"
        Write-ColorOutput "Press any key to exit immediately, or wait 30 seconds..." "Yellow"
        
        # Try to read a key with timeout
        try {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } catch {
            Write-ColorOutput "Press Enter to continue..." "Yellow"
            Read-Host
        }
        
        # If no key was pressed, wait 30 seconds
        Write-ColorOutput "Waiting 30 seconds before closing..." "Yellow"
        Start-Sleep -Seconds 30
        exit 1
    }
}

# Add debugging information
Write-Host ""
Write-ColorOutput "===========================================" "Green"
Write-ColorOutput "STARTING PSYCOPILOT INSTALLER" "Green"
Write-ColorOutput "===========================================" "Green"
Write-Host ""
Write-Info "PowerShell Version: $($PSVersionTable.PSVersion)"
Write-Info "PowerShell Edition: $($PSVersionTable.PSEdition)"
Write-Info "OS Version: $([System.Environment]::OSVersion.Version)"
Write-Info "Current User: $([System.Environment]::UserName)"
Write-Info "Current Directory: $(Get-Location)"
Write-Info "Script Path: $($MyInvocation.MyCommand.Path)"
Write-Host ""
Write-ColorOutput "Script is starting... Press any key to continue or wait 10 seconds..." "Yellow"

# Add a timeout pause to ensure the script is running
try {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} catch {
    Write-ColorOutput "No key pressed, waiting 10 seconds before starting..." "Yellow"
    Start-Sleep -Seconds 10
}

Write-Host ""

# Wrap everything in a try-catch to prevent terminal from closing
try {
    # Run main function
    Main
} catch {
    Write-Host ""
    Write-ColorOutput "===========================================" "Red"
    Write-ColorOutput "INSTALLATION FAILED" "Red"
    Write-ColorOutput "===========================================" "Red"
    Write-Host ""
    Write-Error "Error: $_"
    Write-Error "Error details: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
    Write-Host ""
    Write-ColorOutput "This error occurred outside the main installation process." "Yellow"
    Write-ColorOutput "Please check the error details above and try again." "Yellow"
    Write-Host ""
    Write-ColorOutput "Critical error occurred!" "Red"
    Write-ColorOutput "Keeping terminal open for 30 seconds so you can read the error details..." "Yellow"
    Write-ColorOutput "Press any key to exit immediately, or wait 30 seconds..." "Yellow"
    
    # Try to read a key with timeout
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Write-ColorOutput "Press Enter to continue..." "Red"
        Read-Host
    }
    
    # If no key was pressed, wait 30 seconds
    Write-ColorOutput "Waiting 30 seconds before closing..." "Yellow"
    Start-Sleep -Seconds 30
    exit 1
}
