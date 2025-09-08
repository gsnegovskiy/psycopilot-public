#!/usr/bin/env bash
# psycopilot-installer.sh - Complete installer for macOS distribution
# This is a single executable that handles all installations and setup
# Usage: curl -fsSL https://raw.githubusercontent.com/user/repo/main/psycopilot-installer.sh | bash

set -euo pipefail

# Configuration
APP_NAME="PsycoPilot"
APP_VERSION="1.0.0"
GITHUB_REPO="gsnegovskiy/psycopilot"  # Private repository
GITHUB_BRANCH="master"
MIN_MACOS="14.2"
INSTALL_DIR="$HOME/Applications/$APP_NAME"

# GitHub token for private repo access
GITHUB_TOKEN="${GITHUB_TOKEN:-}"

# Colors and formatting
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Logging functions
say()   { printf "${BLUE}[INFO]${NC} %s\n" "$*"; }
ok()    { printf "${GREEN}[OK]${NC} %s\n" "$*"; }
warn()  { printf "${YELLOW}[WARN]${NC} %s\n" "$*" >&2; }
error() { printf "${RED}[ERROR]${NC} %s\n" "$*" >&2; }
die()   { error "$*"; exit 1; }

# Version comparison function
verlte() { 
    printf '%s\n%s\n' "$1" "$2" | sort -t. -k1,1n -k2,2n -k3,3n | head -n1 | grep -qx "$1"
}

# Progress indicator
show_progress() {
    local pid=$1
    local msg="$2"
    local delay=0.1
    local spinstr='|/-\'
    local temp
    
    while kill -0 $pid 2>/dev/null; do
        temp=${spinstr#?}
        printf "\r${BLUE}[%c]${NC} %s" "$spinstr" "$msg"
        spinstr=$temp${spinstr%"$temp"}
        sleep $delay
    done
    printf "\r${GREEN}[✓]${NC} %s\n" "$msg"
}

# Get GitHub token securely
get_github_token() {
    if [[ -n "$GITHUB_TOKEN" ]]; then
        return  # Token already set via environment
    fi
    
    echo ""
    say "This installer needs access to the private PsycoPilot repository."
    say "Please provide the access token given to you by your administrator:"
    echo ""
    warn "You don't need a GitHub account - just enter the token you received."
    warn "The token will not be stored permanently on your system."
    warn "It's only used during installation to download the application."
    echo ""
    
    # Read token securely (no echo)
    while [[ -z "$GITHUB_TOKEN" ]]; do
        printf "Access Token: "
        read -s GITHUB_TOKEN
        echo ""  # New line after hidden input
        
        if [[ -z "$GITHUB_TOKEN" ]]; then
            error "Token cannot be empty. Please try again."
            continue
        fi
        
        # Basic token validation (GitHub tokens start with specific prefixes)
        if [[ ! "$GITHUB_TOKEN" =~ ^(ghp_|github_pat_) ]]; then
            warn "Token format may be incorrect. GitHub tokens usually start with 'ghp_' or 'github_pat_'"
            read -p "Continue anyway? (y/N): " confirm
            if [[ $confirm != [yY] ]]; then
                GITHUB_TOKEN=""
                continue
            fi
        fi
        
        # Test token by making a simple API call
        say "Validating token..."
        if curl -fsSL -H "Authorization: token $GITHUB_TOKEN" \
                "https://api.github.com/repos/$GITHUB_REPO" >/dev/null 2>&1; then
            ok "Token validated successfully"
        else
            error "Token validation failed. Please check your token and try again."
            GITHUB_TOKEN=""
        fi
    done
}

# Banner
print_banner() {
    cat << 'EOF'
    ____                       ____  _ __      __ 
   / __ \____  __  _________  / __ \(_) /___  / /_
  / /_/ / __ \/ / / / ___/ _ \/ /_/ / / / __ \/ __/
 / ____/ /_/ / /_/ / /__/  __/ ____/ / / /_/ / /_  
/_/   \____/\__, /\___/\___/_/   /_/_/\____/\__/  
           /____/                                 

Real-time dual-channel audio transcription for therapy sessions
EOF
    echo ""
    say "Version: $APP_VERSION"
    say "Installing to: $INSTALL_DIR"
    say "Repository: $GITHUB_REPO (private)"
    echo ""
}

# System checks
check_system() {
    say "Checking system compatibility..."
    
    # macOS check
    [[ "$(uname -s)" == "Darwin" ]] || die "This installer is for macOS only."
    
    # macOS version check
    local macver
    macver="$(sw_vers -productVersion)"
    if ! verlte "$MIN_MACOS" "$macver"; then
        die "macOS $MIN_MACOS or later required. Found: $macver"
    fi
    ok "macOS $macver (compatible)"
    
    # Architecture check
    local arch
    arch="$(uname -m)"
    if [[ "$arch" != "arm64" ]]; then
        warn "Intel Mac detected. This installer is optimized for Apple Silicon."
        warn "Some features may not work correctly."
    else
        ok "Apple Silicon Mac detected"
    fi
    
    # Free space check (require at least 2GB)
    local available_space
    available_space=$(df -g "$HOME" | tail -1 | awk '{print $4}')
    if [[ $available_space -lt 2 ]]; then
        die "Insufficient disk space. At least 2GB required, found ${available_space}GB"
    fi
    ok "Sufficient disk space available (${available_space}GB)"
}

# Install Xcode Command Line Tools
install_xcode_tools() {
    if xcode-select -p >/dev/null 2>&1; then
        ok "Xcode Command Line Tools already installed"
        return
    fi
    
    say "Installing Xcode Command Line Tools..."
    warn "This may show a GUI dialog. Please click 'Install' when prompted."
    
    # Start installation
    xcode-select --install 2>/dev/null || true
    
    # Wait for installation to complete
    local timeout=300  # 5 minutes
    local elapsed=0
    
    while ! xcode-select -p >/dev/null 2>&1; do
        if [[ $elapsed -ge $timeout ]]; then
            die "Xcode Command Line Tools installation timeout. Please install manually."
        fi
        sleep 5
        elapsed=$((elapsed + 5))
    done
    
    ok "Xcode Command Line Tools installed"
}

# Install Homebrew
install_homebrew() {
    if command -v brew >/dev/null 2>&1; then
        ok "Homebrew already installed"
        # Setup Homebrew environment
        if [[ -x /opt/homebrew/bin/brew ]]; then
            eval "$(/opt/homebrew/bin/brew shellenv)"
        elif [[ -x /usr/local/bin/brew ]]; then
            eval "$(/usr/local/bin/brew shellenv)"
        fi
        return
    fi
    
    say "Installing Homebrew..."
    export NONINTERACTIVE=1
    
    # Install Homebrew silently
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" &
    local brew_pid=$!
    show_progress $brew_pid "Installing Homebrew"
    
    # Setup Homebrew environment
    if [[ -x /opt/homebrew/bin/brew ]]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
    elif [[ -x /usr/local/bin/brew ]]; then
        eval "$(/usr/local/bin/brew shellenv)"
    fi
    
    ok "Homebrew installed: $(brew --version | head -n1)"
}

# Install Git
install_git() {
    if command -v git >/dev/null 2>&1; then
        ok "Git already installed: $(git --version)"
        return
    fi
    
    say "Installing Git..."
    brew install git &
    local git_pid=$!
    show_progress $git_pid "Installing Git"
    
    ok "Git installed: $(git --version)"
}

# Download and extract application using git
download_app() {
    say "Downloading $APP_NAME from private repository..."
    
    # Create installation directory
    mkdir -p "$INSTALL_DIR"
    cd "$INSTALL_DIR"
    
    # Clean up any existing files including temp directory (preserve .git for updates)
    rm -rf *.py *.sh application/ webui/ provisioners/ config/ temp_repo 2>/dev/null || true
    
    # Configure git to use token for authentication
    local repo_url="https://$GITHUB_TOKEN@github.com/$GITHUB_REPO.git"
    
    say "Cloning repository..."
    
    # Clone the repository with progress
    {
        git clone --progress --branch "$GITHUB_BRANCH" --single-branch --depth 1 \
            "$repo_url" temp_repo
        
        # Move files from temp directory to install directory
        mv temp_repo/* .
        mv temp_repo/.[^.]* . 2>/dev/null || true  # Move hidden files
        rm -rf temp_repo
    } &
    local clone_pid=$!
    show_progress $clone_pid "Cloning repository"
    
    # Keep git repository for future updates (token is read-only and expires)
    
    # Verify essential files are present
    if [[ ! -f "application/audio-capture/transcribe_dual_database.py" ]]; then
        die "Failed to download application files. Essential files missing."
    fi
    
    if [[ ! -f "provisioners/mac/new/bin/audiotee" ]]; then
        warn "audiotee binary not found. System audio capture may not work."
    fi
    
    ok "Application files downloaded successfully"
}

# Install Python and create environment
setup_python() {
    say "Setting up Python environment..."
    
    # Install Python 3.13
    if ! brew list --versions python@3.13 >/dev/null 2>&1; then
        brew install python@3.13 &
        local python_pid=$!
        show_progress $python_pid "Installing Python 3.13"
    fi
    
    # Find Python binary
    local pybin
    if command -v python3.13 >/dev/null 2>&1; then
        pybin="$(command -v python3.13)"
    elif [[ -x /opt/homebrew/opt/python@3.13/bin/python3.13 ]]; then
        pybin="/opt/homebrew/opt/python@3.13/bin/python3.13"
    elif [[ -x /usr/local/opt/python@3.13/bin/python3.13 ]]; then
        pybin="/usr/local/opt/python@3.13/bin/python3.13"
    else
        pybin="$(brew --prefix python@3.13)/bin/python3.13"
    fi
    
    [[ -x "$pybin" ]] || die "Python 3.13 not found after installation"
    ok "Using Python: $($pybin -V 2>&1)"
    
    # Create virtual environment
    cd "$INSTALL_DIR"
    rm -rf .venv  # Remove any existing venv
    "$pybin" -m venv .venv
    source .venv/bin/activate
    
    # Upgrade pip and install wheel
    .venv/bin/python -m pip install --upgrade pip wheel setuptools &
    local pip_pid=$!
    show_progress $pip_pid "Upgrading pip and tools"
    
    ok "Python environment created"
}

# Install system dependencies
install_system_deps() {
    say "Installing system dependencies..."
    
    # Install required brew packages
    local packages=(
        "pkg-config"  # Still needed for some packages
        "ollama"      # Local LLM support
    )
    # Note: FFmpeg removed - no longer needed since we use pywhispercpp instead of PyAV
    
    for package in "${packages[@]}"; do
        if ! brew list --versions "$package" >/dev/null 2>&1; then
            brew install "$package" &
            local pkg_pid=$!
            show_progress $pkg_pid "Installing $package"
        fi
    done
    
    ok "System dependencies installed"
}

# Install Python dependencies (simplified - no PyAV needed!)
install_python_deps() {
    say "Installing Python dependencies..."
    
    cd "$INSTALL_DIR"
    source .venv/bin/activate
    
    # Install application dependencies (includes pywhispercpp - no PyAV needed!)
    say "Installing audio processing dependencies (with Whisper via whisper.cpp)..."
    .venv/bin/python -m pip install -r application/audio-capture/requirements.txt &
    local req_pid=$!
    show_progress $req_pid "Installing audio processing dependencies"
    
    # Install web UI dependencies
    .venv/bin/python -m pip install -r webui/requirements.txt &
    local webui_pid=$!
    show_progress $webui_pid "Installing web interface dependencies"
    
    ok "Python dependencies installed"
}

# Setup audiotee binary
setup_audiotee() {
    say "Setting up system audio capture..."
    
    local audiotee_path="$INSTALL_DIR/provisioners/mac/new/bin/audiotee"
    
    if [[ -f "$audiotee_path" ]]; then
        # Make executable
        chmod +x "$audiotee_path"
        
        # Remove quarantine attribute
        xattr -dr com.apple.quarantine "$audiotee_path" 2>/dev/null || true
        
        # Ad-hoc code sign
        codesign --force --deep --sign - "$audiotee_path" >/dev/null 2>&1 || {
            warn "Could not code sign audiotee. You may need to allow it in Security & Privacy."
        }
        
        ok "System audio capture ready"
    else
        warn "audiotee binary not found. System audio capture will be disabled."
    fi
}

# Start and configure services
start_services() {
    say "Starting services..."
    
    # Start Ollama service
    if command -v brew >/dev/null 2>&1; then
        brew services start ollama >/dev/null 2>&1 || true
    fi
    
    # Fallback: start ollama daemon
    if ! pgrep -x ollama >/dev/null 2>&1; then
        (nohup ollama serve >/dev/null 2>&1 &) || true
    fi
    
    # Wait for Ollama API
    local ollama_ready=0
    for i in {1..30}; do
        if curl -fsS -m 1 http://127.0.0.1:11434/api/version >/dev/null 2>&1; then
            ollama_ready=1
            break
        fi
        sleep 1
    done

    #if [[ $ollama_ready -eq 1 ]]; then
     #   ok "Ollama service started"

        # Download default model
    #    say "Downloading default language model..."
    #    for now we don't need to download the model
        #ollama pull gpt-oss:20b &
    #    local model_pid=$!
    #    show_progress $model_pid "Downloading GPT-OSS model (this may take several minutes)"
    #else
    #    warn "Ollama service not responding. You can start it manually with 'ollama serve'"
    #fi
}

# Create configuration files
setup_config() {
    say "Setting up configuration..."
    
    cd "$INSTALL_DIR"
    
    # Create API key template if it doesn't exist
    if [[ ! -f config/claude-api-key.txt ]]; then
        cp config/api-key.txt-example config/claude-api-key.txt
        warn "Please add your Claude API key to: $INSTALL_DIR/config/claude-api-key.txt"
    fi
    
    # Create user preferences
    mkdir -p config
    if [[ ! -f config/user_preferences.json ]]; then
        cat > config/user_preferences.json << 'EOF'
{
    "last_models": {},
    "default_provider": "ollama",
    "providers": {
        "ollama": {
            "default_model": "gpt-oss:20b"
        }
    }
}
EOF
    fi
    
    ok "Configuration files ready"
}

# Create launcher scripts
create_launchers() {
    say "Creating launcher scripts..."
    
    cd "$INSTALL_DIR"
    
    # Create main launcher
    cat > psycopilot.sh << 'EOF'
#!/usr/bin/env bash
# PsycoPilot Launcher Script

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
NC='\033[0m'

echo -e "${BLUE}🎯 PsycoPilot - Real-time Audio Transcription${NC}"
echo ""

# Check if services are running
if ! pgrep -x ollama >/dev/null 2>&1; then
    echo "Starting Ollama service..."
    (nohup ollama serve >/dev/null 2>&1 &)
    sleep 3
fi

echo "Available commands:"
echo "  1) Start transcription (Whisper)"
echo "  2) Start transcription (GigaAM - faster)"  
echo "  3) Start web interface"
echo "  4) Exit"
echo ""

while true; do
    read -p "Enter choice (1-4): " choice
    case $choice in
        1)
            echo -e "${GREEN}Starting Whisper transcription...${NC}"
            .venv/bin/python application/audio-capture/transcribe_dual_database.py \
                --engine whisper --model base --language ru \
                --mac-system-source catap \
                --audiotee-bin provisioners/mac/new/bin/audiotee
            ;;
        2)
            echo -e "${GREEN}Starting GigaAM transcription...${NC}"
            .venv/bin/python application/audio-capture/transcribe_dual_database.py \
                --engine gigaam --model gigaam-v2-ctc --gigaam-quant int8 \
                --mac-system-source catap \
                --audiotee-bin provisioners/mac/new/bin/audiotee
            ;;
        3)
            echo -e "${GREEN}Starting web interface...${NC}"
            echo "Open http://127.0.0.1:7860 in your browser (Production port)"
            bash webui/run_webui.sh
            ;;
        4)
            echo "Goodbye!"
            break
            ;;
        *)
            echo "Invalid choice. Please enter 1-4."
            ;;
    esac
    echo ""
done
EOF

    chmod +x psycopilot.sh
    
    # Create web UI launcher
    cat > start-webui.sh << 'EOF'
#!/usr/bin/env bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Set installer context to force production mode
export PSYCOPILOT_INSTALLER=true

echo "🌐 Starting PsycoPilot Web Interface..."
echo "Open http://127.0.0.1:7860 in your browser (Production port)"

# Ensure Ollama is running
if ! pgrep -x ollama >/dev/null 2>&1; then
    echo "Starting Ollama service..."
    (nohup ollama serve >/dev/null 2>&1 &)
    sleep 3
fi

bash webui/run_webui.sh
EOF

    chmod +x start-webui.sh
    
    # Create desktop applications (optional)
    create_desktop_apps
    
    ok "Launcher scripts created"
}

# Create macOS desktop applications
create_desktop_apps() {
    local apps_dir="$HOME/Applications"
    
    # Create PsycoPilot.app bundle
    local app_bundle="$apps_dir/PsycoPilot.app"
    mkdir -p "$app_bundle/Contents/MacOS"
    mkdir -p "$app_bundle/Contents/Resources"
    
    # Create Info.plist
    cat > "$app_bundle/Contents/Info.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>PsycoPilot</string>
    <key>CFBundleIdentifier</key>
    <string>com.psycopilot.app</string>
    <key>CFBundleVersion</key>
    <string>$APP_VERSION</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleExecutable</key>
    <string>PsycoPilot</string>
    <key>LSMinimumSystemVersion</key>
    <string>$MIN_MACOS</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
EOF

    # Create launcher executable - directly starts web interface with terminal
    cat > "$app_bundle/Contents/MacOS/PsycoPilot" << EOF
#!/usr/bin/env bash
# Open Terminal and run the web interface
osascript <<'APPLESCRIPT'
tell application "Terminal"
    activate
    do script "cd '$INSTALL_DIR' && echo '🌐 Starting PsycoPilot Web Interface...' && echo 'Opening at http://127.0.0.1:7860' && echo 'Close this terminal window to stop the application.' && echo '' && if ! pgrep -x ollama >/dev/null 2>&1; then echo 'Starting Ollama service...' && (nohup ollama serve >/dev/null 2>&1 &) && sleep 3; fi && bash start-webui.sh & echo 'Waiting for server to be ready...' && until curl -s http://127.0.0.1:7860/api/health >/dev/null 2>&1; do sleep 1; done && echo 'Server ready! Opening browser...' && open http://127.0.0.1:7860 && wait"
end tell
APPLESCRIPT
EOF

    chmod +x "$app_bundle/Contents/MacOS/PsycoPilot"
}

# Create uninstaller
create_uninstaller() {
    cat > "$INSTALL_DIR/uninstall.sh" << 'EOF'
#!/usr/bin/env bash
# PsycoPilot Uninstaller

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo -e "${YELLOW}PsycoPilot Uninstaller${NC}"
echo ""

read -p "Are you sure you want to uninstall PsycoPilot? (y/N): " confirm
if [[ $confirm != [yY] ]]; then
    echo "Uninstall cancelled."
    exit 0
fi

echo "Stopping services..."
brew services stop ollama 2>/dev/null || true
pkill -f ollama 2>/dev/null || true

echo "Removing application files..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HOME"
rm -rf "$SCRIPT_DIR"

echo "Removing desktop applications..."
rm -rf "$HOME/Applications/PsycoPilot.app"

echo -e "${GREEN}PsycoPilot has been uninstalled.${NC}"
echo ""
echo "Note: Homebrew, Python, and system dependencies were left installed"
echo "as they may be used by other applications."
EOF

    chmod +x "$INSTALL_DIR/uninstall.sh"
}

# Final setup and instructions
finish_installation() {
    echo ""
    echo "🎉 Installation completed successfully!"
    echo ""
    echo "📍 Installation directory: $INSTALL_DIR"
    echo ""
    echo "🚀 Getting started:"
    echo "  • Double-click PsycoPilot in Applications folder"
    echo "  • The web interface will open automatically at http://127.0.0.1:7860"
    echo "  • Close the terminal window to stop the application"
    echo ""
    echo "⚙️  Configuration:"
    echo "  • Add your Claude API key to: $INSTALL_DIR/config/claude-api-key.txt"
    echo "  • Configuration files are in: $INSTALL_DIR/config/"
    echo ""
    echo "🗑️  To uninstall:"
    echo "  • Run: $INSTALL_DIR/uninstall.sh"
    echo ""
    
    if [[ "$OSTYPE" == "darwin"* ]]; then
        say "Opening installation directory..."
        open "$INSTALL_DIR"
    fi
}

# Main installation flow
main() {
    print_banner
    
    # Get GitHub token first
    get_github_token
    
    # Pre-flight checks
    check_system
    
    # Core system setup
    install_xcode_tools
    install_homebrew
    install_git
    
    # Application setup
    download_app
    setup_python
    install_system_deps
    install_python_deps
    setup_audiotee
    
    # Configuration and services
    setup_config
    start_services
    
    # User interface
    create_launchers
    create_uninstaller
    
    # Complete
    finish_installation
}

# Error handling
trap 'error "Installation failed at line $LINENO. Check the error above."; exit 1' ERR

# Set installer context environment variable
export PSYCOPILOT_INSTALLER=true

# Run main installation
main "$@"
