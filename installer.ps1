Set-ExecutionPolicy Bypass -Scope Process -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function CommandExists($cmd) {
    return (Get-Command $cmd -ErrorAction SilentlyContinue) -ne $null
}

function Write-Info($msg) {
    Write-Host $msg -ForegroundColor Cyan
}

function Write-Success($msg) {
    Write-Host $msg -ForegroundColor Green
}

function Write-WarningMsg($msg) {
    Write-Host $msg -ForegroundColor Yellow
}

function Write-ErrorMsg($msg) {
    Write-Host $msg -ForegroundColor Red
} 

# تثبيت Chocolatey إذا مش موجودة
if (-not (CommandExists "choco")) {
    Write-Info "Installing Chocolatey..."
    try {
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-Success "Chocolatey installed successfully."
    } catch {
        Write-ErrorMsg "Failed to install Chocolatey: $_"
        exit 1
    }
} else {
    Write-Success "Chocolatey is already installed."
}

function InstallIfMissing($name, $checkCommand) {
    if (-not (CommandExists $checkCommand)) {
        Write-WarningMsg "Installing $name..."
        choco install $name -y --no-progress | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Success "$name installed successfully."
        } else {
            Write-ErrorMsg "Failed to install $name."
        }
    } else {
        Write-Success "$name is already installed."
    }
}

# تثبيت كل الأدوات ماعدا Postman
$tools = @(
    @{Name="Visual Studio Code"; Cmd="code"; Package="vscode"},
    @{Name="Git"; Cmd="git"; Package="git"},
    @{Name="Node.js"; Cmd="node"; Package="nodejs"},
    @{Name="Python"; Cmd="python"; Package="python"},
    @{Name="OpenJDK"; Cmd="java"; Package="openjdk"},
    @{Name="7-Zip"; Cmd="7z"; Package="7zip"},
    @{Name="curl"; Cmd="curl"; Package="curl"},
    @{Name="OpenSSL"; Cmd="openssl"; Package="openssl.light"},
    @{Name="Docker Desktop"; Cmd="docker"; Package="docker-desktop"}
)

foreach ($tool in $tools) {
    InstallIfMissing $tool.Name $tool.Cmd
}

# تثبيت Postman منفصل عشان ميبقاش فيه تعليق
Write-Info "Installing Postman separately..."
try {
    choco install postman -y --no-progress | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Postman installed successfully."
    } else {
        Write-WarningMsg "Postman installation might require manual intervention or failed."
    }
} catch {
    Write-WarningMsg "Postman installation might require manual intervention or failed."
}

# تثبيت باكيجات بايثون لو بايثون متثبت
if (CommandExists "python") {
    Write-Info "Installing/upgrading Python packages..."
    try {
        python -m ensurepip --upgrade | Out-Null
        python -m pip install --upgrade pip | Out-Null
        python -m pip install --upgrade virtualenv pytest black flake8 | Out-Null
        Write-Success "Python packages installed/upgraded successfully."
    } catch {
        Write-ErrorMsg "Failed to install Python packages: $_"
    }
} else {
    Write-ErrorMsg "Python is not installed, skipping Python packages."
}
