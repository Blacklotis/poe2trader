param(
  [string]$RepoUrl = "https://github.com/Blacklotis/poe2trader",
  [string]$InstallDir = "$HOME\\poe2trader"
)

function Write-Step {
  param([string]$Message)
  Write-Host "==> $Message"
}

function Has-Command {
  param([string]$Name)
  return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

Write-Step "Checking prerequisites"

if (-not (Has-Command git)) {
  if (Has-Command winget) {
    Write-Step "Installing Git (winget)"
    winget install --id Git.Git -e --source winget
  } else {
    throw "Git is not installed and winget is unavailable. Install Git manually."
  }
}

$pythonCmd = $null
if (Has-Command py) {
  $pythonCmd = "py -3"
} elseif (Has-Command python) {
  $pythonCmd = "python"
}

if (-not $pythonCmd) {
  if (Has-Command winget) {
    Write-Step "Installing Python 3.12 (winget)"
    winget install --id Python.Python.3.12 -e --source winget
    $pythonCmd = "py -3"
  } else {
    throw "Python is not installed and winget is unavailable. Install Python 3.12 manually."
  }
}

if (-not (Test-Path $InstallDir)) {
  Write-Step "Cloning repo to $InstallDir"
  git clone $RepoUrl $InstallDir
} else {
  Write-Step "Repo already exists at $InstallDir (skipping clone)"
}

Set-Location $InstallDir

Write-Step "Creating virtual environment"
& $pythonCmd -m venv .venv

Write-Step "Installing Python packages"
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install opencv-python mss numpy pytesseract

Write-Step "Checking Tesseract"
$tessPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
if (-not (Test-Path $tessPath)) {
  $tessPath = "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
}

if (-not (Test-Path $tessPath)) {
  if (Has-Command winget) {
    Write-Step "Installing Tesseract OCR (winget)"
    winget install --id UB-Mannheim.TesseractOCR -e --source winget
  } else {
    Write-Host "Tesseract not found. Install it manually and set TESSERACT_CMD."
  }
}

if (Test-Path $tessPath) {
  Write-Step "Setting TESSERACT_CMD user environment variable"
  [Environment]::SetEnvironmentVariable("TESSERACT_CMD", $tessPath, "User")
  Write-Host "TESSERACT_CMD set to $tessPath (restart terminal to take effect)."
}

Write-Step "Done"
Write-Host "Run: .\\.venv\\Scripts\\python.exe main.py"
