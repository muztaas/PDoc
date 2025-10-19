# Use local Python embedded distribution
$rootDir = Split-Path $PSScriptRoot -Parent
$srcDir = Join-Path $rootDir "src"
$pdocDir = Join-Path $srcDir "PDoc"
$pythonDir = Join-Path $pdocDir "Python"
$toolsSrcDir = Join-Path $rootDir "tools\src"

# Create Python directory if it doesn't exist
New-Item -ItemType Directory -Force -Path $pythonDir | Out-Null

# Copy Python embedded distribution
Write-Host "Copying Python embedded distribution..."
Copy-Item -Path "$toolsSrcDir\PDoc\Python\*" -Destination $pythonDir -Recurse -Force

# Create python310._pth file to enable site-packages
$pthContent = @"
python310.zip
.
Lib/site-packages
"@
Set-Content -Path "$pythonDir\python310._pth" -Value $pthContent -Force

# Install required packages
Write-Host "Installing required packages..."
$env:PYTHONPATH = Join-Path $pythonDir "Lib\site-packages"
$pythonExe = Join-Path $pythonDir "python.exe"

# Install pip and required packages
& $pythonExe -m ensurepip --upgrade
& $pythonExe -m pip install --no-warn-script-location wheel setuptools
& $pythonExe -m pip install --no-warn-script-location pywin32==311 docx2pdf --no-cache-dir

Write-Host "Setting up pywin32..."
$sitePackagesDir = Join-Path $pythonDir "Lib\site-packages"
$win32Dir = Join-Path $sitePackagesDir "win32"
$win32comDir = Join-Path $sitePackagesDir "win32com"
$win32comext = Join-Path $sitePackagesDir "win32comext"
$pywin32_system32 = Join-Path $sitePackagesDir "pywin32_system32"

# Copy DLL and PYD files to Python root
Write-Host "Copying required files to Python root..."
if (Test-Path $pywin32_system32) {
    Copy-Item -Path "$pywin32_system32\*" -Destination $pythonDir -Force
}

# Copy all .pyd files
foreach ($dir in @($win32Dir, $win32comDir, $win32comext)) {
    if (Test-Path $dir) {
        Get-ChildItem -Path $dir -Filter "*.pyd" -Recurse | ForEach-Object {
            Copy-Item $_.FullName -Destination $pythonDir -Force
            Write-Host "Copied $($_.Name) to Python directory"
        }
    }
}

# Register the DLLs
Write-Host "Registering Python COM Server DLLs..."
$pywintypes_dll = Join-Path $pythonDir "pywintypes310.dll"
$pythoncom_dll = Join-Path $pythonDir "pythoncom310.dll"

if (Test-Path $pywintypes_dll) {
    regsvr32.exe /s $pywintypes_dll
    Write-Host "Registered pywintypes310.dll"
} else {
    Write-Error "pywintypes310.dll not found!"
}

if (Test-Path $pythoncom_dll) {
    regsvr32.exe /s $pythoncom_dll
    Write-Host "Registered pythoncom310.dll"
} else {
    Write-Error "pythoncom310.dll not found!"
}

# Create necessary .pyd files from .dll files
Write-Host "Creating PYD files from DLL files..."
Copy-Item -Path (Join-Path $pythonDir "pythoncom310.dll") -Destination (Join-Path $pythonDir "pythoncom.pyd") -Force
Copy-Item -Path (Join-Path $pythonDir "pywintypes310.dll") -Destination (Join-Path $pythonDir "pywintypes.pyd") -Force

# Ensure the site-packages structure is correct
Write-Host "Setting up site-packages structure..."
if (Test-Path $win32Dir) {
    Copy-Item -Path $win32Dir -Destination $sitePackagesDir -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path $win32comDir) {
    Copy-Item -Path $win32comDir -Destination $sitePackagesDir -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path $win32comext) {
    Copy-Item -Path $win32comext -Destination $sitePackagesDir -Recurse -Force -ErrorAction SilentlyContinue
}

# Add win32/lib to Python path in python310._pth
$pthContent = @"
python310.zip
.
Lib/site-packages
Lib/site-packages/win32/lib
"@
Set-Content -Path "$pythonDir\python310._pth" -Value $pthContent -Force

# Run post-install script if it exists
$postinstall = Join-Path $pythonDir "Scripts\pywin32_postinstall.py"
if (Test-Path $postinstall) {
    Write-Host "Running pywin32 post-install script..."
    & $pythonExe $postinstall -install
}

Write-Host "Setup complete! Testing Python COM imports..."
$testImports = @'
import sys
print("Python executable:", sys.executable)
print("Testing imports...")
import pythoncom
print("pythoncom imported successfully")
import win32com.client
print("win32com.client imported successfully")
try:
    word = win32com.client.Dispatch('Word.Application')
    print("Word.Application created successfully")
    word.Quit()
except Exception as e:
    print("Error creating Word.Application:", str(e))
print("Python path:", sys.path)
import win32com
print("win32com imported successfully")
import pythoncom
print("pythoncom imported successfully")
'@

try {
    & $pythonExe -c $testImports
    Write-Host "All imports successful!"
} catch {
    Write-Error "Error testing imports: $_"
    exit 1
}