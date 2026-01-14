# --- PS2EXE Module Check ---
$modName = "PS2EXE"
if (!(Get-Module -ListAvailable $modName)) {
    Write-Host "$modName module not found, installing..." -ForegroundColor Yellow
    Install-Module -Name $modName -Scope CurrentUser -Force
}
Import-Module $modName

# --- Compilation Settings ---
$Params = @{
    InputFile   = ".\HomeBackup.ps1"
    OutputFile  = ".\HomeBackup.exe"
    IconFile    = ".\icon.ico"
    Title       = "Home Backup & Restore"
    Description = "Modern backup solution"
    Company     = "Osman Onur Koç"
    Product     = "Windows Backup"
    Copyright   = "www.osmanonurkoc.com"
    Version     = "11.4.0.0"
    NoConsole   = $true
    STA         = $true  # Critical for WPF UI (Single Threaded Apartment)
}

# --- Icon Check ---
# If icon.ico does not exist in the folder, remove the parameter to use default icon
if (!(Test-Path $Params.IconFile)) {
    Write-Warning "WARNING: icon.ico not found. Default icon will be used."
    $Params.Remove('IconFile')
}

# --- Start Compilation ---
Write-Host "Starting compilation process..." -ForegroundColor Cyan
try {
    Invoke-PS2EXE @Params
    Write-Host "`nSUCCESS: HomeBackup.exe created successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Error occurred during compilation: $_"
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
