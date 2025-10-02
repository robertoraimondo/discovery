# Auto-Administrator Launcher for System Information GUI
# This script ensures the GUI always runs with Administrator privileges

param(
    [string]$GUIVersion = "Stable"
)

Write-Host "`n=== SYSTEM INFORMATION GUI - AUTO-ADMIN LAUNCHER ===" -ForegroundColor Green

# Function to check if running as Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to request Administrator privileges
function Request-Administrator {
    param([string]$ScriptToRun)
    
    Write-Host "`nREQUESTING ADMINISTRATOR PRIVILEGES..." -ForegroundColor Yellow
    Write-Host "A User Account Control (UAC) prompt will appear." -ForegroundColor Cyan
    Write-Host "Please click 'Yes' to grant Administrator access." -ForegroundColor White
    
    try {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptToRun`""
        Start-Process PowerShell -Verb RunAs -ArgumentList $arguments -WorkingDirectory $PSScriptRoot
        Write-Host "`nAdministrator session launching..." -ForegroundColor Green
        exit
    } catch {
        Write-Host "`nFailed to request Administrator privileges:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "`nPlease manually run PowerShell as Administrator." -ForegroundColor Yellow
        return $false
    }
}

# Check current privileges
if (-not (Test-Administrator)) {
    Write-Host "`n[USER] Current session running as standard user" -ForegroundColor Yellow
    Write-Host "Administrator privileges required for:" -ForegroundColor White
    Write-Host "  • System-level software installation" -ForegroundColor Gray
    Write-Host "  • WinRM and service configuration" -ForegroundColor Gray
    Write-Host "  • Network share and registry access" -ForegroundColor Gray
    Write-Host "  • Package manager operations (Chocolatey, winget)" -ForegroundColor Gray
    
    # Determine which GUI to launch
    $guiScript = switch ($GUIVersion.ToLower()) {
        "enhanced" { "SystemInfoGUI-Enhanced.ps1" }
        "debug" { "SystemInfoGUI-Debug.ps1" }
        default { "SystemInfoGUI-Stable.ps1" }
    }
    
    $guiPath = Join-Path $PSScriptRoot $guiScript
    
    if (Test-Path $guiPath) {
        Request-Administrator -ScriptToRun $guiPath
    } else {
        Write-Host "`nGUI script not found: $guiPath" -ForegroundColor Red
        Write-Host "Available versions: Stable, Enhanced, Debug" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n[ADMIN] Already running with Administrator privileges!" -ForegroundColor Green
    Write-Host "You can launch any GUI version directly:" -ForegroundColor White
    Write-Host "  .\SystemInfoGUI-Stable.ps1   (Recommended)" -ForegroundColor Cyan
    Write-Host "  .\SystemInfoGUI-Enhanced.ps1 (Full features)" -ForegroundColor Cyan
    Write-Host "  .\SystemInfoGUI-Debug.ps1    (Basic version)" -ForegroundColor Cyan
}

Write-Host "`nAUTO-ADMIN LAUNCHER USAGE:" -ForegroundColor Yellow
Write-Host "  .\AlwaysAdmin.ps1              (Launch Stable version)" -ForegroundColor White
Write-Host "  .\AlwaysAdmin.ps1 -GUIVersion Enhanced   (Launch Enhanced)" -ForegroundColor White
Write-Host "  .\AlwaysAdmin.ps1 -GUIVersion Debug      (Launch Debug)" -ForegroundColor White