# WinRM Setup Script - Run as Administrator
# This script properly configures WinRM for remote system management

Write-Host "=== WinRM Administrative Setup Script ===" -ForegroundColor Green
Write-Host "Configuring Windows Remote Management for System Information GUI" -ForegroundColor Yellow
Write-Host ""

# Check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "❌ This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "✅ Running with Administrator privileges" -ForegroundColor Green
Write-Host ""

try {
    # Step 1: Enable PowerShell Remoting
    Write-Host "Step 1: Enabling PowerShell Remoting..." -ForegroundColor Cyan
    Enable-PSRemoting -Force -SkipNetworkProfileCheck
    Write-Host "✅ PowerShell Remoting enabled" -ForegroundColor Green
    Write-Host ""

    # Step 2: Configure WinRM
    Write-Host "Step 2: Configuring WinRM..." -ForegroundColor Cyan
    winrm quickconfig -force
    Write-Host "✅ WinRM configured" -ForegroundColor Green
    Write-Host ""

    # Step 3: Set WinRM service to automatic startup
    Write-Host "Step 3: Setting WinRM service startup..." -ForegroundColor Cyan
    Set-Service WinRM -StartupType Automatic
    Start-Service WinRM
    Write-Host "✅ WinRM service configured for automatic startup" -ForegroundColor Green
    Write-Host ""

    # Step 4: Configure trusted hosts for local network (optional)
    Write-Host "Step 4: Configuring trusted hosts..." -ForegroundColor Cyan
    Write-Host "This allows connections to computers not in a domain" -ForegroundColor Yellow
    $response = Read-Host "Configure trusted hosts for local network? (y/n)"
    
    if ($response -eq 'y' -or $response -eq 'Y') {
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "192.168.*,10.*,172.16.*,localhost,127.0.0.1" -Force
        Write-Host "✅ Trusted hosts configured for common local network ranges" -ForegroundColor Green
    } else {
        Write-Host "⏭️  Skipped trusted hosts configuration" -ForegroundColor Yellow
    }
    Write-Host ""

    # Step 5: Test the configuration
    Write-Host "Step 5: Testing WinRM configuration..." -ForegroundColor Cyan
    
    # Test WinRM connectivity
    $wsmanTest = Test-WSMan -ComputerName localhost
    Write-Host "✅ WinRM connectivity test passed" -ForegroundColor Green
    
    # Test PowerShell remoting
    $session = New-PSSession -ComputerName localhost
    $result = Invoke-Command -Session $session -ScriptBlock { $env:COMPUTERNAME }
    Remove-PSSession $session
    Write-Host "✅ PowerShell remoting test passed: Connected to $result" -ForegroundColor Green
    
    # Test CIM queries (what our GUI uses)
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName localhost
    Write-Host "✅ CIM query test passed: $($os.Caption)" -ForegroundColor Green
    Write-Host ""

    # Show current configuration
    Write-Host "Current WinRM Configuration:" -ForegroundColor Cyan
    Write-Host "-" * 30
    winrm get winrm/config
    Write-Host ""

    Write-Host "=== WinRM Setup Complete ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "🎉 SUCCESS! Your system is now configured for:" -ForegroundColor Yellow
    Write-Host "• Local system information gathering" -ForegroundColor Green
    Write-Host "• Remote system information gathering (to other configured computers)" -ForegroundColor Green
    Write-Host "• The System Information GUI will work for both local and remote queries" -ForegroundColor Green
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. Run the SystemInfoGUI.ps1 application" -ForegroundColor White
    Write-Host "2. Test local connections first (use localhost or $env:COMPUTERNAME)" -ForegroundColor White
    Write-Host "3. For remote computers, ensure they also have WinRM configured" -ForegroundColor White
    Write-Host "4. Use appropriate credentials when connecting to remote systems" -ForegroundColor White

} catch {
    Write-Host ""
    Write-Host "❌ Error during WinRM configuration: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "MANUAL STEPS TO TRY:" -ForegroundColor Yellow
    Write-Host "1. Open PowerShell as Administrator" -ForegroundColor White
    Write-Host "2. Run: Enable-PSRemoting -Force" -ForegroundColor Gray
    Write-Host "3. Run: winrm quickconfig -force" -ForegroundColor Gray
    Write-Host "4. Run: Set-Service WinRM -StartupType Automatic" -ForegroundColor Gray
}

Write-Host ""
Read-Host "Press Enter to continue"