# WinRM Configuration and Test Script
# This script helps configure WinRM for the System Information GUI application

Write-Host "=== WinRM Configuration and Testing Tool ===" -ForegroundColor Green
Write-Host "This script will help configure WinRM for remote system information gathering" -ForegroundColor Yellow
Write-Host ""

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check administrator status
$isAdmin = Test-Administrator
Write-Host "Running as Administrator: " -NoNewline
if ($isAdmin) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO" -ForegroundColor Red
    Write-Host "Note: Some configuration changes require administrator privileges" -ForegroundColor Yellow
}
Write-Host ""

# 1. Check WinRM Service Status
Write-Host "1. WINRM SERVICE STATUS" -ForegroundColor Cyan
Write-Host "-" * 30
$winrmService = Get-Service WinRM -ErrorAction SilentlyContinue
if ($winrmService) {
    Write-Host "Service Status: $($winrmService.Status)" -ForegroundColor $(if ($winrmService.Status -eq 'Running') {'Green'} else {'Red'})
    Write-Host "Startup Type: $($winrmService.StartType)"
} else {
    Write-Host "WinRM service not found!" -ForegroundColor Red
}
Write-Host ""

# 2. Test WinRM Connectivity
Write-Host "2. WINRM CONNECTIVITY TEST" -ForegroundColor Cyan
Write-Host "-" * 30
try {
    $wsmanTest = Test-WSMan -ComputerName localhost -ErrorAction Stop
    Write-Host "✅ WinRM is responding on localhost" -ForegroundColor Green
    Write-Host "   Protocol Version: $($wsmanTest.ProtocolVersion)"
    Write-Host "   Product Version: $($wsmanTest.ProductVersion)"
} catch {
    Write-Host "❌ WinRM connectivity test failed: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# 3. Check WinRM Listeners
Write-Host "3. WINRM LISTENERS" -ForegroundColor Cyan
Write-Host "-" * 30
try {
    if ($isAdmin) {
        $listeners = winrm enumerate winrm/config/listener 2>$null
        if ($listeners) {
            Write-Host "✅ WinRM listeners are configured" -ForegroundColor Green
            $listeners
        } else {
            Write-Host "❌ No WinRM listeners found" -ForegroundColor Red
        }
    } else {
        Write-Host "⚠️  Cannot enumerate listeners (requires admin privileges)" -ForegroundColor Yellow
        Write-Host "   Run as administrator to see detailed listener configuration" -ForegroundColor Gray
    }
} catch {
    Write-Host "❌ Error checking listeners: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# 4. Test PowerShell Remoting
Write-Host "4. POWERSHELL REMOTING TEST" -ForegroundColor Cyan
Write-Host "-" * 30
try {
    $session = New-PSSession -ComputerName localhost -ErrorAction Stop
    Write-Host "✅ PowerShell remoting session created successfully" -ForegroundColor Green
    
    # Test a simple command
    $result = Invoke-Command -Session $session -ScriptBlock { $env:COMPUTERNAME } -ErrorAction Stop
    Write-Host "✅ Remote command execution successful: $result" -ForegroundColor Green
    
    Remove-PSSession $session
} catch {
    Write-Host "❌ PowerShell remoting test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   This may indicate PowerShell remoting is not enabled" -ForegroundColor Yellow
}
Write-Host ""

# 5. Firewall Check
Write-Host "5. FIREWALL RULES CHECK" -ForegroundColor Cyan
Write-Host "-" * 30
try {
    $firewallRules = Get-NetFirewallRule -DisplayGroup "Windows Remote Management" -ErrorAction SilentlyContinue
    if ($firewallRules) {
        $enabledRules = $firewallRules | Where-Object { $_.Enabled -eq $true }
        Write-Host "✅ Windows Remote Management firewall rules found: $($firewallRules.Count) total, $($enabledRules.Count) enabled" -ForegroundColor Green
    } else {
        Write-Host "❌ No Windows Remote Management firewall rules found" -ForegroundColor Red
    }
} catch {
    Write-Host "⚠️  Could not check firewall rules: $($_.Exception.Message)" -ForegroundColor Yellow
}
Write-Host ""

# 6. Configuration Recommendations
Write-Host "6. CONFIGURATION RECOMMENDATIONS" -ForegroundColor Cyan
Write-Host "-" * 30

if ($winrmService.Status -ne 'Running') {
    Write-Host "❌ WinRM service is not running" -ForegroundColor Red
    Write-Host "   Resolution: Run 'Start-Service WinRM' as administrator" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "COMMANDS TO ENABLE WINRM (Run as Administrator):" -ForegroundColor Yellow
Write-Host "1. Enable PowerShell Remoting:" -ForegroundColor White
Write-Host "   Enable-PSRemoting -Force" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Configure WinRM:" -ForegroundColor White
Write-Host "   winrm quickconfig -force" -ForegroundColor Gray
Write-Host ""
Write-Host "3. Set trusted hosts (if connecting to non-domain computers):" -ForegroundColor White
Write-Host "   Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force" -ForegroundColor Gray
Write-Host ""
Write-Host "4. Start WinRM service:" -ForegroundColor White
Write-Host "   Start-Service WinRM" -ForegroundColor Gray
Write-Host "   Set-Service WinRM -StartupType Automatic" -ForegroundColor Gray
Write-Host ""

# 7. Test Our System Info Functions
Write-Host "7. TESTING SYSTEM INFO FUNCTIONS" -ForegroundColor Cyan
Write-Host "-" * 30
Write-Host "Testing the same functions used by our GUI application..." -ForegroundColor Yellow

try {
    # Test CIM connectivity
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName localhost -ErrorAction Stop
    Write-Host "✅ CIM/WMI queries working: $($os.Caption)" -ForegroundColor Green
    
    $cpu = Get-CimInstance -ClassName Win32_Processor -ComputerName localhost -ErrorAction Stop
    Write-Host "✅ CPU information accessible: $($cpu.Count) processor(s)" -ForegroundColor Green
    
    $memory = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName localhost -ErrorAction Stop
    Write-Host "✅ Memory information accessible: $($memory.Count) memory module(s)" -ForegroundColor Green
    
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName localhost -ErrorAction Stop | Where-Object { $_.DriveType -eq 3 }
    Write-Host "✅ Disk information accessible: $($disks.Count) drive(s)" -ForegroundColor Green
    
} catch {
    Write-Host "❌ System information query failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Configuration Check Complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "SUMMARY FOR SYSTEM INFORMATION GUI:" -ForegroundColor Yellow
Write-Host "• Local system information gathering: " -NoNewline
if ($os) { Write-Host "✅ WORKING" -ForegroundColor Green } else { Write-Host "❌ FAILED" -ForegroundColor Red }

Write-Host "• Remote connectivity capability: " -NoNewline
try {
    Test-WSMan -ComputerName localhost -ErrorAction Stop | Out-Null
    Write-Host "✅ READY" -ForegroundColor Green
} catch {
    Write-Host "❌ NEEDS CONFIGURATION" -ForegroundColor Red
}

Write-Host ""
Write-Host "The System Information GUI should work for:" -ForegroundColor White
Write-Host "• Local computer queries (always available)" -ForegroundColor Green
Write-Host "• Remote computer queries (requires WinRM configuration on target)" -ForegroundColor $(if ($wsmanTest) {'Green'} else {'Yellow'})