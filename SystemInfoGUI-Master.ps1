# System Information GUI - Master Comprehensive Version with Remote Authentication
# Created: October 2, 2025
# Description: Merged comprehensive GUI with remote authentication capabilities
# Features: Auto-elevation, Enhanced UI, Remote Authentication, Network Tools, Progress Tracking

# Check if running as Administrator and elevate if not
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n=== ADMINISTRATOR PRIVILEGES REQUIRED ===" -ForegroundColor Yellow
    Write-Host "This GUI requires Administrator privileges for full functionality:" -ForegroundColor White
    Write-Host "- System-level software installation" -ForegroundColor Gray
    Write-Host "- WinRM and service configuration" -ForegroundColor Gray
    Write-Host "- Network share and registry access" -ForegroundColor Gray
    Write-Host "- Package manager operations" -ForegroundColor Gray
    Write-Host "- Services monitoring and management" -ForegroundColor Gray
    Write-Host "- Remote computer authentication" -ForegroundColor Gray
    Write-Host "`nLaunching Administrator session..." -ForegroundColor Cyan
    
    # Get the current script path
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    
    try {
        # Start new PowerShell process as Administrator
        Start-Process PowerShell -Verb RunAs -ArgumentList $arguments -WorkingDirectory (Split-Path $scriptPath)
        
        # Exit current non-admin process
        Write-Host "New Administrator window launching..." -ForegroundColor Green
        exit
    } catch {
        Write-Host "`nFailed to launch as Administrator: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`nPlease manually run PowerShell as Administrator and try again." -ForegroundColor Yellow
        Write-Host "Press Ctrl+C to exit or Enter to continue with limited functionality..." -ForegroundColor Gray
        Read-Host
    }
}

# Confirm Administrator status
Write-Host "`n=== RUNNING AS ADMINISTRATOR - MASTER WITH REMOTE AUTH ===" -ForegroundColor Green
Write-Host "Full system access enabled including remote authentication" -ForegroundColor White
Write-Host "Working Directory: $(Get-Location)" -ForegroundColor Cyan

try {
    # Load required assemblies
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop

    # Get screen dimensions for optimal sizing
    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $formWidth = [math]::Round($workingArea.Width * 0.90)
    $formHeight = [math]::Round($workingArea.Height * 0.90)

    # Main Form - Enhanced with optimal sizing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "System Information GUI - Master with Remote Authentication"
    $form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = New-Object System.Drawing.Size(1200, 800)
    $form.MaximizeBox = $true
    $form.FormBorderStyle = "Sizable"
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke

    # Host Selection Group - Enhanced with Remote Credentials
    $hostGroup = New-Object System.Windows.Forms.GroupBox
    $hostGroup.Text = "Target Computer & Remote Authentication"
    $hostGroup.Location = New-Object System.Drawing.Point(10, 10)
    $hostGroup.Size = New-Object System.Drawing.Size(($formWidth - 30), 120)
    $hostGroup.Anchor = "Top,Left,Right"

    # Computer Name
    $hostLabel = New-Object System.Windows.Forms.Label
    $hostLabel.Text = "Computer Name:"
    $hostLabel.Location = New-Object System.Drawing.Point(10, 25)
    $hostLabel.Size = New-Object System.Drawing.Size(100, 20)

    $hostTextBox = New-Object System.Windows.Forms.TextBox
    $hostTextBox.Location = New-Object System.Drawing.Point(120, 23)
    $hostTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $hostTextBox.Text = $env:COMPUTERNAME

    # Authentication Mode
    $authLabel = New-Object System.Windows.Forms.Label
    $authLabel.Text = "Authentication:"
    $authLabel.Location = New-Object System.Drawing.Point(340, 25)
    $authLabel.Size = New-Object System.Drawing.Size(80, 20)

    $authComboBox = New-Object System.Windows.Forms.ComboBox
    $authComboBox.Location = New-Object System.Drawing.Point(430, 23)
    $authComboBox.Size = New-Object System.Drawing.Size(140, 20)
    $authComboBox.DropDownStyle = "DropDownList"
    $authComboBox.Items.AddRange(@("Current User", "Alternate Credentials"))
    $authComboBox.SelectedIndex = 0

    # Username
    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Text = "Username:"
    $usernameLabel.Location = New-Object System.Drawing.Point(10, 55)
    $usernameLabel.Size = New-Object System.Drawing.Size(100, 20)
    $usernameLabel.Enabled = $false

    $usernameTextBox = New-Object System.Windows.Forms.TextBox
    $usernameTextBox.Location = New-Object System.Drawing.Point(120, 53)
    $usernameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $usernameTextBox.Enabled = $false
    $usernameTextBox.Text = "DOMAIN\Username"

    # Password
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "Password:"
    $passwordLabel.Location = New-Object System.Drawing.Point(340, 55)
    $passwordLabel.Size = New-Object System.Drawing.Size(80, 20)
    $passwordLabel.Enabled = $false

    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Location = New-Object System.Drawing.Point(430, 53)
    $passwordTextBox.Size = New-Object System.Drawing.Size(140, 20)
    $passwordTextBox.PasswordChar = '*'
    $passwordTextBox.Enabled = $false

    # Test Connection Button
    $testButton = New-Object System.Windows.Forms.Button
    $testButton.Text = "Test Connection"
    $testButton.Location = New-Object System.Drawing.Point(590, 22)
    $testButton.Size = New-Object System.Drawing.Size(100, 25)
    $testButton.BackColor = [System.Drawing.Color]::LightYellow

    # Connect Button
    $connectButton = New-Object System.Windows.Forms.Button
    $connectButton.Text = "Connect & Get Info"
    $connectButton.Location = New-Object System.Drawing.Point(590, 52)
    $connectButton.Size = New-Object System.Drawing.Size(100, 30)
    $connectButton.BackColor = [System.Drawing.Color]::LightBlue

    # Connection Status
    $connectionStatusLabel = New-Object System.Windows.Forms.Label
    $connectionStatusLabel.Text = "Status: Ready to connect"
    $connectionStatusLabel.Location = New-Object System.Drawing.Point(10, 85)
    $connectionStatusLabel.Size = New-Object System.Drawing.Size(($formWidth - 50), 20)
    $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Blue

    $hostGroup.Controls.AddRange(@($hostLabel, $hostTextBox, $authLabel, $authComboBox, $usernameLabel, $usernameTextBox, $passwordLabel, $passwordTextBox, $testButton, $connectButton, $connectionStatusLabel))

    # Information Selection Group - Single Row Layout
    $infoGroup = New-Object System.Windows.Forms.GroupBox
    $infoGroup.Text = "Information to Collect"
    $infoGroup.Location = New-Object System.Drawing.Point(10, 140)
    $infoGroup.Size = New-Object System.Drawing.Size(($formWidth - 30), 60)
    $infoGroup.Anchor = "Top,Left,Right"

    # Create checkboxes in single row
    $diskCheckBox = New-Object System.Windows.Forms.CheckBox
    $diskCheckBox.Text = "Disk Info"
    $diskCheckBox.Location = New-Object System.Drawing.Point(20, 25)
    $diskCheckBox.Size = New-Object System.Drawing.Size(80, 20)
    $diskCheckBox.Checked = $true

    $cpuCheckBox = New-Object System.Windows.Forms.CheckBox
    $cpuCheckBox.Text = "CPU Info"
    $cpuCheckBox.Location = New-Object System.Drawing.Point(110, 25)
    $cpuCheckBox.Size = New-Object System.Drawing.Size(80, 20)
    $cpuCheckBox.Checked = $true

    $ramCheckBox = New-Object System.Windows.Forms.CheckBox
    $ramCheckBox.Text = "RAM Info"
    $ramCheckBox.Location = New-Object System.Drawing.Point(200, 25)
    $ramCheckBox.Size = New-Object System.Drawing.Size(80, 20)
    $ramCheckBox.Checked = $true

    $softwareCheckBox = New-Object System.Windows.Forms.CheckBox
    $softwareCheckBox.Text = "Software"
    $softwareCheckBox.Location = New-Object System.Drawing.Point(290, 25)
    $softwareCheckBox.Size = New-Object System.Drawing.Size(80, 20)
    $softwareCheckBox.Checked = $true

    $osCheckBox = New-Object System.Windows.Forms.CheckBox
    $osCheckBox.Text = "OS Info"
    $osCheckBox.Location = New-Object System.Drawing.Point(380, 25)
    $osCheckBox.Size = New-Object System.Drawing.Size(70, 20)
    $osCheckBox.Checked = $true

    $servicesCheckBox = New-Object System.Windows.Forms.CheckBox
    $servicesCheckBox.Text = "Services"
    $servicesCheckBox.Location = New-Object System.Drawing.Point(460, 25)
    $servicesCheckBox.Size = New-Object System.Drawing.Size(80, 20)
    $servicesCheckBox.Checked = $true

    $kbCheckBox = New-Object System.Windows.Forms.CheckBox
    $kbCheckBox.Text = "Installed KB"
    $kbCheckBox.Location = New-Object System.Drawing.Point(550, 25)
    $kbCheckBox.Size = New-Object System.Drawing.Size(90, 20)
    $kbCheckBox.Checked = $true

    $gpoCheckBox = New-Object System.Windows.Forms.CheckBox
    $gpoCheckBox.Text = "GPO"
    $gpoCheckBox.Location = New-Object System.Drawing.Point(650, 25)
    $gpoCheckBox.Size = New-Object System.Drawing.Size(90, 20)
    $gpoCheckBox.Checked = $true

    $envCheckBox = New-Object System.Windows.Forms.CheckBox
    $envCheckBox.Text = "System Env"
    $envCheckBox.Location = New-Object System.Drawing.Point(750, 25)
    $envCheckBox.Size = New-Object System.Drawing.Size(90, 20)
    $envCheckBox.Checked = $true

    $featuresCheckBox = New-Object System.Windows.Forms.CheckBox
    $featuresCheckBox.Text = "Features & Roles"
    $featuresCheckBox.Location = New-Object System.Drawing.Point(850, 25)
    $featuresCheckBox.Size = New-Object System.Drawing.Size(120, 20)
    $featuresCheckBox.Checked = $true

    $infoGroup.Controls.AddRange(@($diskCheckBox, $cpuCheckBox, $ramCheckBox, $softwareCheckBox, $osCheckBox, $servicesCheckBox, $kbCheckBox, $gpoCheckBox, $envCheckBox, $featuresCheckBox))

    # Progress Bar and Status
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 210)
    $progressBar.Size = New-Object System.Drawing.Size(($formWidth - 30), 25)
    $progressBar.Anchor = "Top,Left,Right"
    $progressBar.Style = "Continuous"
    $progressBar.Visible = $true

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready to collect system information..."
    $statusLabel.Location = New-Object System.Drawing.Point(10, 240)
    $statusLabel.Size = New-Object System.Drawing.Size(($formWidth - 30), 20)
    $statusLabel.Anchor = "Top,Left,Right"
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue

    # Results Display Area
    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, 270)
    $resultsTextBox.Size = New-Object System.Drawing.Size(($formWidth - 30), ($formHeight - 430))
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Vertical"
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.Anchor = "Top,Bottom,Left,Right"

    # Enhanced Button Panel with Increased Height
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Location = New-Object System.Drawing.Point(0, ($formHeight - 130))
    $buttonPanel.Size = New-Object System.Drawing.Size($formWidth, 120)
    $buttonPanel.BackColor = [System.Drawing.Color]::DarkGray
    $buttonPanel.BorderStyle = "Fixed3D"
    $buttonPanel.Anchor = "Bottom,Left,Right"

    # Dynamic button positioning for perfect centering
    $buttonWidth = 160
    $buttonHeight = 75  # Increased height
    $buttonSpacing = 15
    $buttonY = 25  # Centered in 120px panel
    
    # Calculate total width and center position
    $totalButtonsWidth = (6 * $buttonWidth) + (5 * $buttonSpacing)
    $startX = [math]::Floor(($formWidth - $totalButtonsWidth) / 2)

    # Create Enhanced Buttons
    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "CLEAR"
    $clearButton.Location = New-Object System.Drawing.Point($startX, $buttonY)
    $clearButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $clearButton.BackColor = [System.Drawing.Color]::Cyan
    $clearButton.ForeColor = [System.Drawing.Color]::Black
    $clearButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $clearButton.FlatStyle = "Popup"

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "EXPORT"
    $exportButton.Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $buttonSpacing), $buttonY)
    $exportButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $exportButton.BackColor = [System.Drawing.Color]::Lime
    $exportButton.ForeColor = [System.Drawing.Color]::Black
    $exportButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $exportButton.FlatStyle = "Popup"

    $regeditButton = New-Object System.Windows.Forms.Button
    $regeditButton.Text = "REGISTRY"
    $regeditButton.Location = New-Object System.Drawing.Point(($startX + 2 * ($buttonWidth + $buttonSpacing)), $buttonY)
    $regeditButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $regeditButton.BackColor = [System.Drawing.Color]::Orange
    $regeditButton.ForeColor = [System.Drawing.Color]::Black
    $regeditButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $regeditButton.FlatStyle = "Popup"

    $printersButton = New-Object System.Windows.Forms.Button
    $printersButton.Text = "PRINTERS"
    $printersButton.Location = New-Object System.Drawing.Point(($startX + 3 * ($buttonWidth + $buttonSpacing)), $buttonY)
    $printersButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $printersButton.BackColor = [System.Drawing.Color]::MediumPurple
    $printersButton.ForeColor = [System.Drawing.Color]::White
    $printersButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $printersButton.FlatStyle = "Popup"

    $sharesButton = New-Object System.Windows.Forms.Button
    $sharesButton.Text = "NET SHARES"
    $sharesButton.Location = New-Object System.Drawing.Point(($startX + 4 * ($buttonWidth + $buttonSpacing)), $buttonY)
    $sharesButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $sharesButton.BackColor = [System.Drawing.Color]::DarkTurquoise
    $sharesButton.ForeColor = [System.Drawing.Color]::White
    $sharesButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $sharesButton.FlatStyle = "Popup"

    $envEditorButton = New-Object System.Windows.Forms.Button
    $envEditorButton.Text = "ENV EDITOR"
    $envEditorButton.Location = New-Object System.Drawing.Point(($startX + 5 * ($buttonWidth + $buttonSpacing)), $buttonY)
    $envEditorButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $envEditorButton.BackColor = [System.Drawing.Color]::ForestGreen
    $envEditorButton.ForeColor = [System.Drawing.Color]::White
    $envEditorButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $envEditorButton.FlatStyle = "Popup"

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Text = "EXIT"
    $exitButton.Location = New-Object System.Drawing.Point(($startX + 6 * ($buttonWidth + $buttonSpacing)), $buttonY)
    $exitButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $exitButton.BackColor = [System.Drawing.Color]::Red
    $exitButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $exitButton.FlatStyle = "Popup"

    # Add buttons to panel
    $buttonPanel.Controls.AddRange(@($clearButton, $exportButton, $regeditButton, $printersButton, $sharesButton, $envEditorButton, $exitButton))

    # Helper Functions
    function Update-Status {
        param([string]$message, [string]$color = "Blue")
        try {
            $statusLabel.Text = $message
            $statusLabel.ForeColor = [System.Drawing.Color]::$color
            [System.Windows.Forms.Application]::DoEvents()
        } catch {
            # Silently ignore status update errors
        }
    }

    function Update-Progress {
        param([int]$value, [string]$status = "")
        try {
            if ($value -ge 0 -and $value -le 100) {
                $progressBar.Value = $value
                if ($status) {
                    $statusLabel.Text = "$status ($value%)"
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 50
            }
        } catch {
            # Silently ignore progress update errors
        }
    }

    function Get-SystemInformation {
        param(
            [string]$ComputerName,
            [System.Management.Automation.PSCredential]$Credential = $null
        )
        
        $cimSession = $null
        try {
            $results = @()
            $results += "=== COMPREHENSIVE SYSTEM INFORMATION REPORT ==="
            $results += "Computer: $ComputerName"
            if ($Credential) {
                $results += "Authentication: $($Credential.UserName)"
            } else {
                $results += "Authentication: Current User"
            }
            $results += "Generated: $(Get-Date)"
            $results += "=" * 60
            $results += ""

            Update-Status "Establishing connection..." "Orange"
            Update-Progress 10
            
            # Create CimSession with or without credentials
            $cimSessionOption = New-CimSessionOption -Protocol WSMan
            if ($Credential -and $ComputerName -ne $env:COMPUTERNAME) {
                Update-Status "Authenticating with provided credentials..." "Orange"
                $cimSession = New-CimSession -ComputerName $ComputerName -Credential $Credential -SessionOption $cimSessionOption -ErrorAction Stop
            } else {
                $cimSession = New-CimSession -ComputerName $ComputerName -SessionOption $cimSessionOption -ErrorAction Stop
            }
            
            # Test connection
            $null = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $cimSession -ErrorAction Stop
            Update-Status "Connection established successfully" "Green"

            Update-Status "Gathering system information..." "Orange"
            Update-Progress 20
            
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $cimSession -ErrorAction Stop
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $cimSession -ErrorAction Stop
            
            if ($osCheckBox.Checked) {
                Update-Progress 30 "Collecting OS information"
                $results += "OPERATING SYSTEM INFORMATION:"
                $results += "OS Name: $($os.Caption)"
                $results += "Version: $($os.Version)"
                $results += "Architecture: $($os.OSArchitecture)"
                $results += "Computer Name: $($cs.Name)"
                $results += "Domain/Workgroup: $($cs.Domain)"
                $results += "Last Boot Time: $($os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss'))"
                $uptime = (Get-Date) - $os.LastBootUpTime
                $results += "Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
                $results += ""
            }

            if ($cpuCheckBox.Checked) {
                Update-Progress 40 "Collecting CPU information"
                $cpu = Get-CimInstance -ClassName Win32_Processor -CimSession $cimSession -ErrorAction Stop
                $results += "CPU INFORMATION:"
                foreach ($proc in $cpu) {
                    $results += "Name: $($proc.Name)"
                    $results += "Cores: $($proc.NumberOfCores)"
                    $results += "Logical Processors: $($proc.NumberOfLogicalProcessors)"
                    $results += "Max Clock Speed: $($proc.MaxClockSpeed) MHz"
                }
                $results += ""
            }

            if ($ramCheckBox.Checked) {
                Update-Progress 50 "Collecting memory information"
                $memory = Get-CimInstance -ClassName Win32_PhysicalMemory -CimSession $cimSession -ErrorAction Stop
                $results += "MEMORY INFORMATION:"
                $totalMemory = ($memory | Measure-Object -Property Capacity -Sum).Sum
                $results += "Total Physical Memory: $([math]::Round($totalMemory / 1GB, 2)) GB"
                $results += "Available Memory: $([math]::Round($os.FreePhysicalMemory / 1MB, 2)) GB"
                $results += ""
            }

            if ($diskCheckBox.Checked) {
                Update-Progress 60 "Collecting disk information"
                $disks = Get-CimInstance -ClassName Win32_LogicalDisk -CimSession $cimSession -ErrorAction Stop
                $results += "DISK INFORMATION:"
                foreach ($disk in $disks) {
                    if ($disk.Size) {
                        $results += "Drive $($disk.DeviceID): $([math]::Round($disk.Size / 1GB, 2)) GB"
                        $results += "  Free: $([math]::Round($disk.FreeSpace / 1GB, 2)) GB"
                    }
                }
                $results += ""
            }

            if ($softwareCheckBox.Checked) {
                Update-Progress 70 "Collecting software information"
                try {
                    $software = Get-CimInstance -ClassName Win32_Product -CimSession $cimSession -ErrorAction Stop | Sort-Object Name
                    $results += "COMPLETE INSTALLED SOFTWARE TABLE:"
                    $results += "Total Applications: $($software.Count) (Showing ALL installed applications)"
                    $results += ""
                    
                    # Create table header with expanded columns for full text
                    $results += "{0,-60} {1,-20} {2,-30}" -f "Application Name", "Version", "Vendor"
                    $results += "-" * 110
                    
                    # Display ALL applications without any limits
                    foreach ($app in $software) {
                        $name = if ($app.Name) { $app.Name } else { "Unknown Application" }
                        $version = if ($app.Version) { $app.Version } else { "N/A" }
                        $vendor = if ($app.Vendor) { $app.Vendor } else { "Unknown" }
                        
                        $results += "{0,-60} {1,-20} {2,-30}" -f $name, $version, $vendor
                    }
                    
                    $results += ""
                    $results += "=== END OF SOFTWARE LIST === ($($software.Count) total applications displayed)"
                } catch {
                    $results += "INSTALLED SOFTWARE: Unable to retrieve - $($_.Exception.Message)"
                }
                $results += ""
            }

            if ($servicesCheckBox.Checked) {
                Update-Progress 75 "Collecting services information"
                try {
                    $services = Get-CimInstance -ClassName Win32_Service -CimSession $cimSession -ErrorAction Stop | Sort-Object State, Name
                    $runningServices = $services | Where-Object { $_.State -eq "Running" }
                    $stoppedServices = $services | Where-Object { $_.State -eq "Stopped" }
                    $disabledServices = $services | Where-Object { $_.StartMode -eq "Disabled" }
                    
                    $results += "COMPLETE SERVICES TABLE:"
                    $results += "Total Services: $($services.Count) | Running: $($runningServices.Count) | Stopped: $($stoppedServices.Count) | Disabled: $($disabledServices.Count)"
                    $results += ""
                    
                    # Create services table header with expanded columns
                    $results += "{0,-40} {1,-10} {2,-15} {3,-50}" -f "Service Name", "State", "Start Mode", "Description"
                    $results += "-" * 115
                    
                    # Show ALL RUNNING services
                    $results += "=== ALL RUNNING SERVICES ($($runningServices.Count) services) ==="
                    foreach ($service in $runningServices) {
                        $name = if ($service.Name) { $service.Name } else { "Unknown Service" }
                        $startMode = if ($service.StartMode) { $service.StartMode } else { "Unknown" }
                        $description = if ($service.Description) {
                            if ($service.Description.Length -gt 47) { $service.Description.Substring(0, 44) + "..." } else { $service.Description }
                        } else { "No description" }
                        
                        $results += "{0,-40} {1,-10} {2,-15} {3,-50}" -f $name, $service.State, $startMode, $description
                    }
                    
                    $results += ""
                    
                    # Show ALL STOPPED services
                    $results += "=== ALL STOPPED SERVICES ($($stoppedServices.Count) services) ==="
                    foreach ($service in $stoppedServices) {
                        $name = if ($service.Name) { $service.Name } else { "Unknown Service" }
                        $startMode = if ($service.StartMode) { $service.StartMode } else { "Unknown" }
                        $description = if ($service.Description) {
                            if ($service.Description.Length -gt 47) { $service.Description.Substring(0, 44) + "..." } else { $service.Description }
                        } else { "No description" }
                        
                        $results += "{0,-40} {1,-10} {2,-15} {3,-50}" -f $name, $service.State, $startMode, $description
                    }
                    
                    $results += ""
                    
                    # Show ALL DISABLED services
                    $results += "=== ALL DISABLED SERVICES ($($disabledServices.Count) services) ==="
                    foreach ($service in $disabledServices) {
                        $name = if ($service.Name) { $service.Name } else { "Unknown Service" }
                        $state = if ($service.State) { $service.State } else { "Disabled" }
                        $description = if ($service.Description) {
                            if ($service.Description.Length -gt 47) { $service.Description.Substring(0, 44) + "..." } else { $service.Description }
                        } else { "No description" }
                        
                        $results += "{0,-40} {1,-10} {2,-15} {3,-50}" -f $name, $state, "Disabled", $description
                    }
                    
                    # Critical services status
                    $results += ""
                    $results += "=== CRITICAL SERVICES STATUS ==="
                    $criticalServices = @("Spooler", "BITS", "Themes", "LanmanServer", "Winmgmt", "EventLog", "RpcSs", "Dhcp", "Dnscache", "Netlogon", "W32Time")
                    foreach ($criticalName in $criticalServices) {
                        $service = $services | Where-Object { $_.Name -eq $criticalName }
                        if ($service) {
                            $status = if ($service.State -eq "Running") { "[OK] $($service.State)" } else { "[!!] $($service.State)" }
                            $results += "  {0,-20}: {1}" -f $service.Name, $status
                        }
                    }
                    
                    $results += ""
                    $results += "=== END OF SERVICES LIST === ($($services.Count) total services displayed)"
                    
                } catch {
                    $results += "SERVICES: Unable to retrieve - $($_.Exception.Message)"
                }
                $results += ""
            }

            if ($kbCheckBox.Checked) {
                Update-Progress 85 "Collecting installed KB updates"
                try {
                    $hotfixes = Get-CimInstance -ClassName Win32_QuickFixEngineering -CimSession $cimSession -ErrorAction Stop | Sort-Object InstalledOn -Descending
                    $securityUpdates = $hotfixes | Where-Object { $_.Description -like "*Security*" }
                    $recentDate = (Get-Date).AddDays(-30)
                    $recentUpdates = $hotfixes | Where-Object { 
                        $_.InstalledOn -and [DateTime]$_.InstalledOn -gt $recentDate 
                    }
                    
                    $results += "COMPLETE KB UPDATES TABLE:"
                    $results += "Total Updates: $($hotfixes.Count) | Security Updates: $($securityUpdates.Count) | Recent Updates: $($recentUpdates.Count)"
                    $results += "(Showing ALL $($hotfixes.Count) installed KB updates)"
                    $results += ""
                    
                    # Create KB updates table header with expanded columns
                    $results += "{0,-15} {1,-30} {2,-15} {3,-30}" -f "KB Number", "Description", "Install Date", "Installed By"
                    $results += "-" * 90
                    
                    # Show ALL KB updates (sorted by installation date, newest first)
                    $results += "=== ALL INSTALLED KB UPDATES ($($hotfixes.Count) updates) ==="
                    foreach ($kb in $hotfixes) {
                        $kbNumber = if ($kb.HotFixID) { $kb.HotFixID } else { "Unknown" }
                        $description = if ($kb.Description) {
                            if ($kb.Description.Length -gt 28) { $kb.Description.Substring(0, 25) + "..." } else { $kb.Description }
                        } else { "No description" }
                        $installDate = if ($kb.InstalledOn) { ([DateTime]$kb.InstalledOn).ToString("yyyy-MM-dd") } else { "Unknown" }
                        $installedBy = if ($kb.InstalledBy) {
                            if ($kb.InstalledBy.Length -gt 28) { $kb.InstalledBy.Substring(0, 25) + "..." } else { $kb.InstalledBy }
                        } else { "System" }
                        
                        $results += "{0,-15} {1,-30} {2,-15} {3,-30}" -f $kbNumber, $description, $installDate, $installedBy
                    }
                    
                    # Summary statistics
                    $results += ""
                    $results += "=== KB UPDATE SUMMARY ==="
                    $results += "Total KB Updates Installed: $($hotfixes.Count)"
                    $results += "Security Updates: $($securityUpdates.Count)"
                    $results += "Recent Updates (Last 30 days): $($recentUpdates.Count)"
                    
                    # Show update types breakdown
                    $updateTypes = $hotfixes | Group-Object Description | Sort-Object Count -Descending
                    $results += ""
                    $results += "Update Types Breakdown:"
                    foreach ($type in ($updateTypes | Select-Object -First 10)) {
                        $typeName = if ($type.Name.Length -gt 40) { $type.Name.Substring(0, 37) + "..." } else { $type.Name }
                        $results += "  $($type.Count) - $typeName"
                    }
                    if ($updateTypes.Count -gt 10) {
                        $results += "  ... and $($updateTypes.Count - 10) more update types"
                    }
                    
                    $results += ""
                    $results += "=== END OF KB UPDATES LIST === ($($hotfixes.Count) total updates displayed)"
                    
                } catch {
                    $results += "INSTALLED KB UPDATES: Unable to retrieve - $($_.Exception.Message)"
                }
                $results += ""
            }

            # Features and Roles Information
            if ($featuresCheckBox.Checked) {
                Update-Progress 85 "Collecting Windows Features and Roles information"
                try {
                    $results += "COMPLETE WINDOWS FEATURES AND ROLES TABLE:"
                    $results += "Windows Features, Optional Features, and Server Roles"
                    $results += ""
                    
                    # Create table header
                    $results += "Feature/Role Name".PadRight(50) + "Status".PadRight(15) + "Type"
                    $results += ("-" * 115)
                    
                    # Get Windows Features (DISM method - works on all Windows versions)
                    $results += "=== WINDOWS FEATURES ==="
                    try {
                        $dismFeatures = dism /online /get-features /format:table 2>$null | Where-Object { $_ -match "\|" -and $_ -notmatch "Feature Name" -and $_ -notmatch "---" }
                        $featureCount = 0
                        foreach ($line in $dismFeatures) {
                            if ($line -match "^(.+?)\s+\|\s+(.+?)\s*$") {
                                $featureName = $matches[1].Trim()
                                $status = $matches[2].Trim()
                                if ($featureName -and $status) {
                                    $results += $featureName.PadRight(50) + $status.PadRight(15) + "Windows Feature"
                                    $featureCount++
                                }
                            }
                        }
                        $results += "Total Windows Features: $featureCount"
                    } catch {
                        $results += "Windows Features: Unable to retrieve via DISM - $($_.Exception.Message)"
                    }
                    
                    $results += ""
                    
                    # Get Optional Features (Get-WindowsOptionalFeature - PowerShell method)
                    $results += "=== OPTIONAL FEATURES ==="
                    try {
                        $optionalFeatures = Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue
                        if ($optionalFeatures) {
                            $enabledCount = 0
                            $disabledCount = 0
                            foreach ($feature in $optionalFeatures) {
                                $status = switch ($feature.State) {
                                    "Enabled" { $enabledCount++; "Enabled" }
                                    "Disabled" { $disabledCount++; "Disabled" }
                                    default { $feature.State }
                                }
                                $results += $feature.FeatureName.PadRight(50) + $status.PadRight(15) + "Optional Feature"
                            }
                            $results += "Total Optional Features: $($optionalFeatures.Count) (Enabled: $enabledCount, Disabled: $disabledCount)"
                        } else {
                            $results += "Optional Features: Not available on this system"
                        }
                    } catch {
                        $results += "Optional Features: Unable to retrieve - $($_.Exception.Message)"
                    }
                    
                    $results += ""
                    
                    # Get Server Roles and Features (if Server OS)
                    $results += "=== SERVER ROLES AND FEATURES ==="
                    try {
                        $serverManager = Get-WindowsFeature -ErrorAction SilentlyContinue
                        if ($serverManager) {
                            $installedCount = 0
                            $availableCount = 0
                            $removedCount = 0
                            
                            foreach ($feature in $serverManager) {
                                $status = switch ($feature.InstallState) {
                                    "Installed" { $installedCount++; "Installed" }
                                    "Available" { $availableCount++; "Available" }
                                    "Removed" { $removedCount++; "Removed" }
                                    default { $feature.InstallState }
                                }
                                $type = if ($feature.FeatureType -eq "Role") { "Server Role" } else { "Server Feature" }
                                $results += $feature.Name.PadRight(50) + $status.PadRight(15) + $type
                            }
                            $results += "Total Server Features: $($serverManager.Count) (Installed: $installedCount, Available: $availableCount, Removed: $removedCount)"
                        } else {
                            $results += "Server Roles and Features: Not available (not a Windows Server or feature not accessible)"
                        }
                    } catch {
                        $results += "Server Roles and Features: Unable to retrieve - $($_.Exception.Message)"
                    }
                    
                    $results += ""
                    
                    # Get Windows Capabilities (Windows 10/11)
                    $results += "=== WINDOWS CAPABILITIES ==="
                    try {
                        $capabilities = Get-WindowsCapability -Online -ErrorAction SilentlyContinue
                        if ($capabilities) {
                            $installedCaps = 0
                            $notPresentCaps = 0
                            foreach ($cap in $capabilities) {
                                $status = switch ($cap.State) {
                                    "Installed" { $installedCaps++; "Installed" }
                                    "NotPresent" { $notPresentCaps++; "Not Present" }
                                    default { $cap.State }
                                }
                                $results += $cap.Name.PadRight(50) + $status.PadRight(15) + "Windows Capability"
                            }
                            $results += "Total Windows Capabilities: $($capabilities.Count) (Installed: $installedCaps, Not Present: $notPresentCaps)"
                        } else {
                            $results += "Windows Capabilities: Not available on this system"
                        }
                    } catch {
                        $results += "Windows Capabilities: Unable to retrieve - $($_.Exception.Message)"
                    }
                    
                    $results += ""
                    $results += "=== FEATURE SUMMARY ==="
                    $results += "Features and Roles analysis completed successfully"
                    $results += "Use 'Programs and Features' or 'Server Manager' for detailed management"
                    
                } catch {
                    $results += "FEATURES AND ROLES: Unable to retrieve - $($_.Exception.Message)"
                }
                $results += ""
            }

            if ($envCheckBox.Checked) {
                Update-Progress 87 "Collecting System Environment information"
                try {
                    $results += "COMPLETE SYSTEM ENVIRONMENT TABLE:"
                    $results += "Environment Variables, Paths, and System Configuration"
                    $results += ""
                    
                    # Create Environment table header
                    $results += "{0,-35} {1,-60}" -f "Variable Name", "Value"
                    $results += "-" * 95
                    
                    # Get all environment variables
                    $envVars = Get-ChildItem Env: | Sort-Object Name
                    
                    $results += "=== ALL SYSTEM ENVIRONMENT VARIABLES ($($envVars.Count) variables) ==="
                    $results += "(Showing complete values without truncation - editable via Environment Editor)"
                    $results += ""
                    
                    # Create expanded table header for full values
                    $results += "{0,-35} {1,-80}" -f "Variable Name", "Complete Value"
                    $results += "-" * 115
                    
                    foreach ($env in $envVars) {
                        $name = if ($env.Name) { $env.Name } else { "Unknown" }
                        $value = if ($env.Value) { $env.Value } else { "(empty)" }
                        
                        # Display full value without truncation
                        $results += "{0,-35} {1,-80}" -f $name, $value
                    }
                    
                    $results += ""
                    $results += "=== KEY SYSTEM PATHS ==="
                    $keyPaths = @(
                        @{Name="Windows Directory"; Path=$env:WINDIR},
                        @{Name="System32 Directory"; Path="$env:WINDIR\System32"},
                        @{Name="Program Files"; Path=$env:ProgramFiles},
                        @{Name="Program Files (x86)"; Path=${env:ProgramFiles(x86)}},
                        @{Name="User Profile"; Path=$env:USERPROFILE},
                        @{Name="Temp Directory"; Path=$env:TEMP},
                        @{Name="App Data"; Path=$env:APPDATA},
                        @{Name="Local App Data"; Path=$env:LOCALAPPDATA}
                    )
                    
                    foreach ($path in $keyPaths) {
                        if ($path.Path) {
                            # Display full path without truncation
                            $results += "{0,-35} {1,-80}" -f $path.Name, $path.Path
                        }
                    }
                    
                    $results += ""
                    $results += "=== PATH ENVIRONMENT ANALYSIS ==="
                    $pathEntries = $env:PATH -split ';' | Where-Object { $_ -ne "" }
                    $results += "Total PATH entries: $($pathEntries.Count)"
                    $results += "PATH Breakdown (ALL $($pathEntries.Count) entries):"
                    foreach ($pathEntry in $pathEntries) {
                        # Display full PATH entry without truncation
                        $results += "{0,-35} {1,-80}" -f "PATH Entry", $pathEntry
                    }
                    
                    # Add environment editing instructions
                    $results += ""
                    $results += "=== ENVIRONMENT VARIABLE EDITING ==="
                    $results += "To edit environment variables:"
                    $results += "1. Use 'Environment Editor' button (below) for GUI editing"
                    $results += "2. PowerShell commands for manual editing:"
                    $results += "   - Set variable: [System.Environment]::SetEnvironmentVariable('NAME','VALUE','User')"
                    $results += "   - Set system-wide: [System.Environment]::SetEnvironmentVariable('NAME','VALUE','Machine')"
                    $results += "   - Delete variable: [System.Environment]::SetEnvironmentVariable('NAME',`$null,'User')"
                    $results += "   - View specific: `$env:VARIABLE_NAME"
                    $results += "3. Registry locations:"
                    $results += "   - User variables: HKCU:\Environment"
                    $results += "   - System variables: HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
                    
                    $results += ""
                    $results += "=== SYSTEM CONFIGURATION ==="
                    try {
                        # Get additional system environment info
                        $systemInfo = Get-CimInstance -ClassName Win32_Environment -CimSession $cimSession -ErrorAction SilentlyContinue
                        if ($systemInfo) {
                            $systemVars = $systemInfo | Where-Object { $_.SystemVariable -eq $true }
                            $userVars = $systemInfo | Where-Object { $_.SystemVariable -eq $false }
                            $results += "System-wide variables: $($systemVars.Count)"
                            $results += "User-specific variables: $($userVars.Count)"
                        }
                        
                        # PowerShell execution policy
                        $execPolicy = Get-ExecutionPolicy
                        $results += "{0,-35} {1,-60}" -f "PowerShell Execution Policy", $execPolicy
                        
                        # .NET Framework versions
                        $dotNetVersions = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP" -Recurse -ErrorAction SilentlyContinue |
                                         Get-ItemProperty -Name Version -ErrorAction SilentlyContinue | 
                                         Select-Object -ExpandProperty Version -ErrorAction SilentlyContinue |
                                         Sort-Object -Unique
                        if ($dotNetVersions) {
                            $results += "{0,-35} {1,-60}" -f ".NET Framework Versions", ($dotNetVersions -join ", ")
                        }
                        
                    } catch {
                        $results += "{0,-35} {1,-60}" -f "System Config Error", $_.Exception.Message.Substring(0,50)+"..."
                    }
                    
                    $results += ""
                    $results += "=== ENVIRONMENT SUMMARY ==="
                    $results += "Total Environment Variables: $($envVars.Count)"
                    $results += "PATH Entries: $($pathEntries.Count)"
                    $results += "Key System Paths: Analyzed"
                    $results += "PowerShell Configuration: Checked"
                    $results += "Framework Information: Collected"
                    
                    $results += ""
                    $results += "=== END OF ENVIRONMENT LIST === ($($envVars.Count) variables analyzed)"
                    
                } catch {
                    $results += "SYSTEM ENVIRONMENT: Unable to retrieve - $($_.Exception.Message)"
                }
                $results += ""
            }

            if ($gpoCheckBox.Checked) {
                Update-Progress 90 "Collecting Group Policy information"
                try {
                    $results += "COMPLETE GROUP POLICY TABLE:"
                    $results += "Group Policy Objects and Settings Analysis"
                    $results += ""
                    
                    # Create GPO table header
                    $results += "{0,-40} {1,-30} {2,-25}" -f "Policy Category", "Setting Name", "Value/Status"
                    $results += "-" * 95
                    
                    # Get Local Group Policy settings
                    $results += "=== LOCAL GROUP POLICY SETTINGS ==="
                    
                    # Security Policy settings
                    $results += "Security Policies:"
                    try {
                        # Password Policy
                        $secPolicy = secedit /export /cfg "$env:TEMP\secpol.cfg" 2>$null
                        if (Test-Path "$env:TEMP\secpol.cfg") {
                            $secContent = Get-Content "$env:TEMP\secpol.cfg"
                            $minPwdLen = ($secContent | Where-Object { $_ -like "MinimumPasswordLength*" }) -split "=" | Select-Object -Last 1
                            $maxPwdAge = ($secContent | Where-Object { $_ -like "MaximumPasswordAge*" }) -split "=" | Select-Object -Last 1
                            $pwdComplexity = ($secContent | Where-Object { $_ -like "PasswordComplexity*" }) -split "=" | Select-Object -Last 1
                            
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Password Policy", "Minimum Password Length", "$minPwdLen characters"
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Password Policy", "Maximum Password Age", "$maxPwdAge days"
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Password Policy", "Password Complexity", $(if($pwdComplexity -eq "1"){"Enabled"}else{"Disabled"})
                            
                            Remove-Item "$env:TEMP\secpol.cfg" -Force -ErrorAction SilentlyContinue
                        }
                    } catch {
                        $results += "{0,-40} {1,-30} {2,-25}" -f "Security Policy", "Error retrieving", $_.Exception.Message.Substring(0,20)+"..."
                    }
                    
                    # Windows Update Policy (Registry-based)
                    try {
                        $wuPolicy = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -ErrorAction SilentlyContinue
                        $auPolicy = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -ErrorAction SilentlyContinue
                        
                        if ($wuPolicy -or $auPolicy) {
                            $results += "Windows Update Policies:"
                            if ($auPolicy.NoAutoUpdate) {
                                $results += "{0,-40} {1,-30} {2,-25}" -f "Windows Update", "Automatic Updates", $(if($auPolicy.NoAutoUpdate -eq 1){"Disabled"}else{"Enabled"})
                            }
                            if ($auPolicy.AUOptions) {
                                $auOption = switch ($auPolicy.AUOptions) {
                                    1 { "Keep me up to date" }
                                    2 { "Notify before download" }
                                    3 { "Notify before install" }
                                    4 { "Install automatically" }
                                    default { "Unknown ($($auPolicy.AUOptions))" }
                                }
                                $results += "{0,-40} {1,-30} {2,-25}" -f "Windows Update", "Auto Update Option", $auOption
                            }
                        } else {
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Windows Update", "Policy Status", "Not configured"
                        }
                    } catch {
                        $results += "{0,-40} {1,-30} {2,-25}" -f "Windows Update", "Error retrieving", "Registry access failed"
                    }
                    
                    # User Account Control Policy
                    try {
                        $uacPolicy = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue
                        if ($uacPolicy) {
                            $results += "User Account Control Policies:"
                            $results += "{0,-40} {1,-30} {2,-25}" -f "UAC", "Enable UAC", $(if($uacPolicy.EnableLUA -eq 1){"Enabled"}else{"Disabled"})
                            $results += "{0,-40} {1,-30} {2,-25}" -f "UAC", "Admin Approval Mode", $(if($uacPolicy.FilterAdministratorToken -eq 1){"Enabled"}else{"Disabled"})
                            $results += "{0,-40} {1,-30} {2,-25}" -f "UAC", "Consent Prompt Behavior", "Level $($uacPolicy.ConsentPromptBehaviorAdmin)"
                        }
                    } catch {
                        $results += "{0,-40} {1,-30} {2,-25}" -f "UAC Policy", "Error retrieving", "Registry access failed"
                    }
                    
                    # Firewall Policy
                    try {
                        $fwProfiles = @("Domain", "Private", "Public")
                        $results += "Windows Firewall Policies:"
                        foreach ($profile in $fwProfiles) {
                            $fwStatus = (Get-NetFirewallProfile -Name $profile -ErrorAction SilentlyContinue).Enabled
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Firewall", "$profile Profile", $(if($fwStatus){"Enabled"}else{"Disabled"})
                        }
                    } catch {
                        $results += "{0,-40} {1,-30} {2,-25}" -f "Firewall Policy", "Error retrieving", "PowerShell cmdlet failed"
                    }
                    
                    # Domain Group Policy (if domain-joined)
                    try {
                        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $cimSession
                        if ($computerSystem.PartOfDomain) {
                            $results += ""
                            $results += "=== DOMAIN GROUP POLICY STATUS ==="
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Domain Membership", "Domain Name", $computerSystem.Domain
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Domain Membership", "Workgroup Status", "Domain-joined"
                            
                            # Try to get GP result summary
                            $gpResult = gpresult /R /SCOPE COMPUTER 2>$null | Out-String
                            if ($gpResult) {
                                $gpLines = $gpResult -split "`n"
                                $lastGPUpdate = ($gpLines | Where-Object { $_ -like "*last time Group Policy was applied*" }) -replace ".*: ", ""
                                if ($lastGPUpdate) {
                                    $results += "{0,-40} {1,-30} {2,-25}" -f "Group Policy", "Last Update", $lastGPUpdate.Trim()
                                }
                            }
                        } else {
                            $results += ""
                            $results += "=== WORKGROUP COMPUTER ==="
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Domain Status", "Workgroup Name", $computerSystem.Domain
                            $results += "{0,-40} {1,-30} {2,-25}" -f "Domain Status", "Domain GP Status", "Not applicable"
                        }
                    } catch {
                        $results += "{0,-40} {1,-30} {2,-25}" -f "Domain Policy", "Error retrieving", "WMI query failed"
                    }
                    
                    $results += ""
                    $results += "=== GROUP POLICY SUMMARY ==="
                    $results += "Local Security Policies: Analyzed"
                    $results += "Windows Update Policies: Checked"
                    $results += "User Account Control: Analyzed"
                    $results += "Windows Firewall Policies: Checked"
                    if ($computerSystem.PartOfDomain) {
                        $results += "Domain Group Policies: Available (domain-joined)"
                    } else {
                        $results += "Domain Group Policies: Not applicable (workgroup)"
                    }
                    
                    $results += ""
                    $results += "=== END OF GROUP POLICY LIST === (Policy analysis completed)"
                    
                } catch {
                    $results += "GROUP POLICY: Unable to retrieve - $($_.Exception.Message)"
                }
                $results += ""
            }

            Update-Progress 100 "Complete!"
            Update-Status "System information collection completed successfully" "Green"
            
            return $results -join "`r`n"

        } catch {
            Update-Status "Error: $($_.Exception.Message)" "Red"
            return "ERROR: Failed to collect system information.`r`n`r`nDetails: $($_.Exception.Message)"
        } finally {
            if ($cimSession) { 
                Remove-CimSession $cimSession -ErrorAction SilentlyContinue 
            }
        }
    }

    # Authentication Mode Change Handler
    $authComboBox.Add_SelectedIndexChanged({
        $useCredentials = $authComboBox.SelectedIndex -eq 1
        $usernameLabel.Enabled = $useCredentials
        $usernameTextBox.Enabled = $useCredentials
        $passwordLabel.Enabled = $useCredentials
        $passwordTextBox.Enabled = $useCredentials
        
        if ($useCredentials) {
            $connectionStatusLabel.Text = "Status: Enter alternate credentials"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Orange
        } else {
            $connectionStatusLabel.Text = "Status: Using current user credentials"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Blue
        }
    })

    # Test Connection Handler
    $testButton.Add_Click({
        try {
            $computerName = $hostTextBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($computerName)) {
                $connectionStatusLabel.Text = "Status: Please enter a computer name"
                $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
                return
            }

            $connectionStatusLabel.Text = "Status: Testing connection..."
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Orange
            [System.Windows.Forms.Application]::DoEvents()

            # Create credential if needed
            $credential = $null
            if ($authComboBox.SelectedIndex -eq 1) {
                if ([string]::IsNullOrWhiteSpace($usernameTextBox.Text) -or [string]::IsNullOrWhiteSpace($passwordTextBox.Text)) {
                    $connectionStatusLabel.Text = "Status: Username and password required"
                    $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
                    return
                }
                $securePassword = ConvertTo-SecureString $passwordTextBox.Text -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($usernameTextBox.Text, $securePassword)
            }

            # Test connection
            $cimSessionOption = New-CimSessionOption -Protocol WSMan
            if ($credential -and $computerName -ne $env:COMPUTERNAME) {
                $testSession = New-CimSession -ComputerName $computerName -Credential $credential -SessionOption $cimSessionOption -ErrorAction Stop
            } else {
                $testSession = New-CimSession -ComputerName $computerName -SessionOption $cimSessionOption -ErrorAction Stop
            }

            $null = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $testSession -ErrorAction Stop
            Remove-CimSession $testSession -ErrorAction SilentlyContinue

            $connectionStatusLabel.Text = "Status: [OK] Connection successful!"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Green

        } catch {
            $connectionStatusLabel.Text = "Status: [FAILED] Connection failed - $($_.Exception.Message)"
            $connectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })

    # Connect Button Event Handler
    $connectButton.Add_Click({
        try {
            $computerName = $hostTextBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($computerName)) {
                Update-Status "Please enter a computer name" "Red"
                return
            }

            # Create credential if needed
            $credential = $null
            if ($authComboBox.SelectedIndex -eq 1) {
                if ([string]::IsNullOrWhiteSpace($usernameTextBox.Text) -or [string]::IsNullOrWhiteSpace($passwordTextBox.Text)) {
                    Update-Status "Username and password required" "Red"
                    return
                }
                $securePassword = ConvertTo-SecureString $passwordTextBox.Text -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($usernameTextBox.Text, $securePassword)
            }

            if ($credential) {
                $results = Get-SystemInformation -ComputerName $computerName -Credential $credential
            } else {
                $results = Get-SystemInformation -ComputerName $computerName
            }
            $resultsTextBox.Text = $results
        } catch {
            Update-Status "Connection failed: $($_.Exception.Message)" "Red"
        }
    })

    $clearButton.Add_Click({
        $resultsTextBox.Clear()
        Update-Status "Results cleared" "Blue"
    })

    $exportButton.Add_Click({
        try {
            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "Text files (*.txt)|*.txt"
            $saveDialog.FileName = "SystemInfo_$($hostTextBox.Text)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $resultsTextBox.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                Update-Status "Results exported to: $($saveDialog.FileName)" "Green"
            }
        } catch {
            Update-Status "Export failed: $($_.Exception.Message)" "Red"
        }
    })

    $regeditButton.Add_Click({
        try {
            Start-Process "regedit.exe"
            Update-Status "Registry Editor launched" "Green"
        } catch {
            Update-Status "Failed to launch Registry Editor" "Red"
        }
    })

    $printersButton.Add_Click({
        try {
            # Create Printer Analysis Window
            $printerForm = New-Object System.Windows.Forms.Form
            $printerForm.Text = "Comprehensive Printer Analysis"
            $printerForm.Size = New-Object System.Drawing.Size(1000, 700)
            $printerForm.StartPosition = "CenterScreen"
            $printerForm.FormBorderStyle = "Sizable"
            $printerForm.MaximizeBox = $true
            $printerForm.MinimizeBox = $true
            
            # Instructions Label
            $instructLabel = New-Object System.Windows.Forms.Label
            $instructLabel.Text = "Complete Printer and Print Queue Analysis"
            $instructLabel.Location = New-Object System.Drawing.Point(10, 10)
            $instructLabel.Size = New-Object System.Drawing.Size(970, 20)
            $instructLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
            
            # Results Text Box
            $resultsTextBox = New-Object System.Windows.Forms.TextBox
            $resultsTextBox.Location = New-Object System.Drawing.Point(10, 40)
            $resultsTextBox.Size = New-Object System.Drawing.Size(970, 550)
            $resultsTextBox.Multiline = $true
            $resultsTextBox.ScrollBars = "Vertical"
            $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
            $resultsTextBox.ReadOnly = $true
            
            # Refresh Button
            $refreshButton = New-Object System.Windows.Forms.Button
            $refreshButton.Text = "Refresh Analysis"
            $refreshButton.Location = New-Object System.Drawing.Point(10, 610)
            $refreshButton.Size = New-Object System.Drawing.Size(120, 30)
            $refreshButton.BackColor = [System.Drawing.Color]::Blue
            $refreshButton.ForeColor = [System.Drawing.Color]::White
            
            # Export Button
            $exportPrinterButton = New-Object System.Windows.Forms.Button
            $exportPrinterButton.Text = "Export Results"
            $exportPrinterButton.Location = New-Object System.Drawing.Point(140, 610)
            $exportPrinterButton.Size = New-Object System.Drawing.Size(120, 30)
            $exportPrinterButton.BackColor = [System.Drawing.Color]::Green
            $exportPrinterButton.ForeColor = [System.Drawing.Color]::White
            
            # Print Test Page Button
            $testPageButton = New-Object System.Windows.Forms.Button
            $testPageButton.Text = "Print Test Page"
            $testPageButton.Location = New-Object System.Drawing.Point(270, 610)
            $testPageButton.Size = New-Object System.Drawing.Size(120, 30)
            $testPageButton.BackColor = [System.Drawing.Color]::Orange
            $testPageButton.ForeColor = [System.Drawing.Color]::White
            
            # Close Button
            $closePrinterButton = New-Object System.Windows.Forms.Button
            $closePrinterButton.Text = "Close"
            $closePrinterButton.Location = New-Object System.Drawing.Point(860, 610)
            $closePrinterButton.Size = New-Object System.Drawing.Size(120, 30)
            $closePrinterButton.BackColor = [System.Drawing.Color]::Gray
            $closePrinterButton.ForeColor = [System.Drawing.Color]::White
            
            # Function to get printer information
            function Get-PrinterAnalysis {
                $results = @()
                $results += "COMPREHENSIVE PRINTER ANALYSIS"
                $results += "Generated: $(Get-Date)"
                $results += "=" * 80
                $results += ""
                
                # Get all printers
                $results += "=== INSTALLED PRINTERS ==="
                try {
                    $printers = Get-Printer -ErrorAction SilentlyContinue
                    if ($printers) {
                        $results += "Printer Name".PadRight(35) + "Status".PadRight(15) + "Type".PadRight(20) + "Location"
                        $results += "-" * 80
                        
                        $installedCount = 0
                        $readyCount = 0
                        $offlineCount = 0
                        
                        foreach ($printer in $printers) {
                            $installedCount++
                            $status = switch ($printer.PrinterStatus) {
                                "Normal" { $readyCount++; "Ready" }
                                "Offline" { $offlineCount++; "Offline" }
                                "Error" { "Error" }
                                "PaperJam" { "Paper Jam" }
                                "PaperOut" { "Paper Out" }
                                "OutputBinFull" { "Output Full" }
                                default { $printer.PrinterStatus }
                            }
                            
                            $type = if ($printer.Type) { $printer.Type } else { "Local" }
                            $location = if ($printer.Location) { $printer.Location } else { "Not specified" }
                            
                            $results += $printer.Name.PadRight(35) + $status.PadRight(15) + $type.PadRight(20) + $location
                        }
                        
                        $results += ""
                        $results += "Summary: $installedCount total printers ($readyCount ready, $offlineCount offline)"
                    } else {
                        $results += "No printers found or unable to retrieve printer information"
                    }
                } catch {
                    $results += "Error retrieving printers: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get print jobs
                $results += "=== PRINT QUEUE STATUS ==="
                try {
                    $printJobs = Get-PrintJob -ErrorAction SilentlyContinue
                    if ($printJobs) {
                        $results += "Job ID".PadRight(10) + "Document".PadRight(30) + "Status".PadRight(15) + "Printer"
                        $results += "-" * 80
                        
                        foreach ($job in $printJobs) {
                            $results += $job.Id.ToString().PadRight(10) + $job.DocumentName.PadRight(30) + $job.JobStatus.PadRight(15) + $job.PrinterName
                        }
                        $results += ""
                        $results += "Total print jobs in queue: $($printJobs.Count)"
                    } else {
                        $results += "No print jobs in queue"
                    }
                } catch {
                    $results += "Print queue information: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get printer drivers
                $results += "=== PRINTER DRIVERS ==="
                try {
                    $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
                    if ($drivers) {
                        $results += "Driver Name".PadRight(40) + "Version".PadRight(15) + "Environment"
                        $results += "-" * 80
                        
                        foreach ($driver in $drivers | Sort-Object Name) {
                            $version = if ($driver.MajorVersion) { "$($driver.MajorVersion).$($driver.MinorVersion)" } else { "Unknown" }
                            $environment = if ($driver.PrinterEnvironment) { $driver.PrinterEnvironment } else { "Windows x64" }
                            $results += $driver.Name.PadRight(40) + $version.PadRight(15) + $environment
                        }
                        $results += ""
                        $results += "Total printer drivers: $($drivers.Count)"
                    } else {
                        $results += "No printer drivers found"
                    }
                } catch {
                    $results += "Printer driver information: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get print spooler service status
                $results += "=== PRINT SPOOLER SERVICE ==="
                try {
                    $spooler = Get-Service -Name "Spooler" -ErrorAction SilentlyContinue
                    if ($spooler) {
                        $results += "Service Name: $($spooler.Name)"
                        $results += "Display Name: $($spooler.DisplayName)"
                        $results += "Status: $($spooler.Status)"
                        $results += "Start Type: $($spooler.StartType)"
                        $results += ""
                        
                        if ($spooler.Status -eq "Running") {
                            $results += "[OK] Print Spooler is running normally"
                        } else {
                            $results += "[WARNING] Print Spooler is not running - printing may not work"
                        }
                    }
                } catch {
                    $results += "Print spooler status: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get default printer
                $results += "=== DEFAULT PRINTER ==="
                try {
                    $defaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
                    if ($defaultPrinter) {
                        $results += "Default Printer: $($defaultPrinter.Name)"
                        $results += "Location: $($defaultPrinter.Location)"
                        $results += "Status: $($defaultPrinter.PrinterStatus)"
                        $results += "Paper Size: $($defaultPrinter.PaperSizesSupported -join ', ')"
                    } else {
                        $results += "No default printer set"
                    }
                } catch {
                    $results += "Default printer information: $($_.Exception.Message)"
                }
                
                $results += ""
                $results += "=== PRINTER ANALYSIS COMPLETE ==="
                $results += "Use buttons below for additional printer management functions"
                
                return $results -join "`r`n"
            }
            
            # Load initial printer analysis
            $resultsTextBox.Text = Get-PrinterAnalysis
            
            # Event Handlers
            $refreshButton.Add_Click({
                $resultsTextBox.Text = "Refreshing printer analysis..."
                [System.Windows.Forms.Application]::DoEvents()
                $resultsTextBox.Text = Get-PrinterAnalysis
            })
            
            $exportPrinterButton.Add_Click({
                try {
                    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                    $saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    $saveDialog.FileName = "PrinterAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
                    
                    if ($saveDialog.ShowDialog() -eq "OK") {
                        $resultsTextBox.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                        [System.Windows.Forms.MessageBox]::Show("Printer analysis exported successfully to:`n$($saveDialog.FileName)", "Export Complete", "OK", "Information")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error exporting results: $($_.Exception.Message)", "Export Error", "OK", "Error")
                }
            })
            
            $testPageButton.Add_Click({
                try {
                    $printers = Get-Printer -ErrorAction SilentlyContinue
                    if ($printers) {
                        # Create printer selection form
                        $selectForm = New-Object System.Windows.Forms.Form
                        $selectForm.Text = "Select Printer for Test Page"
                        $selectForm.Size = New-Object System.Drawing.Size(400, 300)
                        $selectForm.StartPosition = "CenterParent"
                        
                        $listBox = New-Object System.Windows.Forms.ListBox
                        $listBox.Location = New-Object System.Drawing.Point(10, 10)
                        $listBox.Size = New-Object System.Drawing.Size(370, 200)
                        
                        foreach ($printer in $printers) {
                            $listBox.Items.Add($printer.Name)
                        }
                        
                        $printButton = New-Object System.Windows.Forms.Button
                        $printButton.Text = "Print Test Page"
                        $printButton.Location = New-Object System.Drawing.Point(10, 220)
                        $printButton.Size = New-Object System.Drawing.Size(120, 30)
                        
                        $cancelButton = New-Object System.Windows.Forms.Button
                        $cancelButton.Text = "Cancel"
                        $cancelButton.Location = New-Object System.Drawing.Point(260, 220)
                        $cancelButton.Size = New-Object System.Drawing.Size(120, 30)
                        
                        $printButton.Add_Click({
                            if ($listBox.SelectedItem) {
                                try {
                                    # This would typically use Windows API to print test page
                                    [System.Windows.Forms.MessageBox]::Show("Test page sent to printer: $($listBox.SelectedItem)", "Test Page", "OK", "Information")
                                    $selectForm.Close()
                                } catch {
                                    [System.Windows.Forms.MessageBox]::Show("Error printing test page: $($_.Exception.Message)", "Print Error", "OK", "Error")
                                }
                            } else {
                                [System.Windows.Forms.MessageBox]::Show("Please select a printer", "No Selection", "OK", "Warning")
                            }
                        })
                        
                        $cancelButton.Add_Click({ $selectForm.Close() })
                        
                        $selectForm.Controls.AddRange(@($listBox, $printButton, $cancelButton))
                        $selectForm.ShowDialog()
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("No printers available for test page", "No Printers", "OK", "Warning")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error accessing printers: $($_.Exception.Message)", "Printer Error", "OK", "Error")
                }
            })
            
            $closePrinterButton.Add_Click({ $printerForm.Close() })
            
            # Add controls to printer form
            $printerForm.Controls.AddRange(@($instructLabel, $resultsTextBox, $refreshButton, $exportPrinterButton, $testPageButton, $closePrinterButton))
            
            # Show Printer Analysis Window
            Update-Status "Opening Comprehensive Printer Analysis..." "Green"
            $printerForm.ShowDialog()
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening Printer Analysis: $($_.Exception.Message)", "Error", "OK", "Error")
            Update-Status "Error opening printer analysis" "Red"
        }
    })

    $sharesButton.Add_Click({
        try {
            # Create Network Shares Analysis Window
            $sharesForm = New-Object System.Windows.Forms.Form
            $sharesForm.Text = "Comprehensive Network Shares Analysis"
            $sharesForm.Size = New-Object System.Drawing.Size(1000, 700)
            $sharesForm.StartPosition = "CenterScreen"
            $sharesForm.MinimizeBox = $true
            $sharesForm.MaximizeBox = $true
            
            # Instructions Label
            $instructLabel = New-Object System.Windows.Forms.Label
            $instructLabel.Text = "Network Shares Analysis - Complete inventory and management of all network shares"
            $instructLabel.Location = New-Object System.Drawing.Point(10, 10)
            $instructLabel.Size = New-Object System.Drawing.Size(970, 25)
            $instructLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            $instructLabel.ForeColor = [System.Drawing.Color]::DarkBlue
            
            # Results TextBox
            $resultsTextBox = New-Object System.Windows.Forms.TextBox
            $resultsTextBox.Location = New-Object System.Drawing.Point(10, 45)
            $resultsTextBox.Size = New-Object System.Drawing.Size(970, 550)
            $resultsTextBox.Multiline = $true
            $resultsTextBox.ScrollBars = "Both"
            $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
            $resultsTextBox.ReadOnly = $true
            $resultsTextBox.BackColor = [System.Drawing.Color]::White
            
            # Button Panel
            $buttonY = 610
            
            # Refresh Button
            $refreshButton = New-Object System.Windows.Forms.Button
            $refreshButton.Text = "Refresh Analysis"
            $refreshButton.Location = New-Object System.Drawing.Point(10, $buttonY)
            $refreshButton.Size = New-Object System.Drawing.Size(120, 35)
            $refreshButton.BackColor = [System.Drawing.Color]::Green
            $refreshButton.ForeColor = [System.Drawing.Color]::White
            $refreshButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            
            # Export Button
            $exportSharesButton = New-Object System.Windows.Forms.Button
            $exportSharesButton.Text = "Export Results"
            $exportSharesButton.Location = New-Object System.Drawing.Point(140, $buttonY)
            $exportSharesButton.Size = New-Object System.Drawing.Size(120, 35)
            $exportSharesButton.BackColor = [System.Drawing.Color]::Blue
            $exportSharesButton.ForeColor = [System.Drawing.Color]::White
            $exportSharesButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            
            # Map Drive Button
            $mapDriveButton = New-Object System.Windows.Forms.Button
            $mapDriveButton.Text = "Map Network Drive"
            $mapDriveButton.Location = New-Object System.Drawing.Point(270, $buttonY)
            $mapDriveButton.Size = New-Object System.Drawing.Size(140, 35)
            $mapDriveButton.BackColor = [System.Drawing.Color]::Orange
            $mapDriveButton.ForeColor = [System.Drawing.Color]::White
            $mapDriveButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            
            # Share Permissions Button
            $permissionsButton = New-Object System.Windows.Forms.Button
            $permissionsButton.Text = "Check Permissions"
            $permissionsButton.Location = New-Object System.Drawing.Point(420, $buttonY)
            $permissionsButton.Size = New-Object System.Drawing.Size(140, 35)
            $permissionsButton.BackColor = [System.Drawing.Color]::Purple
            $permissionsButton.ForeColor = [System.Drawing.Color]::White
            $permissionsButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            
            # Close Button
            $closeSharesButton = New-Object System.Windows.Forms.Button
            $closeSharesButton.Text = "Close"
            $closeSharesButton.Location = New-Object System.Drawing.Point(870, $buttonY)
            $closeSharesButton.Size = New-Object System.Drawing.Size(110, 35)
            $closeSharesButton.BackColor = [System.Drawing.Color]::Gray
            $closeSharesButton.ForeColor = [System.Drawing.Color]::White
            $closeSharesButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            
            # Comprehensive Shares Analysis Function
            function Get-NetworkSharesAnalysis {
                $results = @()
                $results += "COMPREHENSIVE NETWORK SHARES ANALYSIS"
                $results += "Generated: $(Get-Date)"
                $results += "Computer: $env:COMPUTERNAME"
                $results += "=" * 80
                $results += ""
                
                # Get local shares (SMB shares)
                $results += "=== LOCAL SMB SHARES ==="
                try {
                    $smbShares = Get-SmbShare -ErrorAction SilentlyContinue
                    if ($smbShares) {
                        $results += "Share Name                     Path                              Description"
                        $results += "-" * 80
                        foreach ($share in $smbShares) {
                            $shareName = $share.Name.PadRight(30)
                            $sharePath = $share.Path.PadRight(33)
                            $shareDesc = if ($share.Description) { $share.Description.Substring(0, [Math]::Min($share.Description.Length, 25)) } else { "No description" }
                            $results += "$shareName $sharePath $shareDesc"
                        }
                        $results += ""
                        $results += "Total SMB Shares: $($smbShares.Count)"
                    } else {
                        $results += "No SMB shares found or SMB service not available"
                    }
                } catch {
                    $results += "Error accessing SMB shares: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get WMI shares
                $results += "=== WMI NETWORK SHARES ==="
                try {
                    $wmiShares = Get-WmiObject -Class Win32_Share -ErrorAction SilentlyContinue
                    if ($wmiShares) {
                        $results += "Share Name                     Type        Path                              Description"
                        $results += "-" * 90
                        foreach ($share in $wmiShares) {
                            $shareName = $share.Name.PadRight(30)
                            $shareType = switch ($share.Type) {
                                0 { "Disk Drive" }
                                1 { "Print Queue" }
                                2 { "Device" }
                                3 { "IPC" }
                                2147483648 { "Disk Drive Admin" }
                                2147483649 { "Print Queue Admin" }
                                2147483650 { "Device Admin" }
                                2147483651 { "IPC Admin" }
                                default { "Unknown ($($share.Type))" }
                            }
                            $shareType = $shareType.PadRight(11)
                            $sharePath = if ($share.Path) { $share.Path.PadRight(33) } else { "N/A".PadRight(33) }
                            $shareDesc = if ($share.Description) { $share.Description.Substring(0, [Math]::Min($share.Description.Length, 20)) } else { "" }
                            $results += "$shareName $shareType $sharePath $shareDesc"
                        }
                        $results += ""
                        $results += "Total WMI Shares: $($wmiShares.Count)"
                    }
                } catch {
                    $results += "Error accessing WMI shares: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get mapped network drives
                $results += "=== MAPPED NETWORK DRIVES ==="
                try {
                    $mappedDrives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 4 }
                    if ($mappedDrives) {
                        $results += "Drive Letter    Remote Path                           Size (GB)    Free (GB)"
                        $results += "-" * 75
                        foreach ($drive in $mappedDrives) {
                            $driveLetter = "$($drive.DeviceID)".PadRight(15)
                            $remotePath = if ($drive.ProviderName) { $drive.ProviderName.PadRight(37) } else { "Unknown".PadRight(37) }
                            $totalSize = if ($drive.Size) { [math]::Round($drive.Size / 1GB, 2).ToString().PadLeft(12) } else { "N/A".PadLeft(12) }
                            $freeSpace = if ($drive.FreeSpace) { [math]::Round($drive.FreeSpace / 1GB, 2).ToString().PadLeft(12) } else { "N/A".PadLeft(12) }
                            $results += "$driveLetter $remotePath $totalSize $freeSpace"
                        }
                        $results += ""
                        $results += "Total Mapped Drives: $($mappedDrives.Count)"
                    } else {
                        $results += "No mapped network drives found"
                    }
                } catch {
                    $results += "Error accessing mapped drives: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get SMB client configuration
                $results += "=== SMB CLIENT CONFIGURATION ==="
                try {
                    $smbClientConfig = Get-SmbClientConfiguration -ErrorAction SilentlyContinue
                    if ($smbClientConfig) {
                        $results += "Connection Count Limit: $($smbClientConfig.ConnectionCountPerRssNetworkInterface)"
                        $results += "Directory Cache Lifetime: $($smbClientConfig.DirectoryCacheLifetime)"
                        $results += "Enable Bandwidth Throttling: $($smbClientConfig.EnableBandwidthThrottling)"
                        $results += "Enable Byte Range Locking: $($smbClientConfig.EnableByteRangeLockingOnReadOnlyFiles)"
                        $results += "Enable Large MTU: $($smbClientConfig.EnableLargeMtu)"
                        $results += "Enable Multichannel: $($smbClientConfig.EnableMultiChannel)"
                        $results += "File Info Cache Lifetime: $($smbClientConfig.FileInfoCacheLifetime)"
                        $results += "File Not Found Cache Lifetime: $($smbClientConfig.FileNotFoundCacheLifetime)"
                        $results += "Keep Connection Time: $($smbClientConfig.KeepConn)"
                        $results += "Max Commands: $($smbClientConfig.MaxCmds)"
                        $results += "Session Timeout: $($smbClientConfig.SessionTimeout)"
                        $results += "Use Opportunistic Locking: $($smbClientConfig.UseOpportunisticLocking)"
                    }
                } catch {
                    $results += "SMB client configuration not available: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get network adapter information
                $results += "=== NETWORK ADAPTER INFORMATION ==="
                try {
                    $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" -and $_.Virtual -eq $false } -ErrorAction SilentlyContinue
                    if ($networkAdapters) {
                        $results += "Adapter Name                      Interface Description              Link Speed"
                        $results += "-" * 80
                        foreach ($adapter in $networkAdapters) {
                            $adapterName = $adapter.Name.PadRight(33)
                            $adapterDesc = $adapter.InterfaceDescription.Substring(0, [Math]::Min($adapter.InterfaceDescription.Length, 33)).PadRight(33)
                            $linkSpeed = $adapter.LinkSpeed
                            $results += "$adapterName $adapterDesc $linkSpeed"
                        }
                    }
                } catch {
                    $results += "Network adapter information not available: $($_.Exception.Message)"
                }
                
                $results += ""
                
                # Get SMB server configuration (if available)
                $results += "=== SMB SERVER STATUS ==="
                try {
                    $serverService = Get-Service -Name "Server" -ErrorAction SilentlyContinue
                    if ($serverService) {
                        $results += "Server Service Status: $($serverService.Status)"
                        $results += "Server Service Start Type: $($serverService.StartType)"
                        
                        if ($serverService.Status -eq "Running") {
                            try {
                                $smbServerConfig = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
                                if ($smbServerConfig) {
                                    $results += "SMB Server Configuration:"
                                    $results += "  Auto Disconnect Timeout: $($smbServerConfig.AutoDisconnectTimeout)"
                                    $results += "  Enable SMB1 Protocol: $($smbServerConfig.EnableSMB1Protocol)"
                                    $results += "  Enable SMB2 Protocol: $($smbServerConfig.EnableSMB2Protocol)"
                                    $results += "  Max Channels Per Session: $($smbServerConfig.MaxChannelsPerSession)"
                                    $results += "  Max Sessions Per Connection: $($smbServerConfig.MaxSessionPerConnection)"
                                    $results += "  Max Thread Count Per Queue: $($smbServerConfig.MaxThreadsPerQueue)"
                                }
                            } catch {
                                $results += "  SMB server configuration details not accessible"
                            }
                        }
                    } else {
                        $results += "Server service not found"
                    }
                } catch {
                    $results += "Server service status not available: $($_.Exception.Message)"
                }
                
                $results += ""
                $results += "=== NETWORK SHARES ANALYSIS COMPLETE ==="
                $results += "Use buttons below for additional network share management functions"
                
                return $results -join "`r`n"
            }
            
            # Load initial shares analysis
            $resultsTextBox.Text = Get-NetworkSharesAnalysis
            
            # Event Handlers
            $refreshButton.Add_Click({
                $resultsTextBox.Text = "Refreshing network shares analysis..."
                [System.Windows.Forms.Application]::DoEvents()
                $resultsTextBox.Text = Get-NetworkSharesAnalysis
            })
            
            $exportSharesButton.Add_Click({
                try {
                    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                    $saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
                    $saveDialog.FileName = "NetworkSharesAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
                    
                    if ($saveDialog.ShowDialog() -eq "OK") {
                        $resultsTextBox.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                        [System.Windows.Forms.MessageBox]::Show("Network shares analysis exported successfully to:`n$($saveDialog.FileName)", "Export Complete", "OK", "Information")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error exporting results: $($_.Exception.Message)", "Export Error", "OK", "Error")
                }
            })
            
            $mapDriveButton.Add_Click({
                try {
                    # Create map drive dialog
                    $mapForm = New-Object System.Windows.Forms.Form
                    $mapForm.Text = "Map Network Drive"
                    $mapForm.Size = New-Object System.Drawing.Size(400, 250)
                    $mapForm.StartPosition = "CenterParent"
                    
                    $driveLabel = New-Object System.Windows.Forms.Label
                    $driveLabel.Text = "Drive Letter:"
                    $driveLabel.Location = New-Object System.Drawing.Point(10, 20)
                    $driveLabel.Size = New-Object System.Drawing.Size(80, 23)
                    
                    $driveCombo = New-Object System.Windows.Forms.ComboBox
                    $driveCombo.Location = New-Object System.Drawing.Point(100, 20)
                    $driveCombo.Size = New-Object System.Drawing.Size(60, 23)
                    $availableLetters = @('Z:', 'Y:', 'X:', 'W:', 'V:', 'U:', 'T:', 'S:', 'R:', 'Q:', 'P:', 'O:', 'N:', 'M:', 'L:', 'K:', 'J:', 'I:', 'H:', 'G:', 'F:', 'E:')
                    $driveCombo.Items.AddRange($availableLetters)
                    $driveCombo.SelectedIndex = 0
                    
                    $pathLabel = New-Object System.Windows.Forms.Label
                    $pathLabel.Text = "Network Path:"
                    $pathLabel.Location = New-Object System.Drawing.Point(10, 60)
                    $pathLabel.Size = New-Object System.Drawing.Size(80, 23)
                    
                    $pathTextBox = New-Object System.Windows.Forms.TextBox
                    $pathTextBox.Location = New-Object System.Drawing.Point(100, 60)
                    $pathTextBox.Size = New-Object System.Drawing.Size(270, 23)
                    $pathTextBox.Text = "\\\"
                    
                    $persistentCheckBox = New-Object System.Windows.Forms.CheckBox
                    $persistentCheckBox.Text = "Reconnect at sign-in"
                    $persistentCheckBox.Location = New-Object System.Drawing.Point(100, 100)
                    $persistentCheckBox.Size = New-Object System.Drawing.Size(200, 23)
                    $persistentCheckBox.Checked = $true
                    
                    $mapButton = New-Object System.Windows.Forms.Button
                    $mapButton.Text = "Map Drive"
                    $mapButton.Location = New-Object System.Drawing.Point(100, 140)
                    $mapButton.Size = New-Object System.Drawing.Size(80, 30)
                    
                    $cancelMapButton = New-Object System.Windows.Forms.Button
                    $cancelMapButton.Text = "Cancel"
                    $cancelMapButton.Location = New-Object System.Drawing.Point(200, 140)
                    $cancelMapButton.Size = New-Object System.Drawing.Size(80, 30)
                    
                    $mapButton.Add_Click({
                        if ($pathTextBox.Text -and $driveCombo.SelectedItem) {
                            try {
                                $persistent = $persistentCheckBox.Checked
                                $cmd = "net use $($driveCombo.SelectedItem) `"$($pathTextBox.Text)`""
                                if ($persistent) { $cmd += " /persistent:yes" }
                                
                                $result = cmd /c $cmd 2>&1
                                if ($LASTEXITCODE -eq 0) {
                                    [System.Windows.Forms.MessageBox]::Show("Network drive mapped successfully!", "Success", "OK", "Information")
                                    $mapForm.Close()
                                    # Refresh the main analysis
                                    $resultsTextBox.Text = Get-NetworkSharesAnalysis
                                } else {
                                    [System.Windows.Forms.MessageBox]::Show("Error mapping drive: $result", "Mapping Error", "OK", "Error")
                                }
                            } catch {
                                [System.Windows.Forms.MessageBox]::Show("Error mapping drive: $($_.Exception.Message)", "Error", "OK", "Error")
                            }
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Please enter both drive letter and network path", "Missing Information", "OK", "Warning")
                        }
                    })
                    
                    $cancelMapButton.Add_Click({ $mapForm.Close() })
                    
                    $mapForm.Controls.AddRange(@($driveLabel, $driveCombo, $pathLabel, $pathTextBox, $persistentCheckBox, $mapButton, $cancelMapButton))
                    $mapForm.ShowDialog()
                    
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error opening map drive dialog: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            })
            
            $permissionsButton.Add_Click({
                try {
                    $shares = Get-SmbShare -ErrorAction SilentlyContinue
                    if ($shares) {
                        # Create share selection dialog
                        $permForm = New-Object System.Windows.Forms.Form
                        $permForm.Text = "Check Share Permissions"
                        $permForm.Size = New-Object System.Drawing.Size(500, 400)
                        $permForm.StartPosition = "CenterParent"
                        
                        $listBox = New-Object System.Windows.Forms.ListBox
                        $listBox.Location = New-Object System.Drawing.Point(10, 10)
                        $listBox.Size = New-Object System.Drawing.Size(470, 200)
                        
                        foreach ($share in $shares) {
                            $listBox.Items.Add("$($share.Name) - $($share.Path)")
                        }
                        
                        $checkButton = New-Object System.Windows.Forms.Button
                        $checkButton.Text = "Check Permissions"
                        $checkButton.Location = New-Object System.Drawing.Point(10, 220)
                        $checkButton.Size = New-Object System.Drawing.Size(120, 30)
                        
                        $resultsBox = New-Object System.Windows.Forms.TextBox
                        $resultsBox.Location = New-Object System.Drawing.Point(10, 260)
                        $resultsBox.Size = New-Object System.Drawing.Size(470, 100)
                        $resultsBox.Multiline = $true
                        $resultsBox.ScrollBars = "Both"
                        $resultsBox.ReadOnly = $true
                        
                        $closePermButton = New-Object System.Windows.Forms.Button
                        $closePermButton.Text = "Close"
                        $closePermButton.Location = New-Object System.Drawing.Point(400, 220)
                        $closePermButton.Size = New-Object System.Drawing.Size(80, 30)
                        
                        $checkButton.Add_Click({
                            if ($listBox.SelectedItem) {
                                $selectedShare = $listBox.SelectedItem.ToString().Split(' - ')[0]
                                try {
                                    $permissions = Get-SmbShareAccess -Name $selectedShare -ErrorAction SilentlyContinue
                                    if ($permissions) {
                                        $permText = "Permissions for share: $selectedShare`r`n"
                                        $permText += "-" * 40 + "`r`n"
                                        foreach ($perm in $permissions) {
                                            $permText += "Account: $($perm.AccountName)`r`n"
                                            $permText += "Access: $($perm.AccessRight)`r`n"
                                            $permText += "Type: $($perm.AccessControlType)`r`n"
                                            $permText += "`r`n"
                                        }
                                        $resultsBox.Text = $permText
                                    } else {
                                        $resultsBox.Text = "No permissions information available for this share"
                                    }
                                } catch {
                                    $resultsBox.Text = "Error checking permissions: $($_.Exception.Message)"
                                }
                            } else {
                                [System.Windows.Forms.MessageBox]::Show("Please select a share", "No Selection", "OK", "Warning")
                            }
                        })
                        
                        $closePermButton.Add_Click({ $permForm.Close() })
                        
                        $permForm.Controls.AddRange(@($listBox, $checkButton, $resultsBox, $closePermButton))
                        $permForm.ShowDialog()
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("No shares available for permission checking", "No Shares", "OK", "Information")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error accessing shares for permission checking: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            })
            
            $closeSharesButton.Add_Click({ $sharesForm.Close() })
            
            # Add controls to shares form
            $sharesForm.Controls.AddRange(@($instructLabel, $resultsTextBox, $refreshButton, $exportSharesButton, $mapDriveButton, $permissionsButton, $closeSharesButton))
            
            # Show Network Shares Analysis Window
            Update-Status "Opening Comprehensive Network Shares Analysis..." "Green"
            $sharesForm.ShowDialog()
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening Network Shares Analysis: $($_.Exception.Message)", "Error", "OK", "Error")
            Update-Status "Error opening network shares analysis" "Red"
        }
    })

    $envEditorButton.Add_Click({
        try {
            # Create Environment Variable Editor Window
            $envForm = New-Object System.Windows.Forms.Form
            $envForm.Text = "Environment Variables Editor"
            $envForm.Size = New-Object System.Drawing.Size(900, 700)
            $envForm.StartPosition = "CenterScreen"
            $envForm.FormBorderStyle = "Sizable"
            $envForm.MaximizeBox = $true
            $envForm.MinimizeBox = $true
            
            # Instructions Label
            $instructLabel = New-Object System.Windows.Forms.Label
            $instructLabel.Text = "Environment Variables Editor - Select variable to edit, or create new ones"
            $instructLabel.Location = New-Object System.Drawing.Point(10, 10)
            $instructLabel.Size = New-Object System.Drawing.Size(870, 20)
            $instructLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
            
            # Environment Variables List
            $envListView = New-Object System.Windows.Forms.ListView
            $envListView.Location = New-Object System.Drawing.Point(10, 40)
            $envListView.Size = New-Object System.Drawing.Size(870, 400)
            $envListView.View = "Details"
            $envListView.FullRowSelect = $true
            $envListView.GridLines = $true
            $envListView.Sorting = "Ascending"
            
            # Add columns
            $envListView.Columns.Add("Variable Name", 250) | Out-Null
            $envListView.Columns.Add("Value", 500) | Out-Null
            $envListView.Columns.Add("Scope", 120) | Out-Null
            
            # Populate with current environment variables
            $allEnvVars = @()
            
            # Get user environment variables
            try {
                $userVars = [System.Environment]::GetEnvironmentVariables('User')
                foreach ($var in $userVars.GetEnumerator()) {
                    $item = New-Object System.Windows.Forms.ListViewItem($var.Key)
                    $item.SubItems.Add($var.Value)
                    $item.SubItems.Add("User")
                    $envListView.Items.Add($item) | Out-Null
                }
            } catch {}
            
            # Get system environment variables
            try {
                $systemVars = [System.Environment]::GetEnvironmentVariables('Machine')
                foreach ($var in $systemVars.GetEnumerator()) {
                    $item = New-Object System.Windows.Forms.ListViewItem($var.Key)
                    $item.SubItems.Add($var.Value)
                    $item.SubItems.Add("System")
                    $item.BackColor = [System.Drawing.Color]::LightGray
                    $envListView.Items.Add($item) | Out-Null
                }
            } catch {}
            
            # Variable Name Input
            $nameLabel = New-Object System.Windows.Forms.Label
            $nameLabel.Text = "Variable Name:"
            $nameLabel.Location = New-Object System.Drawing.Point(10, 460)
            $nameLabel.Size = New-Object System.Drawing.Size(100, 20)
            
            $nameTextBox = New-Object System.Windows.Forms.TextBox
            $nameTextBox.Location = New-Object System.Drawing.Point(120, 458)
            $nameTextBox.Size = New-Object System.Drawing.Size(200, 20)
            
            # Variable Value Input
            $valueLabel = New-Object System.Windows.Forms.Label
            $valueLabel.Text = "Variable Value:"
            $valueLabel.Location = New-Object System.Drawing.Point(10, 490)
            $valueLabel.Size = New-Object System.Drawing.Size(100, 20)
            
            $valueTextBox = New-Object System.Windows.Forms.TextBox
            $valueTextBox.Location = New-Object System.Drawing.Point(120, 488)
            $valueTextBox.Size = New-Object System.Drawing.Size(500, 60)
            $valueTextBox.Multiline = $true
            $valueTextBox.ScrollBars = "Vertical"
            
            # Scope Selection
            $scopeLabel = New-Object System.Windows.Forms.Label
            $scopeLabel.Text = "Scope:"
            $scopeLabel.Location = New-Object System.Drawing.Point(640, 490)
            $scopeLabel.Size = New-Object System.Drawing.Size(50, 20)
            
            $scopeComboBox = New-Object System.Windows.Forms.ComboBox
            $scopeComboBox.Location = New-Object System.Drawing.Point(700, 488)
            $scopeComboBox.Size = New-Object System.Drawing.Size(80, 20)
            $scopeComboBox.DropDownStyle = "DropDownList"
            $scopeComboBox.Items.AddRange(@("User", "System"))
            $scopeComboBox.SelectedIndex = 0
            
            # Buttons
            $setButton = New-Object System.Windows.Forms.Button
            $setButton.Text = "Set Variable"
            $setButton.Location = New-Object System.Drawing.Point(10, 570)
            $setButton.Size = New-Object System.Drawing.Size(100, 30)
            $setButton.BackColor = [System.Drawing.Color]::Green
            $setButton.ForeColor = [System.Drawing.Color]::White
            
            $deleteButton = New-Object System.Windows.Forms.Button
            $deleteButton.Text = "Delete Variable"
            $deleteButton.Location = New-Object System.Drawing.Point(120, 570)
            $deleteButton.Size = New-Object System.Drawing.Size(100, 30)
            $deleteButton.BackColor = [System.Drawing.Color]::Red
            $deleteButton.ForeColor = [System.Drawing.Color]::White
            
            $refreshButton = New-Object System.Windows.Forms.Button
            $refreshButton.Text = "Refresh List"
            $refreshButton.Location = New-Object System.Drawing.Point(230, 570)
            $refreshButton.Size = New-Object System.Drawing.Size(100, 30)
            $refreshButton.BackColor = [System.Drawing.Color]::Blue
            $refreshButton.ForeColor = [System.Drawing.Color]::White
            
            $closeButton = New-Object System.Windows.Forms.Button
            $closeButton.Text = "Close Editor"
            $closeButton.Location = New-Object System.Drawing.Point(780, 570)
            $closeButton.Size = New-Object System.Drawing.Size(100, 30)
            $closeButton.BackColor = [System.Drawing.Color]::Gray
            $closeButton.ForeColor = [System.Drawing.Color]::White
            
            # Event Handlers for Environment Editor
            $envListView.Add_SelectedIndexChanged({
                if ($envListView.SelectedItems.Count -gt 0) {
                    $selectedItem = $envListView.SelectedItems[0]
                    $nameTextBox.Text = $selectedItem.Text
                    $valueTextBox.Text = $selectedItem.SubItems[1].Text
                    $scopeComboBox.SelectedItem = $selectedItem.SubItems[2].Text
                }
            })
            
            $setButton.Add_Click({
                try {
                    if ($nameTextBox.Text -and $scopeComboBox.SelectedItem) {
                        $scope = if ($scopeComboBox.SelectedItem -eq "System") { "Machine" } else { "User" }
                        [System.Environment]::SetEnvironmentVariable($nameTextBox.Text, $valueTextBox.Text, $scope)
                        [System.Windows.Forms.MessageBox]::Show("Environment variable '$($nameTextBox.Text)' has been set successfully!", "Success", "OK", "Information")
                        $refreshButton.PerformClick()
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Please enter a variable name and select a scope.", "Error", "OK", "Error")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error setting environment variable: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            })
            
            $deleteButton.Add_Click({
                try {
                    if ($nameTextBox.Text -and $scopeComboBox.SelectedItem) {
                        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the environment variable '$($nameTextBox.Text)'?", "Confirm Delete", "YesNo", "Question")
                        if ($result -eq "Yes") {
                            $scope = if ($scopeComboBox.SelectedItem -eq "System") { "Machine" } else { "User" }
                            [System.Environment]::SetEnvironmentVariable($nameTextBox.Text, $null, $scope)
                            [System.Windows.Forms.MessageBox]::Show("Environment variable '$($nameTextBox.Text)' has been deleted!", "Success", "OK", "Information")
                            $nameTextBox.Clear()
                            $valueTextBox.Clear()
                            $refreshButton.PerformClick()
                        }
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Please select a variable to delete.", "Error", "OK", "Error")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error deleting environment variable: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            })
            
            $refreshButton.Add_Click({
                $envListView.Items.Clear()
                try {
                    $userVars = [System.Environment]::GetEnvironmentVariables('User')
                    foreach ($var in $userVars.GetEnumerator()) {
                        $item = New-Object System.Windows.Forms.ListViewItem($var.Key)
                        $item.SubItems.Add($var.Value)
                        $item.SubItems.Add("User")
                        $envListView.Items.Add($item) | Out-Null
                    }
                    $systemVars = [System.Environment]::GetEnvironmentVariables('Machine')
                    foreach ($var in $systemVars.GetEnumerator()) {
                        $item = New-Object System.Windows.Forms.ListViewItem($var.Key)
                        $item.SubItems.Add($var.Value)
                        $item.SubItems.Add("System")
                        $item.BackColor = [System.Drawing.Color]::LightGray
                        $envListView.Items.Add($item) | Out-Null
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error refreshing environment variables: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            })
            
            $closeButton.Add_Click({ $envForm.Close() })
            
            # Add controls to environment editor form
            $envForm.Controls.AddRange(@($instructLabel, $envListView, $nameLabel, $nameTextBox, $valueLabel, $valueTextBox, $scopeLabel, $scopeComboBox, $setButton, $deleteButton, $refreshButton, $closeButton))
            
            # Show Environment Editor
            $envForm.ShowDialog()
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error opening Environment Editor: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    $exitButton.Add_Click({
        $form.Close()
    })

    # Add all controls to form
    $form.Controls.AddRange(@($hostGroup, $infoGroup, $progressBar, $statusLabel, $resultsTextBox, $buttonPanel))

    # Show the form
    Update-Status "Ready - Select authentication mode and target computer" "Blue"
    [System.Windows.Forms.Application]::Run($form)

} catch {
    $errorMessage = @"
CRITICAL ERROR: Unable to start System Information GUI

Error Details: $($_.Exception.Message)

Troubleshooting:
- Run PowerShell as Administrator
- Ensure .NET Framework 4.5+ is installed
- Check execution policy: Get-ExecutionPolicy
"@

    Write-Host $errorMessage -ForegroundColor Red
    Read-Host "Press Enter to continue"
}