# System Information GUI Tool

A PowerShell-based GUI application for remotely gathering system information from Windows computers.

## Features

- **Host Selection**: Connect to local or remote computers by name or IP address
- **Credential Support**: Optional username/password for remote connections
- **Multiple Information Types**:
  - Disk Information (drive sizes, free space, file systems)
  - CPU Information (processor details, cores, clock speed, load)
  - Memory Information (total RAM, available memory, memory modules)
  - Installed Software (programs, versions, vendors)
  - Operating System Information (OS version, architecture, boot time)

## Requirements

- Windows PowerShell 5.1 or later
- WinRM enabled on target computers (for remote connections)
- Appropriate permissions on target computers

## Usage

### Method 1: Direct PowerShell Execution
```powershell
.\SystemInfoGUI.ps1
```

### Method 2: Batch File (Easier)
Double-click `Launch_SystemInfoGUI.bat`

## GUI Components

1. **Host Selection**: 
   - Enter computer name or IP address
   - Optionally provide credentials for remote access

2. **Information Selection**: 
   - Check boxes to select what information to retrieve
   - All options except "Installed Software" are selected by default

3. **Progress Tracking**: 
   - Progress bar shows completion status
   - Status messages indicate current operation

4. **Results Display**: 
   - Formatted text output with all requested information
   - Scrollable text area for large results

5. **Action Buttons**:
   - **Connect & Get Info**: Starts the information gathering process
   - **Clear Results**: Clears the results display
   - **Export to File**: Saves results to a text file
   - **Exit**: Closes the application

## Remote Connection Setup

For remote connections, ensure:

1. **Enable WinRM on target computer**:
   ```powershell
   Enable-PSRemoting -Force
   ```

2. **Configure trusted hosts (if needed)**:
   ```powershell
   Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
   ```

3. **Firewall**: Ensure Windows Remote Management is allowed through firewall

## Security Notes

- Credentials are handled securely using PowerShell's credential system
- Password field uses system password character masking
- No credentials are stored or logged by the application

## Troubleshooting

### Common Issues:

1. **"Access Denied" errors**: 
   - Provide valid credentials with administrative privileges
   - Ensure WinRM is enabled on target computer

2. **Connection timeouts**: 
   - Verify network connectivity
   - Check Windows Firewall settings
   - Ensure target computer is powered on

3. **Slow software retrieval**: 
   - Installed Software queries can be slow on some systems
   - Consider unchecking this option if not needed

## Output Format

The tool generates a comprehensive report including:
- Report header with computer name and timestamp
- Organized sections for each selected information type
- Detailed specifications and current usage statistics
- Export capability to timestamped text files

## File Structure

```
discovery/
├── SystemInfoGUI.ps1          # Main GUI application
├── Launch_SystemInfoGUI.bat   # Batch launcher
└── README.md                  # This documentation
```

## Version Information

- Created: October 1, 2025
- Compatible with: Windows PowerShell 5.1+, PowerShell 7.x
- Framework: Windows Forms (.NET Framework)