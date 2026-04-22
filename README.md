# Collect-Win32Logs

A PowerShell log collector for Intune Win32 (IME) app deployment troubleshooting. Produces a single ZIP file containing all relevant diagnostic artifacts.

## Quick Start

Run as **Administrator** on the target device:

```powershell
irm https://raw.githubusercontent.com/n1ghtmarauder/Collect-Win32Logs/main/Collect-Win32Logs.ps1 | iex
```

Or download and run locally:

```powershell
.\Collect-Win32Logs.ps1
```

The ZIP is saved to your Desktop and Explorer opens to it automatically.

## What It Collects

| Artifact | File(s) | Time Filtered |
|----------|---------|:---:|
| **IME Sidecar Logs** | AppWorkload, AppActionProcessor, IntuneManagementExtension, AgentExecutor | No (all rolled logs merged) |
| **Delivery Optimization** | Get-DeliveryOptimizationLog output | No |
| **Registry** | IntuneManagementExtension, EnterpriseDesktopAppManagement key exports | No |
| **Application Event Log** | XML export | Yes |
| **Store Event Log** | XML export | Yes |
| **AppxDeployment-Server Event Log** | XML export | Yes |
| **BITS Event Log** | XML export | Yes |

IME Sidecar logs are always collected in full because rolled logs are essential for proper iteration tracking and are often older than the time filter window.

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-DaysBack` | `7` | Days of history for event log exports |
| `-OutputRoot` | Desktop | Folder where the ZIP is created |
| `-MaxAppEvents` | `500` | Cap for Application event log (0 = no cap) |
| `-NoDOLog` | off | Skip Delivery Optimization log collection |
| `-NoZip` | off | Keep staging folder, skip ZIP creation |
| `-NoOpen` | off | Don't open Explorer after completion |

### Examples

```powershell
# Default - collect last 7 days of event logs, save ZIP to Desktop
.\Collect-Win32Logs.ps1

# Last 14 days, save to C:\Temp
.\Collect-Win32Logs.ps1 -DaysBack 14 -OutputRoot C:\Temp

# Skip Delivery Optimization log for faster collection
.\Collect-Win32Logs.ps1 -NoDOLog

# Uncapped Application event log
.\Collect-Win32Logs.ps1 -MaxAppEvents 0
```

## Output

```
COMPUTERNAME_Win32Logs_20260211-143022.zip
  ├── COMPUTERNAME_AppWorkload.log
  ├── COMPUTERNAME_AppActionProcessor.log
  ├── COMPUTERNAME_IntuneManagementExtension.log
  ├── COMPUTERNAME_AgentExecutor.log
  ├── COMPUTERNAME_Get-DeliveryOptimizationLog.txt
  ├── COMPUTERNAME_REG_SW_Microsoft_IntuneManagementExtension.txt
  ├── COMPUTERNAME_REG_SW_Microsoft_EnterpriseDesktopAppManagement.txt
  ├── COMPUTERNAME_Application.xml
  ├── COMPUTERNAME_Store.xml
  ├── COMPUTERNAME_AppxDeployment-Server.xml
  ├── COMPUTERNAME_BITS.xml
  └── COMPUTERNAME_Collector.log
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator elevation recommended (registry and event log access)

## License

MIT
