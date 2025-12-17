# Win11Debloat - Standalone Edition

## Overview

`Win11Debloat-Standalone.ps1` is a single-file version of Win11Debloat that consolidates all functionality from the multi-file project into one easy-to-use PowerShell script.

### Key Features

✅ **Single File** - No external dependencies on .reg files, .txt files, or binary files  
✅ **Easy Configuration** - All settings at the top of the script  
✅ **All Features Preserved** - Every feature from the original project  
✅ **Embedded Data** - 54 registry files, app lists, and menus built-in  
✅ **Fully Compatible** - Works with Windows 10 and Windows 11  
✅ **45+ Parameters** - Complete command-line support  

## Quick Start

### Interactive Mode (Recommended for first-time users)

```powershell
.\Win11Debloat-Standalone.ps1
```

This will show you a menu where you can choose:
1. Default Mode - Apply recommended settings
2. Custom Mode - Pick and choose what you want
3. App Removal Mode - Just remove apps
4. Saved Settings - Reapply your previous choices

### Default Mode (Recommended settings automatically)

```powershell
.\Win11Debloat-Standalone.ps1 -RunDefaults
```

### Custom App Selection (GUI)

```powershell
.\Win11Debloat-Standalone.ps1 -RunAppConfigurator
```

### Command-Line Mode

```powershell
.\Win11Debloat-Standalone.ps1 -RemoveApps -DisableTelemetry -DisableCopilot
```

## Customization

The script has easy-to-edit configuration sections at the top:

### 1. App Configuration (starts around line 119)

Control which apps are removed by default:

```powershell
$global:AppConfig = @{
    DefaultRemove = @(
        'Microsoft.BingNews'           # Remove by default
        'Microsoft.MicrosoftSolitaireCollection'
        # ... more apps
    )
    
    OptionalRemove = @(
        # 'Microsoft.OneDrive'         # Uncomment to remove
        # 'Microsoft.Edge'             # Uncomment to remove
    )
}
```

**To customize:**
- To KEEP an app normally removed: Add `#` in front of it in `DefaultRemove`
- To REMOVE an app normally kept: Remove `#` in front of it in `OptionalRemove`

### 2. Policy Configuration (starts around line 268)

Control which registry tweaks to apply:

```powershell
$global:PolicyConfig = @{
    DisableTelemetry = $false      # Change to $true to enable by default
    DisableCopilot = $false        # Change to $true to enable by default
    ShowKnownFileExt = $false      # Change to $true to enable by default
    # ... more options
}
```

**To customize:**
- Change `$false` to `$true` to enable a tweak by default
- Change `$true` to `$false` to disable a tweak by default

## Available Parameters

### App Removal
- `-RemoveApps` - Remove default selection of bloatware
- `-RemoveAppsCustom` - Remove apps from custom selection
- `-RemoveGamingApps` - Remove Xbox and gaming apps
- `-RemoveCommApps` - Remove Mail, Calendar, and People
- `-RemoveDevApps` - Remove developer apps
- `-RemoveW11Outlook` - Remove new Outlook for Windows
- `-ForceRemoveEdge` - Forcefully remove Microsoft Edge (not recommended)

### Privacy & Telemetry
- `-DisableTelemetry` - Disable telemetry and diagnostic data
- `-DisableBing` - Disable Bing search in Windows Search
- `-DisableSuggestions` - Disable tips, tricks, and ads
- `-DisableLockscreenTips` - Disable lockscreen tips
- `-DisableCopilot` - Disable Windows Copilot (Win11 only)
- `-DisableRecall` - Disable Windows Recall (Win11 AI feature)

### File Explorer
- `-ShowHiddenFolders` - Show hidden files and folders
- `-ShowKnownFileExt` - Show file extensions
- `-HideHome` - Hide Home section (Win11)
- `-HideGallery` - Hide Gallery section (Win11)
- `-HideDupliDrive` - Hide duplicate removable drives
- `-Hide3dObjects` - Hide 3D Objects folder (Win10)
- `-HideMusic` - Hide Music folder (Win10)
- `-HideOnedrive` - Hide OneDrive folder (Win10)

### Taskbar
- `-TaskbarAlignLeft` - Align taskbar icons left (Win11)
- `-HideSearchTb` - Hide search icon
- `-ShowSearchIconTb` - Show search icon
- `-ShowSearchLabelTb` - Show search icon with label
- `-ShowSearchBoxTb` - Show search box
- `-HideTaskview` - Hide taskview button
- `-DisableWidgets` - Disable widgets service
- `-HideChat` - Hide chat icon

### Context Menu
- `-RevertContextMenu` - Restore Windows 10 context menu (Win11)
- `-HideIncludeInLibrary` - Hide "Include in library"
- `-HideGiveAccessTo` - Hide "Give access to"
- `-HideShare` - Hide "Share"

### Other
- `-DisableDVR` - Disable Xbox game recording
- `-ClearStart` - Clear start menu for current user (Win11)
- `-ClearStartAllUsers` - Clear start menu for all users (Win11)

### Execution Modes
- `-RunDefaults` - Run with default settings
- `-RunAppConfigurator` - Open app selection GUI
- `-Sysprep` - Apply to default user profile
- `-Silent` - Run without prompts

## Examples

### Remove bloatware and disable telemetry
```powershell
.\Win11Debloat-Standalone.ps1 -RemoveApps -DisableTelemetry
```

### Full privacy-focused setup
```powershell
.\Win11Debloat-Standalone.ps1 -RemoveApps -DisableTelemetry -DisableBing -DisableCopilot -DisableSuggestions
```

### Just tweak File Explorer
```powershell
.\Win11Debloat-Standalone.ps1 -ShowHiddenFolders -ShowKnownFileExt -HideHome -HideGallery
```

### Customize for gaming PC
```powershell
.\Win11Debloat-Standalone.ps1 -RemoveApps -DisableTelemetry -DisableDVR:$false
```
(Note: `-DisableDVR:$false` keeps DVR enabled for game recording)

### Sysprep mode (for Windows image preparation)
```powershell
.\Win11Debloat-Standalone.ps1 -Sysprep -RemoveApps -DisableTelemetry
```

## Comparison with Original

| Feature | Original Multi-File | Standalone |
|---------|-------------------|------------|
| External Files Required | Yes (54 files) | No |
| Easy Configuration | Moderate | Very Easy |
| All Features | ✅ | ✅ |
| Command-Line Support | ✅ | ✅ |
| Distribution | Multiple files | Single file |
| Customization | Edit multiple files | Edit one section |

## Technical Details

- **Lines of Code:** ~2,700
- **Embedded Registry Files:** 54 (27 main + 27 Sysprep)
- **Apps Database:** 150+ apps categorized
- **Configuration Options:** 20+ registry tweaks
- **Parameters:** 45+ command-line switches
- **Size:** ~140 KB

## Security

✅ No hardcoded credentials  
✅ No code injection vulnerabilities  
✅ Proper input validation  
✅ Temp file cleanup  
✅ Embedded, validated registry content  
✅ Requires Administrator privileges  
✅ Language mode security check  

## Compatibility

- **Windows 11:** Fully supported (all builds)
- **Windows 10:** Fully supported
- **PowerShell:** 5.1 or higher
- **Architecture:** 32-bit and 64-bit

## Support

For issues, questions, or contributions:
- Original Project: https://github.com/Raphire/Win11Debloat
- Report bugs in the main project's issue tracker

## License

This standalone script is part of the Win11Debloat project and follows the same license.

## Credits

- Original Win11Debloat project by Raphire
- Standalone edition consolidation

---

**Note:** This standalone script is designed to complement, not replace, the original multi-file project. Choose whichever format works best for your needs!
