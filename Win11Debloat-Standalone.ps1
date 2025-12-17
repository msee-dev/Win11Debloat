#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Win11Debloat - Standalone Edition
    
.DESCRIPTION
    A consolidated, single-file version of Win11Debloat that removes bloatware, 
    disables telemetry, and declutters Windows 11/10.
    
    This standalone script contains ALL features from the original Win11Debloat project,
    with easy-to-configure sections at the top for customizing:
    - Which apps to remove (150+ apps organized by category)
    - Registry tweaks and policies (20+ options)
    - Feature toggles and execution modes
    
    All external dependencies (registry files, app lists, menus) are embedded directly
    in this script. No other files are needed!
    
.PARAMETER RunDefaults
    Run with default recommended settings without prompting
    
.PARAMETER RunAppConfigurator
    Open the app selection GUI to customize which apps to remove
    
.PARAMETER Sysprep
    Run in Sysprep mode to apply changes to the default user profile
    
.PARAMETER Silent
    Run without user prompts (for automated deployments)
    
.EXAMPLE
    .\Win11Debloat-Standalone.ps1
    Runs the script in interactive mode with menu
    
.EXAMPLE
    .\Win11Debloat-Standalone.ps1 -RunDefaults
    Applies the default recommended settings automatically
    
.EXAMPLE
    .\Win11Debloat-Standalone.ps1 -RemoveApps -DisableTelemetry
    Removes bloatware apps and disables telemetry
    
.EXAMPLE
    .\Win11Debloat-Standalone.ps1 -RunAppConfigurator
    Opens GUI to select specific apps to remove
    
.NOTES
    Author: Win11Debloat Project (Standalone Edition)
    Version: Standalone 1.0
    Requires: PowerShell running as Administrator
    
    To customize default behavior:
    1. Edit the $global:AppConfig section (starts around line 119)
    2. Edit the $global:PolicyConfig section (starts around line 268)
    3. Save and run the script
    
.LINK
    https://github.com/Raphire/Win11Debloat

#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$Silent,
    [switch]$Sysprep,
    [switch]$RunAppConfigurator,
    [switch]$RunDefaults, [switch]$RunWin11Defaults,
    [switch]$RemoveApps, 
    [switch]$RemoveAppsCustom,
    [switch]$RemoveGamingApps,
    [switch]$RemoveCommApps,
    [switch]$RemoveDevApps,
    [switch]$RemoveW11Outlook,
    [switch]$ForceRemoveEdge,
    [switch]$DisableDVR,
    [switch]$DisableTelemetry,
    [switch]$DisableBingSearches, [switch]$DisableBing,
    [switch]$DisableLockscrTips, [switch]$DisableLockscreenTips,
    [switch]$DisableWindowsSuggestions, [switch]$DisableSuggestions,
    [switch]$ShowHiddenFolders,
    [switch]$ShowKnownFileExt,
    [switch]$HideDupliDrive,
    [switch]$TaskbarAlignLeft,
    [switch]$HideSearchTb, [switch]$ShowSearchIconTb, [switch]$ShowSearchLabelTb, [switch]$ShowSearchBoxTb,
    [switch]$HideTaskview,
    [switch]$DisableCopilot,
    [switch]$DisableRecall,
    [switch]$DisableWidgets,
    [switch]$HideWidgets,
    [switch]$DisableChat,
    [switch]$HideChat,
    [switch]$ClearStart,
    [switch]$ClearStartAllUsers,
    [switch]$RevertContextMenu,
    [switch]$HideHome,
    [switch]$HideGallery,
    [switch]$DisableOnedrive, [switch]$HideOnedrive,
    [switch]$Disable3dObjects, [switch]$Hide3dObjects,
    [switch]$DisableMusic, [switch]$HideMusic,
    [switch]$DisableIncludeInLibrary, [switch]$HideIncludeInLibrary,
    [switch]$DisableGiveAccessTo, [switch]$HideGiveAccessTo,
    [switch]$DisableShare, [switch]$HideShare
)


#                                                                                                                #
#                                          TABLE OF CONTENTS                                                     #
#                                                                                                                #
##################################################################################################################
#
# 1. CONFIGURATION SECTION
#    - App Configuration
#    - Registry & Policy Configuration  
#    - Feature Toggles
#
# 2. EMBEDDED DATA
#    - Registry File Contents
#    - Menu Text Content
#    - Apps List Data
#
# 3. HELPER FUNCTIONS
#    - ShowAppSelectionForm
#    - ReadAppslistFromFile
#    - RemoveApps
#    - ForceRemoveEdge
#    - Strip-Progress
#    - RegImport
#    - RestartExplorer
#    - ReplaceStartMenu functions
#    - Utility functions
#
# 4. MAIN EXECUTION LOGIC
#    - Parameter handling
#    - Mode selection
#    - Feature execution
#



##################################################################################################################
#                                                                                                                #
#                                     CONFIGURATION SECTION                                                      #
#                                                                                                                #
##################################################################################################################

# ============================================
# CONFIGURATION SECTION - APPS TO REMOVE
# ============================================
# This section allows you to customize which apps are removed by default.
# The apps are categorized into three groups:
#   - DefaultRemove: Bloatware apps removed by default
#   - OptionalRemove: Apps NOT removed by default (gaming, dev, communication)
#   - NeverRemove: System-critical apps (kept for reference, commented out)
#
# To change default behavior:
#   - To KEEP an app that's normally removed: Comment it out with # in DefaultRemove
#   - To REMOVE an app that's normally kept: Uncomment it (remove #) in OptionalRemove

$global:AppConfig = @{
    # Apps that ARE removed by default
    DefaultRemove = @(
        # Microsoft bloatware
        'Clipchamp.Clipchamp'
        'Microsoft.3DBuilder'
        'Microsoft.549981C3F5F10'   # Cortana app
        'Microsoft.BingFinance'
        'Microsoft.BingFoodAndDrink'
        'Microsoft.BingHealthAndFitness'
        'Microsoft.BingNews'
        'Microsoft.BingSports'
        'Microsoft.BingTranslator'
        'Microsoft.BingTravel'
        'Microsoft.BingWeather'
        'Microsoft.Getstarted'   # Cannot be uninstalled in Windows 11
        'Microsoft.Messaging'
        'Microsoft.Microsoft3DViewer'
        'Microsoft.MicrosoftJournal'
        'Microsoft.MicrosoftOfficeHub'
        'Microsoft.MicrosoftPowerBIForWindows'
        'Microsoft.MicrosoftSolitaireCollection'
        'Microsoft.MicrosoftStickyNotes'
        'Microsoft.MixedReality.Portal'
        'Microsoft.NetworkSpeedTest'
        'Microsoft.News'
        'Microsoft.Office.OneNote'
        'Microsoft.Office.Sway'
        'Microsoft.OneConnect'
        'Microsoft.Print3D'
        'Microsoft.SkypeApp'
        'Microsoft.Todos'
        'Microsoft.WindowsAlarms'
        'Microsoft.WindowsFeedbackHub'
        'Microsoft.WindowsMaps'
        'Microsoft.WindowsSoundRecorder'
        'Microsoft.XboxApp'   # Old Xbox Console Companion App, no longer supported
        'Microsoft.ZuneVideo'
        'MicrosoftCorporationII.MicrosoftFamily'   # Family Safety App
        'MicrosoftCorporationII.QuickAssist'
        'MicrosoftTeams'   # Old MS Teams personal (MS Store)
        'MSTeams'   # New MS Teams app
        
        # Third-party bloatware
        'ACGMediaPlayer'
        'ActiproSoftwareLLC'
        'AdobeSystemsIncorporated.AdobePhotoshopExpress'
        'Amazon.com.Amazon'
        'AmazonVideo.PrimeVideo'
        'Asphalt8Airborne'
        'AutodeskSketchBook'
        'CaesarsSlotsFreeCasino'
        'COOKINGFEVER'
        'CyberLinkMediaSuiteEssentials'
        'DisneyMagicKingdoms'
        'Disney'
        'DrawboardPDF'
        'Duolingo-LearnLanguagesforFree'
        'EclipseManager'
        'Facebook'
        'FarmVille2CountryEscape'
        'fitbit'
        'Flipboard'
        'HiddenCity'
        'HULULLC.HULUPLUS'
        'iHeartRadio'
        'Instagram'
        'king.com.BubbleWitch3Saga'
        'king.com.CandyCrushSaga'
        'king.com.CandyCrushSodaSaga'
        'LinkedInforWindows'
        'MarchofEmpires'
        'Netflix'
        'NYTCrossword'
        'OneCalendar'
        'PandoraMediaInc'
        'PhototasticCollage'
        'PicsArt-PhotoStudio'
        'Plex'
        'PolarrPhotoEditorAcademicEdition'
        'Royal Revolt'
        'Shazam'
        'Sidia.LiveWallpaper'
        'SlingTV'
        'Spotify'
        'TikTok'
        'TuneInRadio'
        'Twitter'
        'Viber'
        'WinZipUniversal'
        'Wunderlist'
        'XING'
    )
    
    # Apps that are NOT removed by default - uncomment to remove
    OptionalRemove = @(
        # Uncomment any of these to remove them by default:
        # 'Microsoft.BingSearch'                   # Web Search from Microsoft Bing
        # 'Microsoft.Copilot'                      # New Windows Copilot app
        # 'Microsoft.Edge'                         # Edge browser (Can only be uninstalled in EEA)
        # 'Microsoft.GetHelp'                      # Required for some Win11 Troubleshooters
        # 'Microsoft.MSPaint'                      # Paint 3D
        # 'Microsoft.OneDrive'                     # OneDrive consumer
        # 'Microsoft.Paint'                        # Classic Paint
        # 'Microsoft.ScreenSketch'                 # Snipping Tool
        # 'Microsoft.Whiteboard'                   # Only preinstalled on touchscreen devices
        # 'Microsoft.Windows.Photos'
        # 'Microsoft.WindowsCalculator'
        # 'Microsoft.WindowsCamera'
        # 'Microsoft.WindowsStore'                 # Microsoft Store, WARNING: Cannot be reinstalled!
        # 'Microsoft.WindowsTerminal'              # New default terminal app in Windows 11
        # 'Microsoft.Xbox.TCUI'                    # Xbox UI framework, required for Store
        # 'Microsoft.XboxIdentityProvider'         # Xbox sign-in, required for some games
        # 'Microsoft.XboxSpeechToTextOverlay'      # WARNING: Cannot be reinstalled!
        # 'Microsoft.YourPhone'                    # Phone Link
        # 'Microsoft.ZuneMusic'                    # Modern Media Player
        # 'MicrosoftWindows.CrossDevice'           # Phone integration
    )
    
    # Apps for specific modes (gaming, dev, communication)
    GamingApps = @(
        'Microsoft.GamingApp'                    # Modern Xbox Gaming App
        'Microsoft.XboxGameOverlay'              # Game overlay
        'Microsoft.XboxGamingOverlay'            # Game overlay
    )
    
    CommunicationApps = @(
        'Microsoft.windowscommunicationsapps'    # Mail & Calendar
        'Microsoft.People'                       # Required for Mail & Calendar
    )
    
    DeveloperApps = @(
        'Microsoft.PowerAutomateDesktop'
        'Microsoft.RemoteDesktop'
        'Microsoft.Windows.DevHome'
    )
    
    NewOutlookApp = @(
        'Microsoft.OutlookForWindows'            # New Outlook for Windows
    )
}


# ============================================
# CONFIGURATION SECTION - POLICIES & TWEAKS
# ============================================
# Configure which registry tweaks and policies to apply by default.
# Set to $true to enable, $false to disable.

$global:PolicyConfig = @{
    # Telemetry & Privacy
    DisableTelemetry = $false           # Disable telemetry, diagnostic data, activity history
    DisableBing = $false                # Disable Bing search & Cortana in Windows search
    DisableLockscreenTips = $false      # Disable tips & tricks on lockscreen
    DisableSuggestions = $false         # Disable suggestions and ads across Windows
    DisableCopilot = $false             # Disable and remove Windows Copilot (Win11 22621+)
    DisableRecall = $false              # Disable Windows Recall snapshots (Win11 AI feature)
    
    # Xbox & Gaming
    DisableDVR = $false                 # Disable Xbox game/screen recording
    
    # File Explorer
    ShowHiddenFolders = $false          # Show hidden files, folders and drives
    ShowKnownFileExt = $false           # Show file extensions for known file types
    HideHome = $false                   # Hide Home section from Explorer (Win11)
    HideGallery = $false                # Hide Gallery section from Explorer (Win11)
    HideDupliDrive = $false             # Hide duplicate removable drives from sidebar
    HideOnedrive = $false               # Hide OneDrive folder from sidebar (Win10)
    Hide3dObjects = $false              # Hide 3D Objects folder (Win10)
    HideMusic = $false                  # Hide Music folder (Win10)
    
    # Taskbar
    TaskbarAlignLeft = $false           # Align taskbar icons to left (Win11)
    HideSearchTb = $false               # Hide search icon from taskbar (Win11)
    ShowSearchIconTb = $false           # Show search icon on taskbar (Win11)
    ShowSearchLabelTb = $false          # Show search icon with label (Win11)
    ShowSearchBoxTb = $false            # Show search box on taskbar (Win11)
    HideTaskview = $false               # Hide taskview button from taskbar (Win11)
    DisableWidgets = $false             # Disable widgets service and hide icon
    HideChat = $false                   # Hide chat (meet now) icon from taskbar
    
    # Context Menu
    RevertContextMenu = $false          # Restore old Windows 10 style context menu (Win11)
    HideIncludeInLibrary = $false       # Hide 'Include in library' context option
    HideGiveAccessTo = $false           # Hide 'Give access to' context option
    HideShare = $false                  # Hide 'Share' context option
    
    # Start Menu
    ClearStart = $false                 # Clear start menu for current user (Win11)
    ClearStartAllUsers = $false         # Clear start menu for all users (Win11)
}


# ============================================
# CONFIGURATION SECTION - FEATURE TOGGLES
# ============================================
# Control script behavior and execution modes

$global:FeatureConfig = @{
    # Execution modes
    DefaultMode = $false                # Run with default recommended settings
    SysprepMode = $false                # Apply changes to default user profile
    SilentMode = $false                 # Run without user prompts
    
    # App removal options
    RemoveApps = $false                 # Remove default bloatware apps
    RemoveAppsCustom = $false           # Remove custom app selection
    RemoveGamingApps = $false           # Remove gaming apps
    RemoveCommApps = $false             # Remove communication apps (Mail, Calendar)
    RemoveDevApps = $false              # Remove developer apps
    RemoveW11Outlook = $false           # Remove new Outlook for Windows
    ForceRemoveEdge = $false            # Forcefully remove Edge (not recommended)
}




##################################################################################################################
#                                                                                                                #
#                                    EMBEDDED DATA - REGISTRY FILES                                              #
#                                                                                                                #
##################################################################################################################

# This section contains all registry file contents embedded as here-strings
# These are used by the RegImport function to apply registry tweaks

$global:EmbeddedRegFiles = @{
    'Align_Taskbar_Left' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"TaskbarAl"=dword:00000000

'@

    'Disable_AI_Recall' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\WindowsAI]
"DisableAIDataAnalysis"=dword:00000001

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsAI]
"DisableAIDataAnalysis"=dword:00000001
'@

    'Disable_Bing_Cortana_In_Search' = @'
Windows Registry Editor Version 5.00

; Disable bing in search
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Explorer]
"DisableSearchBoxSuggestions"=dword:00000001

; Disable cortana in search
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search]
"AllowCortana"=dword:00000000
"CortanaConsent"=dword:00000000

'@

    'Disable_Chat_Taskbar' = @'
Windows Registry Editor Version 5.00

; Windows 11
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"TaskbarMn"=dword:00000000

; Windows 10
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]
"HideSCAMeetNow"=dword:00000001

'@

    'Disable_Copilot' = @'
Windows Registry Editor Version 5.00

; Disable Copilot button on taskbar
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowCopilotButton"=dword:00000000

; Disable Copilot service for current user
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\WindowsCopilot]
"TurnOffWindowsCopilot"=dword:00000001

; Disable Copilot service for all users
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot]
"TurnOffWindowsCopilot"=dword:00000001
'@

    'Disable_DVR' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\System\GameConfigStore]
"GameDVR_Enabled"=dword:00000000

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\GameDVR]
"AllowGameDVR"=dword:00000000

'@

    'Disable_Give_access_to_context_menu' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\shellex\CopyHookHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\shellex\PropertySheetHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Drive\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Drive\shellex\PropertySheetHandlers\Sharing]

[-HKEY_CLASSES_ROOT\LibraryFolder\background\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\UserLibraryFolder\shellex\ContextMenuHandlers\Sharing]

'@

    'Disable_Include_in_library_from_context_menu' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\Folder\ShellEx\ContextMenuHandlers\Library Location]

[-HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Folder\ShellEx\ContextMenuHandlers\Library Location]

'@

    'Disable_Lockscreen_Tips' = @'
Windows Registry Editor Version 5.00

; Get fun facts, tips and more from Windows and Cortana on your lock screen
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338387Enabled"=dword:00000000
"RotatingLockScreenOverlayEnabled"=dword:00000000
'@

    'Disable_Share_from_context_menu' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\ModernSharing]

'@

    'Disable_Show_More_Options_Context_Menu' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32]
@=""

'@

    'Disable_Telemetry' = @'
Windows Registry Editor Version 5.00

; Let Apps use Advertising ID for Relevant Ads in Windows 10
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo]
"Enabled"=dword:00000000

; Tailored experiences with diagnostic data for Current User
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Privacy]
"TailoredExperiencesWithDiagnosticDataEnabled"=dword:00000000

; Online Speech Recognition
[HKEY_CURRENT_USER\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy]
"HasAccepted"=dword:00000000

; Improve Inking & Typing Recognition
[HKEY_CURRENT_USER\Software\Microsoft\Input\TIPC]
"Enabled"=dword:00000000

; Inking & Typing Personalization
[HKEY_CURRENT_USER\Software\Microsoft\InputPersonalization]
"RestrictImplicitInkCollection"=dword:00000001
"RestrictImplicitTextCollection"=dword:00000001

[HKEY_CURRENT_USER\Software\Microsoft\InputPersonalization\TrainedDataStore]
"HarvestContacts"=dword:00000000

[HKEY_CURRENT_USER\Software\Microsoft\Personalization\Settings]
"AcceptedPrivacyPolicy"=dword:00000000

; Send only Required Diagnostic and Usage Data
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection]
"AllowTelemetry"=dword:00000000

; Disable Let Windows improve Start and search results by tracking app launches
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Start_TrackProgs"=dword:00000000

; Disable Activity History
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\System]
"PublishUserActivities"=dword:00000000

; Set Feedback Frequency to Never
[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Siuf\Rules]
"NumberOfSIUFInPeriod"=dword:00000000
"PeriodInNanoSeconds"=-

'@

    'Disable_Widgets_Taskbar' = @'
Windows Registry Editor Version 5.00

; Windows 11
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"TaskbarDa"=dword:00000000

; Windows 10
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Feeds]
"ShellFeedsTaskbarViewMode"=dword:00000002

; Disable widgets service
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PolicyManager\default\NewsAndInterests\AllowNewsAndInterests]
"value"=dword:00000000

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Dsh]
"AllowNewsAndInterests"=dword:00000000

'@

    'Disable_Windows_Suggestions' = @'
Windows Registry Editor Version 5.00

; Show me the Windows welcome experience after updates and occasionally when I sign in to highlight what's new and suggested
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-310093Enabled"=dword:00000000

; Occasionally show suggestions in Start
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338388Enabled"=dword:00000000
"SystemPaneSuggestionsEnabled"=dword:00000000

; Show recommendations for tips, shortcuts, new apps, and more in start
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Start_IrisRecommendations"=dword:00000000

; Get tips, tricks, and suggestions as you use Windows
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338389Enabled"=dword:00000000
"SoftLandingEnabled"=dword:00000000

; Show me suggested content in the Settings app
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338393Enabled"=dword:00000000
"SubscribedContent-353694Enabled"=dword:00000000
"SubscribedContent-353696Enabled"=dword:00000000

; Disable Show me notifications in the Settings app
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\SystemSettings\AccountNotifications]
"EnableAccountNotifications"=dword:00000000

; Suggest ways I can finish setting up my device to get the most out of Windows
[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\UserProfileEngagement]
"ScoobeSystemSettingEnabled"=dword:00000000

; Sync provider ads
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowSyncProviderNotifications"=dword:00000000

; Automatic Installation of Suggested Apps
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SilentInstalledAppsEnabled"=dword:00000000

; Disable "Suggested" app notifications (Ads for MS services)
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested]
"Enabled"=dword:00000000

; Disable Show me suggestions for using my mobile device with Windows (Phone Link suggestions)
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Mobility]
"OptedIn"=dword:00000000

; Disable Show account-related notifications
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Start_AccountNotifications"=dword:00000000

'@

    'Hide_3D_Objects_Folder' = @'
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}]

[-HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}]

'@

    'Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer' = @'
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders\{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}]

'@

    'Hide_Gallery_from_Explorer' = @'
Windows Registry Editor Version 5.00

; Hide Gallery on Navigation Pane for current user
[HKEY_CURRENT_USER\Software\Classes\CLSID\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}]
"System.IsPinnedToNameSpaceTree"=dword:00000000

; Add `Show Gallery` option to File Explorer folder options, with default set to disabled
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowGallery]
"CheckedValue"=dword:00000001
"DefaultValue"=dword:00000000
"HKeyRoot"=dword:80000001
"Id"=dword:0000000d
"RegPath"="Software\\Classes\\CLSID\\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}"
"Text"="Show Gallery"
"Type"="checkbox"
"UncheckedValue"=dword:00000000
"ValueName"="System.IsPinnedToNameSpaceTree"

'@

    'Hide_Home_from_Explorer' = @'
Windows Registry Editor Version 5.00

; Hide Home on Navigation Pane for current user
[HKEY_CURRENT_USER\Software\Classes\CLSID\{f874310e-b6b7-47dc-bc84-b9e6b38f5903}]
@="CLSID_MSGraphHomeFolder"
"System.IsPinnedToNameSpaceTree"=dword:00000000

; Add `Show Home` option to File Explorer folder options, with default set to disabled
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowHome]
"CheckedValue"=dword:00000001
"DefaultValue"=dword:00000000
"HKeyRoot"=dword:80000001
"Id"=dword:0000000d
"RegPath"="Software\\Classes\\CLSID\\{f874310e-b6b7-47dc-bc84-b9e6b38f5903}"
"Text"="Show Home"
"Type"="checkbox"
"UncheckedValue"=dword:00000000
"ValueName"="System.IsPinnedToNameSpaceTree"

'@

    'Hide_Music_Folder' = @'
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}]

[-HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}]

'@

    'Hide_Onedrive_Folder' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]
"System.IsPinnedToNameSpaceTree"=dword:0

[-HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]
"System.IsPinnedToNameSpaceTree"=dword:0

'@

    'Hide_Search_Taskbar' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Search]
"SearchboxTaskbarMode"=dword:00000000

'@

    'Hide_Taskview_Taskbar' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowTaskViewButton"=dword:00000000

'@

    'Show_Extensions_For_Known_File_Types' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"HideFileExt"=dword:00000000

'@

    'Show_Hidden_Folders' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Hidden"=dword:00000001

'@

    'Show_Search_Box' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Search]
"SearchboxTaskbarMode"=dword:00000002

'@

    'Show_Search_Icon_And_Label' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Search]
"SearchboxTaskbarMode"=dword:00000003

'@

    'Show_Search_Icon' = @'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Search]
"SearchboxTaskbarMode"=dword:00000001

'@

}


# Sysprep mode registry files (for default user profile)
$global:EmbeddedRegFilesSysprep = @{
    'Align_Taskbar_Left' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"TaskbarAl"=dword:00000000


'@

    'Disable_AI_Recall' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\Software\Policies\Microsoft\Windows\WindowsAI]
"DisableAIDataAnalysis"=dword:00000001

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsAI]
"DisableAIDataAnalysis"=dword:00000001
'@

    'Disable_Bing_Cortana_In_Search' = @'
Windows Registry Editor Version 5.00

; Disable bing in search
[hkey_users\default\Software\Policies\Microsoft\Windows\Explorer]
"DisableSearchBoxSuggestions"=dword:00000001

; Disable cortana in search
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Windows Search]
"AllowCortana"=dword:00000000
"CortanaConsent"=dword:00000000

'@

    'Disable_Chat_Taskbar' = @'
Windows Registry Editor Version 5.00

; Windows 11
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"TaskbarMn"=dword:00000000

; Windows 10
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]
"HideSCAMeetNow"=dword:00000001

'@

    'Disable_Copilot' = @'
Windows Registry Editor Version 5.00

; Disable Copilot button on taskbar
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowCopilotButton"=dword:00000000

; Disable Copilot service for current user
[hkey_users\default\Software\Policies\Microsoft\Windows\WindowsCopilot]
"TurnOffWindowsCopilot"=dword:00000001

; Disable Copilot service for all users
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot]
"TurnOffWindowsCopilot"=dword:00000001
'@

    'Disable_DVR' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\System\GameConfigStore]
"GameDVR_Enabled"=dword:00000000

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\GameDVR]
"AllowGameDVR"=dword:00000000

'@

    'Disable_Give_access_to_context_menu' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\shellex\CopyHookHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Directory\shellex\PropertySheetHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Drive\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\Drive\shellex\PropertySheetHandlers\Sharing]

[-HKEY_CLASSES_ROOT\LibraryFolder\background\shellex\ContextMenuHandlers\Sharing]

[-HKEY_CLASSES_ROOT\UserLibraryFolder\shellex\ContextMenuHandlers\Sharing]

'@

    'Disable_Include_in_library_from_context_menu' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\Folder\ShellEx\ContextMenuHandlers\Library Location]

[-HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Folder\ShellEx\ContextMenuHandlers\Library Location]

'@

    'Disable_Lockscreen_Tips' = @'
Windows Registry Editor Version 5.00

; Get fun facts, tips and more from Windows and Cortana on your lock screen
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338387Enabled"=dword:00000000
"RotatingLockScreenOverlayEnabled"=dword:00000000
'@

    'Disable_Share_from_context_menu' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\ModernSharing]

'@

    'Disable_Show_More_Options_Context_Menu' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"RestoreWin10ContextMenu"="reg add HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\\InprocServer32 /f /ve"

'@

    'Disable_Telemetry' = @'
Windows Registry Editor Version 5.00

; Let Apps use Advertising ID for Relevant Ads in Windows 10
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo]
"Enabled"=dword:00000000

; Tailored experiences with diagnostic data for Current User
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Privacy]
"TailoredExperiencesWithDiagnosticDataEnabled"=dword:00000000

; Online Speech Recognition
[hkey_users\default\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy]
"HasAccepted"=dword:00000000

; Improve Inking & Typing Recognition
[hkey_users\default\Software\Microsoft\Input\TIPC]
"Enabled"=dword:00000000

; Inking & Typing Personalization
[hkey_users\default\Software\Microsoft\InputPersonalization]
"RestrictImplicitInkCollection"=dword:00000001
"RestrictImplicitTextCollection"=dword:00000001

[hkey_users\default\Software\Microsoft\InputPersonalization\TrainedDataStore]
"HarvestContacts"=dword:00000000

[hkey_users\default\Software\Microsoft\Personalization\Settings]
"AcceptedPrivacyPolicy"=dword:00000000

; Send only Required Diagnostic and Usage Data
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection]
"AllowTelemetry"=dword:00000000

; Disable Let Windows improve Start and search results by tracking app launches
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Start_TrackProgs"=dword:00000000

; Disable Activity History
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\System]
"PublishUserActivities"=dword:00000000

; Set Feedback Frequency to Never
[hkey_users\default\SOFTWARE\Microsoft\Siuf\Rules]
"NumberOfSIUFInPeriod"=dword:00000000
"PeriodInNanoSeconds"=-

'@

    'Disable_Widgets_Taskbar' = @'
Windows Registry Editor Version 5.00

; Windows 11
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"TaskbarDa"=dword:00000000

; Windows 10
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Feeds]
"ShellFeedsTaskbarViewMode"=dword:00000002

; Disable widgets service
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PolicyManager\default\NewsAndInterests\AllowNewsAndInterests]
"value"=dword:00000000

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Dsh]
"AllowNewsAndInterests"=dword:00000000

'@

    'Disable_Windows_Suggestions' = @'
Windows Registry Editor Version 5.00

; Show me the Windows welcome experience after updates and occasionally when I sign in to highlight what's new and suggested
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-310093Enabled"=dword:00000000

; Occasionally show suggestions in Start
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338388Enabled"=dword:00000000
"SystemPaneSuggestionsEnabled"=dword:00000000

; Show recommendations for tips, shortcuts, new apps, and more in start
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Start_IrisRecommendations"=dword:00000000

; Get tips, tricks, and suggestions as you use Windows
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338389Enabled"=dword:00000000
"SoftLandingEnabled"=dword:00000000

; Show me suggested content in the Settings app
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SubscribedContent-338393Enabled"=dword:00000000
"SubscribedContent-353694Enabled"=dword:00000000
"SubscribedContent-353696Enabled"=dword:00000000

; Disable Show me notifications in the Settings app
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\SystemSettings\AccountNotifications]
"EnableAccountNotifications"=dword:00000000

; Suggest ways I can finish setting up my device to get the most out of Windows
[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\UserProfileEngagement]
"ScoobeSystemSettingEnabled"=dword:00000000

; Sync provider ads
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowSyncProviderNotifications"=dword:00000000

; Automatic Installation of Suggested Apps
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager]
"SilentInstalledAppsEnabled"=dword:00000000

; Disable "Suggested" app notifications (Ads for MS services)
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested]
"Enabled"=dword:00000000

; Disable Show me suggestions for using my mobile device with Windows (Phone Link suggestions)
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Mobility]
"OptedIn"=dword:00000000

; Disable Show account-related notifications
[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Start_AccountNotifications"=dword:00000000

'@

    'Hide_3D_Objects_Folder' = @'
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}]

[-HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}]

'@

    'Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer' = @'
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\DelegateFolders\{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}]

'@

    'Hide_Gallery_from_Explorer' = @'
Windows Registry Editor Version 5.00

; Hide Gallery on Navigation Pane
[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"HideGalleryExplorer"="reg add HKCU\\Software\\Classes\\CLSID\\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c} /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 0 /f"

; Add `Show Gallery` option to File Explorer folder options, with default set to disabled
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowGallery]
"CheckedValue"=dword:00000001
"DefaultValue"=dword:00000000
"HKeyRoot"=dword:80000001
"Id"=dword:0000000d
"RegPath"="Software\\Classes\\CLSID\\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}"
"Text"="Show Gallery"
"Type"="checkbox"
"UncheckedValue"=dword:00000000
"ValueName"="System.IsPinnedToNameSpaceTree"

'@

    'Hide_Home_from_Explorer' = @'
Windows Registry Editor Version 5.00

; Hide Home on Navigation Pane
[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"HideHomeExplorer1"="reg add HKCU\\Software\\Classes\\CLSID\\{f874310e-b6b7-47dc-bc84-b9e6b38f5903} /d CLSID_MSGraphHomeFolder /f /ve"
"HideHomeExplorer2"="reg add HKCU\\Software\\Classes\\CLSID\\{f874310e-b6b7-47dc-bc84-b9e6b38f5903} /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 0 /f"

; Add `Show Home` option to File Explorer folder options, with default set to disabled
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowHome]
"CheckedValue"=dword:00000001
"DefaultValue"=dword:00000000
"HKeyRoot"=dword:80000001
"Id"=dword:0000000d
"RegPath"="Software\\Classes\\CLSID\\{f874310e-b6b7-47dc-bc84-b9e6b38f5903}"
"Text"="Show Home"
"Type"="checkbox"
"UncheckedValue"=dword:00000000
"ValueName"="System.IsPinnedToNameSpaceTree"

'@

    'Hide_Music_Folder' = @'
Windows Registry Editor Version 5.00

[-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}]

[-HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}]

'@

    'Hide_Onedrive_Folder' = @'
Windows Registry Editor Version 5.00

[-HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]
"System.IsPinnedToNameSpaceTree"=dword:0

[-HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}]
"System.IsPinnedToNameSpaceTree"=dword:0

'@

    'Hide_Search_Taskbar' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"SetTaskbarSearchbox"="reg add HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search /v SearchboxTaskbarMode /t REG_DWORD /d 0 /f"

'@

    'Hide_Taskview_Taskbar' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowTaskViewButton"=dword:00000000

'@

    'Show_Extensions_For_Known_File_Types' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"HideFileExt"=dword:00000000

'@

    'Show_Hidden_Folders' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"Hidden"=dword:00000001

'@

    'Show_Search_Box' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"SetTaskbarSearchbox"="reg add HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search /v SearchboxTaskbarMode /t REG_DWORD /d 2 /f"

'@

    'Show_Search_Icon_And_Label' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"SetTaskbarSearchbox"="reg add HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search /v SearchboxTaskbarMode /t REG_DWORD /d 3 /f"

'@

    'Show_Search_Icon' = @'
Windows Registry Editor Version 5.00

[hkey_users\default\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce]
"SetTaskbarSearchbox"="reg add HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search /v SearchboxTaskbarMode /t REG_DWORD /d 1 /f"

'@

}

##################################################################################################################
#                                                                                                                #
#                                    EMBEDDED DATA - MENU TEXT CONTENT                                           #
#                                                                                                                #
##################################################################################################################

# Default Settings menu text
$global:DefaultSettingsText = @'
-------------------------------------------------------------------------------------------
 Win11Debloat Script - Default Mode
-------------------------------------------------------------------------------------------
Win11Debloat will make the following changes:
- Remove the default selection of apps, the list can be found in the 'Appslist.txt' file.
- Disable telemetry, diagnostic data, app-launch tracking & targeted ads.
- Disable & remove Bing search & Cortana in Windows search.
- Disable tips & tricks on the lockscreen. (This may change your lockscreen wallpaper)
- Disable tips, tricks, suggestions and ads in start, settings, notifications and more.
- Disable Windows Copilot. (Windows 11 build 22621+)
- Show file extensions for known file types.
- Disable the widget service & hide the icon from the taskbar. 
- Hide the Chat (meet now) icon from the taskbar.
- Hide the 3D objects folder in Windows Explorer. (Windows 10 only)

'@

# Information menu text
$global:InfoText = @'
-------------------------------------------------------------------------------------------
 Win11Debloat Script - Information
-------------------------------------------------------------------------------------------
Win11Debloat is a simple, easy to use and lightweight PowerShell script that can remove
pre-installed Windows bloatware apps, disable telemetry and declutter the experience by
disabling or removing intrusive interface elements, ads and more. No need to go through
all the settings yourself, or remove apps one by one.

-------------------------------------------------------------------------------------------
 All Features
-------------------------------------------------------------------------------------------
App Removal
- Remove a wide variety of bloatware apps.
- Remove all pinned apps from start for the current user, or for all new & existing users. (Windows 11 only)

Telemetry, Tracking & Suggested Content
- Disable telemetry, diagnostic data, activity history, app-launch tracking & targeted ads.
- Disable tips, tricks, suggestions & ads across Windows.

Bing, Copilot & More
- Disable & remove Bing web search & Cortana from Windows search.
- Disable Windows Copilot. (Windows 11 only)
- Disable Windows Recall snapshots. (Windows 11 only)

File Explorer
- Show hidden files, folders & drives.
- Show file extensions for known file types.
- Hide the gallery section from the File Explorer sidepanel. (Windows 11 only)
- Hide the 3D objects, music or onedrive folder from the File Explorer sidepanel. (Windows 10 only)
- Hide duplicate removable drive entries from the File Explorer sidepanel.

Taskbar
- Align taskbar icons to the left. (Windows 11 only)
- Hide or change the search icon/box on the taskbar. (Windows 11 only)
- Hide the taskview button from the taskbar. (Windows 11 only)
- Disable the widgets service & hide icon from the taskbar.
- Hide the chat (meet now) icon from the taskbar.

Context menu
- Restore the old Windows 10 style context menu. (Windows 11 only)
- Hide the 'Include in library', 'Give access to' & 'Share' options from the context menu. (Windows 10 only)

Other
- Disable Xbox game/screen recording (Also stops gaming overlay popups)

Advanced Features
- Sysprep mode to apply changes to the Windows Default user profile.

-------------------------------------------------------------------------------------------
 Default mode
-------------------------------------------------------------------------------------------
The default mode applies the changes that are recommended for most users, this includes:
- Remove the default selection of apps, the list can be found in the 'Appslist.txt' file.
- Disable telemetry, diagnostic data, activity history, app-launch tracking & targeted ads.
- Disable & remove Bing web search & Cortana from Windows search.
- Disable tips, tricks, suggestions & ads across Windows.
- Disable Windows Copilot. (Windows 11 only)
- Show file extensions for known file types.
- Disable the widget service & hide the icon from the taskbar. 
- Hide the Chat (meet now) icon from the taskbar.
- Hide the 3D objects folder under 'This pc' from File Explorer. (Windows 10 only)
- Hide the 'Include in library', 'Give access to' and 'Share' options from the context menu. (Windows 10 only)

'@


##################################################################################################################
#                                                                                                                #
#                                    EMBEDDED DATA - START MENU TEMPLATE                                         #
#                                                                                                                #
##################################################################################################################

# The start menu template (start2.bin) is a binary file that cannot be embedded as text.
# Instead, we create it programmatically or use a base64 encoded version.
# For this standalone version, we'll create a minimal start menu file when needed.

# Base64 encoded start2.bin (empty start menu for Windows 11)
$global:StartMenuTemplate = @'
0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAPgADAPz/CgAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAEAAA/v///wAAAAAA////AAAAAFEAAAAA////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////AAAA/v8AAAQAAP7/CgAAAAYJEgAA6AIAAA0AAAACAAAAAQAAAAEAAAACAAAAA9D+/wAAAAD+/wAAAAD+/wAAAAD+/wAAAAD+/wAAAAD+/wAAAAD+/wAAAAAAAACqAAAA/v8AAAD+/wAAAAD+/wAAAAD+/wAAAAD+/wAAAAD+/wAABgEBAAMAAAAFAAAA/////w0AAAAFAAAADQAAAAAAAACqAAAAAAAAAP7/AAAAAAAA/v8AAAD+/wAAAAD+/wAAAAD+/wAAAAD+/wAA
'@




##################################################################################################################
#                                                                                                                #
#                                          HELPER FUNCTIONS                                                      #
#                                                                                                                #
##################################################################################################################

# CmdletBinding for SupportsShouldProcess is added at the end with parameters

# Shows application selection form that allows the user to select what apps they want to remove or keep
function ShowAppSelectionForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    # Initialise form objects
    $form = New-Object System.Windows.Forms.Form
    $label = New-Object System.Windows.Forms.Label
    $button1 = New-Object System.Windows.Forms.Button
    $button2 = New-Object System.Windows.Forms.Button
    $selectionBox = New-Object System.Windows.Forms.CheckedListBox 
    $loadingLabel = New-Object System.Windows.Forms.Label
    $onlyInstalledCheckBox = New-Object System.Windows.Forms.CheckBox
    $checkUncheckCheckBox = New-Object System.Windows.Forms.CheckBox
    $initialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $global:selectionBoxIndex = -1

    # saveButton eventHandler
    $handler_saveButton_Click= 
    {
        if ($selectionBox.CheckedItems -contains "Microsoft.WindowsStore" -and -not $Silent) {
            $warningSelection = [System.Windows.Forms.Messagebox]::Show('Are you sure you wish to uninstall the Microsoft Store? This app cannot easily be reinstalled.', 'Are you sure?', 'YesNo', 'Warning')
        
            if ($warningSelection -eq 'No') {
                return
            }
        }

        $global:SelectedApps = $selectionBox.CheckedItems

        # Create file that stores selected apps if it doesn't exist
        if (!(Test-Path "$PSScriptRoot/CustomAppsList")) {
            $null = New-Item "$PSScriptRoot/CustomAppsList"
        } 

        Set-Content -Path "$PSScriptRoot/CustomAppsList" -Value $global:SelectedApps

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    # cancelButton eventHandler
    $handler_cancelButton_Click= 
    {
        $form.Close()
    }

    $selectionBox_SelectedIndexChanged= 
    {
        $global:selectionBoxIndex = $selectionBox.SelectedIndex
    }

    $selectionBox_MouseDown=
    {
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
                if ($global:selectionBoxIndex -ne -1) {
                    $topIndex = $global:selectionBoxIndex

                    if ($selectionBox.SelectedIndex -gt $topIndex) {
                        for (($i = ($topIndex)); $i -le $selectionBox.SelectedIndex; $i++){
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                    elseif ($topIndex -gt $selectionBox.SelectedIndex) {
                        for (($i = ($selectionBox.SelectedIndex)); $i -le $topIndex; $i++){
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                }
            }
            elseif ($global:selectionBoxIndex -ne $selectionBox.SelectedIndex) {
                $selectionBox.SetItemChecked($selectionBox.SelectedIndex, -not $selectionBox.GetItemChecked($selectionBox.SelectedIndex))
            }
        }
    }

    $check_All=
    {
        for (($i = 0); $i -lt $selectionBox.Items.Count; $i++){
            $selectionBox.SetItemChecked($i, $checkUncheckCheckBox.Checked)
        }
    }

    $load_Apps=
    {
        # Correct the initial state of the form to prevent the .Net maximized form issue
        $form.WindowState = $initialFormWindowState

        # Reset state to default before loading appslist again
        $global:selectionBoxIndex = -1
        $checkUncheckCheckBox.Checked = $False

        # Show loading indicator
        $loadingLabel.Visible = $true
        $form.Refresh()

        # Clear selectionBox before adding any new items
        $selectionBox.Items.Clear()

        # Build app list from embedded configuration
        $allApps = $global:AppConfig.DefaultRemove + $global:AppConfig.OptionalRemove + `
                   $global:AppConfig.GamingApps + $global:AppConfig.CommunicationApps + `
                   $global:AppConfig.DeveloperApps + $global:AppConfig.NewOutlookApp
        
        $listOfApps = ""

        if ($onlyInstalledCheckBox.Checked -and ($global:wingetInstalled -eq $true)) {
            # Attempt to get a list of installed apps via winget, times out after 10 seconds
            $job = Start-Job { return winget list --accept-source-agreements --disable-interactivity }
            $jobDone = $job | Wait-Job -TimeOut 10

            if (-not $jobDone) {
                # Show error that the script was unable to get list of apps from winget
                [System.Windows.MessageBox]::Show('Unable to load list of installed apps via winget, some apps may not be displayed in the list.', 'Error', 'Ok', 'Error')
            }
            else {
                # Add output of job (list of apps) to $listOfApps
                $listOfApps = Receive-Job -Job $job
            }
        }

        # Go through appslist and add items one by one to the selectionBox
        Foreach ($app in $allApps) { 
            # Check if app is in DefaultRemove (should be checked by default)
            $appChecked = $global:AppConfig.DefaultRemove -contains $app
            
            # Remove any leading/trailing spaces
            $app = $app.Trim()
            $appString = $app.Trim('*')

            # Make sure appString is not empty
            if ($appString.length -gt 0) {
                if ($onlyInstalledCheckBox.Checked) {
                    # onlyInstalledCheckBox is checked, check if app is installed before adding it to selectionBox
                    if (-not ($listOfApps -like ("*$appString*")) -and -not (Get-AppxPackage -Name $app)) {
                        # App is not installed, continue with next item
                        continue
                    }
                    if (($appString -eq "Microsoft.Edge") -and -not ($listOfApps -like "* Microsoft.Edge *")) {
                        # App is not installed, continue with next item
                        continue
                    }
                }

                # Add the app to the selectionBox and set it's checked status
                $selectionBox.Items.Add($appString, $appChecked) | Out-Null
            }
        }
        
        # Hide loading indicator
        $loadingLabel.Visible = $False

        # Sort selectionBox alphabetically
        $selectionBox.Sorted = $True
    }

    $form.Text = "Win11Debloat Application Selection"
    $form.Name = "appSelectionForm"
    $form.DataBindings.DefaultDataSourceUpdateMode = 0
    $form.ClientSize = New-Object System.Drawing.Size(400,502)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $False

    $button1.TabIndex = 4
    $button1.Name = "saveButton"
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "Confirm"
    $button1.Location = New-Object System.Drawing.Point(27,472)
    $button1.Size = New-Object System.Drawing.Size(75,23)
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    $button1.add_Click($handler_saveButton_Click)

    $form.Controls.Add($button1)

    $button2.TabIndex = 5
    $button2.Name = "cancelButton"
    $button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "Cancel"
    $button2.Location = New-Object System.Drawing.Point(129,472)
    $button2.Size = New-Object System.Drawing.Size(75,23)
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $button2.add_Click($handler_cancelButton_Click)

    $form.Controls.Add($button2)

    $label.Location = New-Object System.Drawing.Point(13,5)
    $label.Size = New-Object System.Drawing.Size(400,14)
    $Label.Font = 'Microsoft Sans Serif,8'
    $label.Text = 'Check apps that you wish to remove, uncheck apps that you wish to keep'

    $form.Controls.Add($label)

    $loadingLabel.Location = New-Object System.Drawing.Point(16,46)
    $loadingLabel.Size = New-Object System.Drawing.Size(300,418)
    $loadingLabel.Text = 'Loading apps...'
    $loadingLabel.BackColor = "White"
    $loadingLabel.Visible = $false

    $form.Controls.Add($loadingLabel)

    $onlyInstalledCheckBox.TabIndex = 6
    $onlyInstalledCheckBox.Location = New-Object System.Drawing.Point(230,474)
    $onlyInstalledCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $onlyInstalledCheckBox.Text = 'Only show installed apps'
    $onlyInstalledCheckBox.add_CheckedChanged($load_Apps)

    $form.Controls.Add($onlyInstalledCheckBox)

    $checkUncheckCheckBox.TabIndex = 7
    $checkUncheckCheckBox.Location = New-Object System.Drawing.Point(16,22)
    $checkUncheckCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $checkUncheckCheckBox.Text = 'Check/Uncheck all'
    $checkUncheckCheckBox.add_CheckedChanged($check_All)

    $form.Controls.Add($checkUncheckCheckBox)

    $selectionBox.FormattingEnabled = $True
    $selectionBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $selectionBox.Name = "selectionBox"
    $selectionBox.Location = New-Object System.Drawing.Point(13,43)
    $selectionBox.Size = New-Object System.Drawing.Size(374,424)
    $selectionBox.TabIndex = 3
    $selectionBox.add_SelectedIndexChanged($selectionBox_SelectedIndexChanged)
    $selectionBox.add_Click($selectionBox_MouseDown)

    $form.Controls.Add($selectionBox)

    # Save the initial state of the form
    $initialFormWindowState = $form.WindowState

    # Load apps into selectionBox
    $form.add_Load($load_Apps)

    # Focus selectionBox when form opens
    $form.Add_Shown({$form.Activate(); $selectionBox.Focus()})

    # Show the Form
    return $form.ShowDialog()
}


# Returns list of apps from embedded configuration or custom file
function ReadAppslistFromFile {
    param (
        $appsFilePath
    )

    $appsList = @()

    # If a custom file path is provided, read from that file
    if ($appsFilePath -and (Test-Path $appsFilePath)) {
        # Get list of apps from file at the path provided
        Foreach ($app in (Get-Content -Path $appsFilePath | Where-Object { $_ -notmatch '^#.*' -and $_ -notmatch '^\s*$' } )) { 
            # Remove any comments from the Appname
            if (-not ($app.IndexOf('#') -eq -1)) {
                $app = $app.Substring(0, $app.IndexOf('#'))
            }

            # Remove any spaces before and after the Appname
            $app = $app.Trim()
            
            $appString = $app.Trim('*')
            $appsList += $appString
        }
    }
    else {
        # Use embedded app configuration
        $appsList = $global:AppConfig.DefaultRemove
    }

    return $appsList
}


# Removes apps specified during function call from all user accounts and from the OS image.
function RemoveApps {
    param (
        $appslist
    )

    Foreach ($app in $appsList) { 
        Write-Output "Attempting to remove $app..."

        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            # Use winget to remove OneDrive and Edge
            if ($global:wingetInstalled -eq $false) {
                Write-Host "Error: WinGet is either not installed or is outdated, $app could not be removed" -ForegroundColor Red
            }
            else {
                # Uninstall app via winget
                Strip-Progress -ScriptBlock { winget uninstall --accept-source-agreements --disable-interactivity --id $app } | Tee-Object -Variable wingetOutput 

                If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "93")) {
                    Write-Host "Unable to uninstall Microsoft Edge via Winget" -ForegroundColor Red
                    Write-Output ""

                    if ($( Read-Host -Prompt "Would you like to forcefully uninstall Edge? NOT RECOMMENDED! (y/n)" ) -eq 'y') {
                        Write-Output ""
                        ForceRemoveEdge
                    }
                }
            }
        }
        else {
            # Use Remove-AppxPackage to remove all other apps
            $app = '*' + $app + '*'

            # Remove installed app for all existing users
            if ($WinVersion -ge 22000){
                # Windows 11 build 22000 or later
                try {
                    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue
                }
                catch {
                    Write-Host "Unable to remove $app for all users" -ForegroundColor Yellow
                    Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                }
            }
            else {
                # Windows 10
                try {
                    Get-AppxPackage -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Unable to remove $app for current user" -ForegroundColor Yellow
                    Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                }
                
                try {
                    Get-AppxPackage -Name $app -PackageTypeFilter Main, Bundle, Resource -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Unable to remove $app for all users" -ForegroundColor Yellow
                    Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                }
            }

            # Remove provisioned app from OS image, so the app won't be installed for any new users
            try {
                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $app } | ForEach-Object { Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName }
            }
            catch {
                Write-Host "Unable to remove $app from windows image" -ForegroundColor Yellow
                Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
            }
        }
    }
            
    Write-Output ""
}


# Forcefully removes Microsoft Edge using it's uninstaller
function ForceRemoveEdge {
    # Based on work from loadstring1 & ave9858
    Write-Output "> Forcefully uninstalling Microsoft Edge..."

    $regView = [Microsoft.Win32.RegistryView]::Registry32
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
    $hklm.CreateSubKey('SOFTWARE\Microsoft\EdgeUpdateDev').SetValue('AllowUninstall', '')

    # Create stub (Creating this somehow allows uninstalling edge)
    $edgeStub = "$env:SystemRoot\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe"
    New-Item $edgeStub -ItemType Directory | Out-Null
    New-Item "$edgeStub\MicrosoftEdge.exe" | Out-Null

    # Remove edge
    $uninstallRegKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge')
    if ($null -ne $uninstallRegKey) {
        Write-Output "Running uninstaller..."
        $uninstallString = $uninstallRegKey.GetValue('UninstallString') + ' --force-uninstall'
        Start-Process cmd.exe "/c $uninstallString" -WindowStyle Hidden -Wait

        Write-Output "Removing leftover files..."

        $edgePaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Tombstones\Microsoft Edge.lnk",
            "$env:PUBLIC\Desktop\Microsoft Edge.lnk",
            "$env:USERPROFILE\Desktop\Microsoft Edge.lnk",
            "$edgeStub"
        )

        foreach ($path in $edgePaths){
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force -Recurse -ErrorAction SilentlyContinue
                Write-Host "  Removed $path" -ForegroundColor DarkGray
            }
        }

        Write-Output "Cleaning up registry..."

        # Remove ms edge from autostart
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "Microsoft Edge Update" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "Microsoft Edge Update" /f *>$null

        Write-Output "Microsoft Edge was uninstalled"
    }
    else {
        Write-Output ""
        Write-Host "Error: Unable to forcefully uninstall Microsoft Edge, uninstaller could not be found" -ForegroundColor Red
    }
    
    Write-Output ""
}


# Execute provided command and strips progress spinners/bars from console output
function Strip-Progress {
    param(
        [ScriptBlock]$ScriptBlock
    )

    # Regex pattern to match spinner characters and progress bar patterns
    $progressPattern = 'Γû[Æê]|^\s+[-\\|/]\s+$'

    # Corrected regex pattern for size formatting, ensuring proper capture groups are utilized
    $sizePattern = '(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB) /\s+(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB)'

    & $ScriptBlock 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            "ERROR: $($_.Exception.Message)"
        } else {
            $line = $_ -replace $progressPattern, '' -replace $sizePattern, ''
            if (-not ([string]::IsNullOrWhiteSpace($line)) -and -not ($line.StartsWith('  '))) {
                $line
            }
        }
    }
}


# Import & execute embedded registry content
function RegImport {
    param (
        $message,
        $regFileName
    )

    Write-Output $message

    # Select the appropriate registry file collection
    if (!$global:Params.ContainsKey("Sysprep")) {
        $regContent = $global:EmbeddedRegFiles[$regFileName]
    }
    else {
        $regContent = $global:EmbeddedRegFilesSysprep[$regFileName]
    }

    if ([string]::IsNullOrWhiteSpace($regContent)) {
        Write-Host "Warning: Registry file '$regFileName' not found in embedded data" -ForegroundColor Yellow
        Write-Output ""
        return
    }

    # Create temporary registry file
    $tempRegFile = [System.IO.Path]::GetTempFileName() + ".reg"
    Set-Content -Path $tempRegFile -Value $regContent -Encoding Unicode

    # Import the registry file
    if (!$global:Params.ContainsKey("Sysprep")) {
        reg import $tempRegFile 2>&1 | Out-Null
    }
    else {
        $defaultUserPath = $env:USERPROFILE.Replace($env:USERNAME, 'Default\NTUSER.DAT')
        
        reg load "HKU\Default" $defaultUserPath | Out-Null
        reg import $tempRegFile 2>&1 | Out-Null
        reg unload "HKU\Default" | Out-Null
    }

    # Clean up temporary file
    Remove-Item -Path $tempRegFile -Force -ErrorAction SilentlyContinue

    Write-Output ""
}


# Restart the Windows Explorer process
function RestartExplorer {
    Write-Output "> Restarting Windows Explorer process to apply all changes... (This may cause some flickering)"

    # Only restart if the powershell process matches the OS architecture
    # Restarting explorer from a 32bit Powershell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem)
    {
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Warning "Unable to restart Windows Explorer process, please manually restart your PC to apply all changes."
    }
}


# Replace the startmenu for all users
function ReplaceStartMenuForAllUsers {
    param (
        $startMenuTemplate = $null
    )

    Write-Output "> Removing all pinned apps from the start menu for all users..."

    # Create start menu template from embedded base64 if not provided
    if ($null -eq $startMenuTemplate) {
        $tempStartFile = [System.IO.Path]::GetTempFileName()
        $bytes = [System.Convert]::FromBase64String($global:StartMenuTemplate)
        [System.IO.File]::WriteAllBytes($tempStartFile, $bytes)
        $startMenuTemplate = $tempStartFile
    }

    # Check if template bin file exists
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "Error: Unable to clear start menu, start2.bin file could not be created" -ForegroundColor Red
        Write-Output ""
        return
    }

    # Get path to start menu file for all users
    $userPathString = $env:USERPROFILE.Replace($env:USERNAME, "*\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState")
    $usersStartMenuPaths = get-childitem -path $userPathString

    # Go through all users and replace the start menu file
    ForEach ($startMenuPath in $usersStartMenuPaths) {
        ReplaceStartMenu "$($startMenuPath.Fullname)\start2.bin" $startMenuTemplate
    }

    # Also replace the start menu file for the default user profile
    $defaultStartMenuPath = $env:USERPROFILE.Replace($env:USERNAME, 'Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState')

    # Create folder if it doesn't exist
    if (-not(Test-Path $defaultStartMenuPath)) {
        new-item $defaultStartMenuPath -ItemType Directory -Force | Out-Null
        Write-Output "Created LocalState folder for default user profile"
    }

    # Copy template to default profile
    Copy-Item -Path $startMenuTemplate -Destination $defaultStartMenuPath -Force
    Write-Output "Replaced start menu for the default user profile"
    
    # Clean up temp file if we created one
    if ($null -ne $tempStartFile -and (Test-Path $tempStartFile)) {
        Remove-Item -Path $tempStartFile -Force -ErrorAction SilentlyContinue
    }
    
    Write-Output ""
}


# Replace the startmenu for current user
function ReplaceStartMenu {
    param (
        $startMenuBinFile = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin",
        $startMenuTemplate = $null
    )

    $userName = $startMenuBinFile.Split("\")[2]

    # Create start menu template from embedded base64 if not provided
    if ($null -eq $startMenuTemplate) {
        $tempStartFile = [System.IO.Path]::GetTempFileName()
        $bytes = [System.Convert]::FromBase64String($global:StartMenuTemplate)
        [System.IO.File]::WriteAllBytes($tempStartFile, $bytes)
        $startMenuTemplate = $tempStartFile
    }

    # Check if template bin file exists
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "Error: Unable to clear start menu, start2.bin file could not be created" -ForegroundColor Red
        return
    }

    # Check if bin file exists
    if (-not (Test-Path $startMenuBinFile)) {
        Write-Host "Error: Unable to clear start menu for user $userName, start2.bin file could not found" -ForegroundColor Red
        return
    }

    $backupBinFile = $startMenuBinFile + ".bak"

    # Backup current start menu file
    Move-Item -Path $startMenuBinFile -Destination $backupBinFile -Force

    # Copy template file
    Copy-Item -Path $startMenuTemplate -Destination $startMenuBinFile -Force

    Write-Output "Replaced start menu for user $userName"
    
    # Clean up temp file if we created one
    if ($null -ne $tempStartFile -and (Test-Path $tempStartFile)) {
        Remove-Item -Path $tempStartFile -Force -ErrorAction SilentlyContinue
    }
}


# Add parameter to script and write to file
function AddParameter {
    param (
        $parameterName,
        $message
    )

    # Add key if it doesn't already exist
    if (-not $global:Params.ContainsKey($parameterName)) {
        $global:Params.Add($parameterName, $true)
    }

    # Create or clear file that stores last used settings
    if (!(Test-Path "$PSScriptRoot/SavedSettings")) {
        $null = New-Item "$PSScriptRoot/SavedSettings"
    } 
    elseif ($global:FirstSelection) {
        $null = Clear-Content "$PSScriptRoot/SavedSettings"
    }
    
    $global:FirstSelection = $false

    # Create entry and add it to the file
    $entry = "$parameterName#- $message"
    Add-Content -Path "$PSScriptRoot/SavedSettings" -Value $entry
}


function PrintHeader {
    param (
        $title
    )

    $fullTitle = " Win11Debloat Script - $title"

    if ($global:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle (Sysprep mode)"
    }
    else {
        $fullTitle = "$fullTitle (User: $Env:UserName)"
    }

    Clear-Host
    Write-Output "-------------------------------------------------------------------------------------------"
    Write-Output $fullTitle
    Write-Output "-------------------------------------------------------------------------------------------"
}


function PrintFromText {
    param (
        $text
    )

    Clear-Host
    Write-Output $text
}


function AwaitKeyToExit {
    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output "Press any key to exit..."
        $null = [System.Console]::ReadKey()
    }
}


# Check if winget is installed & if it is, check if the version is at least v1.4
if ((Get-AppxPackage -Name "*Microsoft.DesktopAppInstaller*") -and ((winget -v) -replace 'v','' -gt 1.4)) {
    $global:wingetInstalled = $true
}
else {
    $global:wingetInstalled = $false

    # Show warning that requires user confirmation, Suppress confirmation if Silent parameter was passed
    if (-not $Silent) {
        Write-Warning "Winget is not installed or outdated. This may prevent Win11Debloat from removing certain apps."
        Write-Output ""
        Write-Output "Press any key to continue anyway..."
        $null = [System.Console]::ReadKey()
    }
}

# Get current Windows build version to compare against features
$WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild

$global:Params = $PSBoundParameters
$global:FirstSelection = $true
$SPParams = 'WhatIf', 'Confirm', 'Verbose', 'Silent', 'Sysprep'
$SPParamCount = 0

# Count how many SPParams exist within Params
# This is later used to check if any options were selected
foreach ($Param in $SPParams) {
    if ($global:Params.ContainsKey($Param)) {
        $SPParamCount++
    }
}

# Hide progress bars for app removal, as they block Win11Debloat's output
if (-not ($global:Params.ContainsKey("Verbose"))) {
    $ProgressPreference = 'SilentlyContinue'
}
else {
    Read-Host "Verbose mode is enabled, press enter to continue"
    $ProgressPreference = 'Continue'
}

if ($global:Params.ContainsKey("Sysprep")) {
    $defaultUserPath = $env:USERPROFILE.Replace($env:USERNAME, 'Default\NTUSER.DAT')

    # Exit script if default user directory or NTUSER.DAT file cannot be found
    if (-not (Test-Path "$defaultUserPath")) {
        Write-Host "Error: Unable to start Win11Debloat in Sysprep mode, cannot find default user folder at '$defaultUserPath'" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
    # Exit script if run in Sysprep mode on Windows 10
    if ($WinVersion -lt 22000) {
        Write-Host "Error: Win11Debloat Sysprep mode is not supported on Windows 10" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
}

# Remove SavedSettings file if it exists and is empty
if ((Test-Path "$PSScriptRoot/SavedSettings") -and ([String]::IsNullOrWhiteSpace((Get-content "$PSScriptRoot/SavedSettings")))) {
    Remove-Item -Path "$PSScriptRoot/SavedSettings" -recurse
}

# Only run the app selection form if the 'RunAppConfigurator' parameter was passed to the script
if ($RunAppConfigurator) {
    PrintHeader "App Configurator"

    $result = ShowAppSelectionForm

    # Show different message based on whether the app selection was saved or cancelled
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "App configurator was closed without saving." -ForegroundColor Red
    }
    else {
        Write-Output "Your app selection was saved to the 'CustomAppsList' file in the root folder of the script."
    }

    AwaitKeyToExit

    Exit
}

# Change script execution based on provided parameters or user input
if ((-not $global:Params.Count) -or $RunDefaults -or $RunWin11Defaults -or ($SPParamCount -eq $global:Params.Count)) {
    if ($RunDefaults -or $RunWin11Defaults) {
        $Mode = '1'
    }
    else {
        # Show menu and wait for user input, loops until valid input is provided
        Do { 
            $ModeSelectionMessage = "Please select an option (1/2/3/0)" 

            PrintHeader 'Menu'

            Write-Output "(1) Default Mode: Apply the default settings"
            Write-Output "(2) Custom Mode: Modify the script to your needs"
            Write-Output "(3) App removal mode: Select & remove apps, without making other changes"

            # Only show this option if SavedSettings file exists
            if (Test-Path "$PSScriptRoot/SavedSettings") {
                Write-Output "(4) Apply saved custom settings from last time"
                
                $ModeSelectionMessage = "Please select an option (1/2/3/4/0)" 
            }

            Write-Output ""
            Write-Output "(0) Show more information"
            Write-Output ""
            Write-Output ""

            $Mode = Read-Host $ModeSelectionMessage

            # Show information based on user input, Suppress user prompt if Silent parameter was passed
            if ($Mode -eq '0') {
                # Print script information from embedded text
                PrintFromText $global:InfoText

                Write-Output ""
                Write-Output "Press any key to go back..."
                $null = [System.Console]::ReadKey()
            }
            elseif (($Mode -eq '4')-and -not (Test-Path "$PSScriptRoot/SavedSettings")) {
                $Mode = $null
            }
        }
        while ($Mode -ne '1' -and $Mode -ne '2' -and $Mode -ne '3' -and $Mode -ne '4') 
    }

    # Add execution parameters based on the mode
    switch ($Mode) {
        # Default mode, loads defaults after confirmation
        '1' { 
            # Print the default settings & require userconfirmation, unless Silent parameter was passed
            if (-not $Silent) {
                PrintFromText $global:DefaultSettingsText

                Write-Output ""
                Write-Output "Press enter to execute the script or press CTRL+C to quit..."
                Read-Host | Out-Null
            }

            $DefaultParameterNames = 'RemoveApps','DisableTelemetry','DisableBing','DisableLockscreenTips','DisableSuggestions','ShowKnownFileExt','DisableWidgets','HideChat','DisableCopilot'

            PrintHeader 'Default Mode'

            # Add default parameters if they don't already exist
            foreach ($ParameterName in $DefaultParameterNames) {
                if (-not $global:Params.ContainsKey($ParameterName)){
                    $global:Params.Add($ParameterName, $true)
                }
            }

            # Only add this option for Windows 10 users, if it doesn't already exist
            if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -and (-not $global:Params.ContainsKey('Hide3dObjects'))) {
                $global:Params.Add('Hide3dObjects', $Hide3dObjects)
            }
        }

        # Custom mode, show & add options based on user input
        '2' { 
            # Get current Windows build version to compare against features
            $WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild
            
            PrintHeader 'Custom Mode'

            # Show options for removing apps, only continue on valid input
            Do {
                Write-Host "Options:" -ForegroundColor Yellow
                Write-Host " (n) Don't remove any apps" -ForegroundColor Yellow
                Write-Host " (1) Only remove the default selection of bloatware apps from embedded configuration" -ForegroundColor Yellow
                Write-Host " (2) Remove default selection of bloatware apps, aswell as mail & calendar apps, developer apps and gaming apps"  -ForegroundColor Yellow
                Write-Host " (3) Select which apps to remove and which to keep" -ForegroundColor Yellow
                $RemoveAppsInput = Read-Host "Remove any pre-installed apps? (n/1/2/3)" 

                # Show app selection form if user entered option 3
                if ($RemoveAppsInput -eq '3') {
                    $result = ShowAppSelectionForm

                    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                        # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                        Write-Output ""
                        Write-Host "Cancelled application selection, please try again" -ForegroundColor Red

                        $RemoveAppsInput = 'c'
                    }
                    
                    Write-Output ""
                }
            }
            while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2' -and $RemoveAppsInput -ne '3') 

            # Select correct option based on user input
            switch ($RemoveAppsInput) {
                '1' {
                    AddParameter 'RemoveApps' 'Remove default selection of bloatware apps'
                }
                '2' {
                    AddParameter 'RemoveApps' 'Remove default selection of bloatware apps'
                    AddParameter 'RemoveCommApps' 'Remove the Mail, Calendar, and People apps'
                    AddParameter 'RemoveW11Outlook' 'Remove the new Outlook for Windows app'
                    AddParameter 'RemoveDevApps' 'Remove developer-related apps'
                    AddParameter 'RemoveGamingApps' 'Remove the Xbox App and Xbox Gamebar'
                    AddParameter 'DisableDVR' 'Disable Xbox game/screen recording'
                }
                '3' {
                    Write-Output "You have selected $($global:SelectedApps.Count) apps for removal"

                    AddParameter 'RemoveAppsCustom' "Remove $($global:SelectedApps.Count) apps:"

                    Write-Output ""

                    if ($( Read-Host -Prompt "Disable Xbox game/screen recording? Also stops gaming overlay popups (y/n)" ) -eq 'y') {
                        AddParameter 'DisableDVR' 'Disable Xbox game/screen recording'
                    }
                }
            }

            # Only show this option for Windows 11 users running build 22621 or later
            if ($WinVersion -ge 22621){
                Write-Output ""

                if ($global:Params.ContainsKey("Sysprep")) {
                    if ($( Read-Host -Prompt "Remove all pinned apps from the start menu for all existing and new users? (y/n)" ) -eq 'y') {
                        AddParameter 'ClearStartAllUsers' 'Remove all pinned apps from the start menu for existing and new users'
                    }
                }
                else {
                    Do {
                        Write-Host "Options:" -ForegroundColor Yellow
                        Write-Host " (n) Don't remove any pinned apps from the start menu" -ForegroundColor Yellow
                        Write-Host " (1) Remove all pinned apps from the start menu for this user only ($env:USERNAME)" -ForegroundColor Yellow
                        Write-Host " (2) Remove all pinned apps from the start menu for all existing and new users"  -ForegroundColor Yellow
                        $ClearStartInput = Read-Host "Remove all pinned apps from the start menu? (n/1/2)" 
                    }
                    while ($ClearStartInput -ne 'n' -and $ClearStartInput -ne '0' -and $ClearStartInput -ne '1' -and $ClearStartInput -ne '2') 
    
                    # Select correct option based on user input
                    switch ($ClearStartInput) {
                        '1' {
                            AddParameter 'ClearStart' "Remove all pinned apps from the start menu for this user only"
                        }
                        '2' {
                            AddParameter 'ClearStartAllUsers' "Remove all pinned apps from the start menu for all existing and new users"
                        }
                    }
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "Disable telemetry, diagnostic data, activity history, app-launch tracking and targeted ads? (y/n)" ) -eq 'y') {
                AddParameter 'DisableTelemetry' 'Disable telemetry, diagnostic data, activity history, app-launch tracking & targeted ads'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "Disable tips, tricks, suggestions and ads in start, settings, notifications, explorer and lockscreen? (y/n)" ) -eq 'y') {
                AddParameter 'DisableSuggestions' 'Disable tips, tricks, suggestions and ads in start, settings, notifications and File Explorer'
                AddParameter 'DisableLockscreenTips' 'Disable tips & tricks on the lockscreen'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "Disable & remove bing web search, bing AI & cortana in Windows search? (y/n)" ) -eq 'y') {
                AddParameter 'DisableBing' 'Disable & remove bing web search, bing AI & cortana in Windows search'
            }

            # Only show this option for Windows 11 users running build 22621 or later
            if ($WinVersion -ge 22621){
                Write-Output ""

                if ($( Read-Host -Prompt "Disable and remove Windows Copilot? This applies to all users (y/n)" ) -eq 'y') {
                    AddParameter 'DisableCopilot' 'Disable and remove Windows Copilot'
                }

                Write-Output ""

                if ($( Read-Host -Prompt "Disable Windows Recall snapshots? This applies to all users (y/n)" ) -eq 'y') {
                    AddParameter 'DisableRecall' 'Disable Windows Recall snapshots'
                }
            }

            # Only show this option for Windows 11 users running build 22000 or later
            if ($WinVersion -ge 22000){
                Write-Output ""

                if ($( Read-Host -Prompt "Restore the old Windows 10 style context menu? (y/n)" ) -eq 'y') {
                    AddParameter 'RevertContextMenu' 'Restore the old Windows 10 style context menu'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "Do you want to make any changes to the taskbar and related services? (y/n)" ) -eq 'y') {
                # Only show these specific options for Windows 11 users running build 22000 or later
                if ($WinVersion -ge 22000){
                    Write-Output ""

                    if ($( Read-Host -Prompt "   Align taskbar buttons to the left side? (y/n)" ) -eq 'y') {
                        AddParameter 'TaskbarAlignLeft' 'Align taskbar icons to the left'
                    }

                    # Show options for search icon on taskbar, only continue on valid input
                    Do {
                        Write-Output ""
                        Write-Host "   Options:" -ForegroundColor Yellow
                        Write-Host "    (n) No change" -ForegroundColor Yellow
                        Write-Host "    (1) Hide search icon from the taskbar" -ForegroundColor Yellow
                        Write-Host "    (2) Show search icon on the taskbar" -ForegroundColor Yellow
                        Write-Host "    (3) Show search icon with label on the taskbar" -ForegroundColor Yellow
                        Write-Host "    (4) Show search box on the taskbar" -ForegroundColor Yellow
                        $TbSearchInput = Read-Host "   Hide or change the search icon on the taskbar? (n/1/2/3/4)" 
                    }
                    while ($TbSearchInput -ne 'n' -and $TbSearchInput -ne '0' -and $TbSearchInput -ne '1' -and $TbSearchInput -ne '2' -and $TbSearchInput -ne '3' -and $TbSearchInput -ne '4') 

                    # Select correct taskbar search option based on user input
                    switch ($TbSearchInput) {
                        '1' {
                            AddParameter 'HideSearchTb' 'Hide search icon from the taskbar'
                        }
                        '2' {
                            AddParameter 'ShowSearchIconTb' 'Show search icon on the taskbar'
                        }
                        '3' {
                            AddParameter 'ShowSearchLabelTb' 'Show search icon with label on the taskbar'
                        }
                        '4' {
                            AddParameter 'ShowSearchBoxTb' 'Show search box on the taskbar'
                        }
                    }

                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the taskview button from the taskbar? (y/n)" ) -eq 'y') {
                        AddParameter 'HideTaskview' 'Hide the taskview button from the taskbar'
                    }
                }

                Write-Output ""

                if ($( Read-Host -Prompt "   Disable the widgets service and hide the icon from the taskbar? (y/n)" ) -eq 'y') {
                    AddParameter 'DisableWidgets' 'Disable the widget service & hide the widget (news and interests) icon from the taskbar'
                }

                # Only show this options for Windows users running build 22621 or earlier
                if ($WinVersion -le 22621){
                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the chat (meet now) icon from the taskbar? (y/n)" ) -eq 'y') {
                        AddParameter 'HideChat' 'Hide the chat (meet now) icon from the taskbar'
                    }
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "Do you want to make any changes to File Explorer? (y/n)" ) -eq 'y') {
                Write-Output ""

                if ($( Read-Host -Prompt "   Show hidden files, folders and drives? (y/n)" ) -eq 'y') {
                    AddParameter 'ShowHiddenFolders' 'Show hidden files, folders and drives'
                }

                Write-Output ""

                if ($( Read-Host -Prompt "   Show file extensions for known file types? (y/n)" ) -eq 'y') {
                    AddParameter 'ShowKnownFileExt' 'Show file extensions for known file types'
                }

                # Only show this option for Windows 11 users running build 22000 or later
                if ($WinVersion -ge 22000){
                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the Home section from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                        AddParameter 'HideHome' 'Hide the Home section from the File Explorer sidepanel'
                    }

                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the Gallery section from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                        AddParameter 'HideGallery' 'Hide the Gallery section from the File Explorer sidepanel'
                    }
                }

                Write-Output ""

                if ($( Read-Host -Prompt "   Hide duplicate removable drive entries from the File Explorer sidepanel so they only show under This PC? (y/n)" ) -eq 'y') {
                    AddParameter 'HideDupliDrive' 'Hide duplicate removable drive entries from the File Explorer sidepanel'
                }

                # Only show option for disabling these specific folders for Windows 10 users
                if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'"){
                    Write-Output ""

                    if ($( Read-Host -Prompt "Do you want to hide any folders from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                        Write-Output ""

                        if ($( Read-Host -Prompt "   Hide the OneDrive folder from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                            AddParameter 'HideOnedrive' 'Hide the OneDrive folder in the File Explorer sidepanel'
                        }

                        Write-Output ""
                        
                        if ($( Read-Host -Prompt "   Hide the 3D objects folder from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                            AddParameter 'Hide3dObjects' "Hide the 3D objects folder under 'This pc' in File Explorer" 
                        }
                        
                        Write-Output ""

                        if ($( Read-Host -Prompt "   Hide the music folder from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                            AddParameter 'HideMusic' "Hide the music folder under 'This pc' in File Explorer"
                        }
                    }
                }
            }

            # Only show option for disabling context menu items for Windows 10 users or if the user opted to restore the Windows 10 context menu
            if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -or $global:Params.ContainsKey('RevertContextMenu')){
                Write-Output ""

                if ($( Read-Host -Prompt "Do you want to disable any context menu options? (y/n)" ) -eq 'y') {
                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the 'Include in library' option in the context menu? (y/n)" ) -eq 'y') {
                        AddParameter 'HideIncludeInLibrary' "Hide the 'Include in library' option in the context menu"
                    }

                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the 'Give access to' option in the context menu? (y/n)" ) -eq 'y') {
                        AddParameter 'HideGiveAccessTo' "Hide the 'Give access to' option in the context menu"
                    }

                    Write-Output ""

                    if ($( Read-Host -Prompt "   Hide the 'Share' option in the context menu? (y/n)" ) -eq 'y') {
                        AddParameter 'HideShare' "Hide the 'Share' option in the context menu"
                    }
                }
            }

            # Suppress prompt if Silent parameter was passed
            if (-not $Silent) {
                Write-Output ""
                Write-Output ""
                Write-Output ""
                Write-Output "Press enter to confirm your choices and execute the script or press CTRL+C to quit..."
                Read-Host | Out-Null
            }

            PrintHeader 'Custom Mode'
        }

        # App removal, remove apps based on user selection
        '3' {
            PrintHeader "App Removal"

            $result = ShowAppSelectionForm

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Output "You have selected $($global:SelectedApps.Count) apps for removal"
                AddParameter 'RemoveAppsCustom' "Remove $($global:SelectedApps.Count) apps:"

                # Suppress prompt if Silent parameter was passed
                if (-not $Silent) {
                    Write-Output ""
                    Write-Output "Press enter to remove the selected apps or press CTRL+C to quit..."
                    Read-Host | Out-Null
                    PrintHeader "App Removal"
                }
            }
            else {
                Write-Host "Selection was cancelled, no apps have been removed" -ForegroundColor Red
                Write-Output ""
            }
        }

        # Load custom options selection from the "SavedSettings" file
        '4' {
            if (-not $Silent) {
                PrintHeader 'Custom Mode'
                Write-Output "Win11Debloat will make the following changes:"

                # Get & print default settings info from file
                Foreach ($line in (Get-Content -Path "$PSScriptRoot/SavedSettings" )) { 
                    # Remove any spaces before and after the line
                    $line = $line.Trim()
                
                    # Check if the line contains a comment
                    if (-not ($line.IndexOf('#') -eq -1)) {
                        $parameterName = $line.Substring(0, $line.IndexOf('#'))

                        # Print parameter description and add parameter to Params list
                        if ($parameterName -eq "RemoveAppsCustom") {
                            if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                                # Apps file does not exist, skip
                                continue
                            }
                            
                            $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
                            Write-Output "- Remove $($appsList.Count) apps:"
                            Write-Host $appsList -ForegroundColor DarkGray
                        }
                        else {
                            Write-Output $line.Substring(($line.IndexOf('#') + 1), ($line.Length - $line.IndexOf('#') - 1))
                        }

                        if (-not $global:Params.ContainsKey($parameterName)){
                            $global:Params.Add($parameterName, $true)
                        }
                    }
                }

                Write-Output ""
                Write-Output ""
                Write-Output "Press enter to execute the script or press CTRL+C to quit..."
                Read-Host | Out-Null
            }

            PrintHeader 'Custom Mode'
        }
    }
}
else {
    PrintHeader 'Custom Mode'
}


# If the number of keys in SPParams equals the number of keys in Params then no modifications/changes were selected
#  or added by the user, and the script can exit without making any changes.
if ($SPParamCount -eq $global:Params.Keys.Count) {
    Write-Output "The script completed without making any changes."

    AwaitKeyToExit
}
else {
    # Execute all selected/provided parameters
    switch ($global:Params.Keys) {
        'RemoveApps' {
            $appsList = $global:AppConfig.DefaultRemove
            Write-Output "> Removing default selection of $($appsList.Count) apps..."
            RemoveApps $appsList
            continue
        }
        'RemoveAppsCustom' {
            if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                Write-Host "> Error: Could not load custom apps list from file, no apps were removed" -ForegroundColor Red
                Write-Output ""
                continue
            }
            
            $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
            Write-Output "> Removing $($appsList.Count) apps..."
            RemoveApps $appsList
            continue
        }
        'RemoveCommApps' {
            Write-Output "> Removing Mail, Calendar and People apps..."
            
            $appsList = $global:AppConfig.CommunicationApps
            RemoveApps $appsList
            continue
        }
        'RemoveW11Outlook' {
            $appsList = $global:AppConfig.NewOutlookApp
            Write-Output "> Removing new Outlook for Windows app..."
            RemoveApps $appsList
            continue
        }
        'RemoveDevApps' {
            $appsList = $global:AppConfig.DeveloperApps
            Write-Output "> Removing developer-related related apps..."
            RemoveApps $appsList
            continue
        }
        'RemoveGamingApps' {
            $appsList = $global:AppConfig.GamingApps
            Write-Output "> Removing gaming related apps..."
            RemoveApps $appsList
            continue
        }
        "ForceRemoveEdge" {
            ForceRemoveEdge
            continue
        }
        'DisableDVR' {
            RegImport "> Disabling Xbox game/screen recording..." "Disable_DVR"
            continue
        }
        'ClearStart' {
            Write-Output "> Removing all pinned apps from the start menu for user $env:USERNAME..."
            ReplaceStartMenu
            Write-Output ""
            continue
        }
        'ClearStartAllUsers' {
            ReplaceStartMenuForAllUsers
            continue
        }
        'DisableTelemetry' {
            RegImport "> Disabling telemetry, diagnostic data, activity history, app-launch tracking and targeted ads..." "Disable_Telemetry"
            continue
        }
        {$_ -in "DisableBingSearches", "DisableBing"} {
            RegImport "> Disabling bing web search, bing AI & cortana in Windows search..." "Disable_Bing_Cortana_In_Search"
            
            # Also remove the app package for bing search
            $appsList = @('Microsoft.BingSearch')
            RemoveApps $appsList
            continue
        }
        {$_ -in "DisableLockscrTips", "DisableLockscreenTips"} {
            RegImport "> Disabling tips & tricks on the lockscreen..." "Disable_Lockscreen_Tips"
            continue
        }
        {$_ -in "DisableSuggestions", "DisableWindowsSuggestions"} {
            RegImport "> Disabling tips, tricks, suggestions and ads across Windows..." "Disable_Windows_Suggestions"
            continue
        }
        'RevertContextMenu' {
            RegImport "> Restoring the old Windows 10 style context menu..." "Disable_Show_More_Options_Context_Menu"
            continue
        }
        'TaskbarAlignLeft' {
            RegImport "> Aligning taskbar buttons to the left..." "Align_Taskbar_Left"

            continue
        }
        'HideSearchTb' {
            RegImport "> Hiding the search icon from the taskbar..." "Hide_Search_Taskbar"
            continue
        }
        'ShowSearchIconTb' {
            RegImport "> Changing taskbar search to icon only..." "Show_Search_Icon"
            continue
        }
        'ShowSearchLabelTb' {
            RegImport "> Changing taskbar search to icon with label..." "Show_Search_Icon_And_Label"
            continue
        }
        'ShowSearchBoxTb' {
            RegImport "> Changing taskbar search to search box..." "Show_Search_Box"
            continue
        }
        'HideTaskview' {
            RegImport "> Hiding the taskview button from the taskbar..." "Hide_Taskview_Taskbar"
            continue
        }
        'DisableCopilot' {
            RegImport "> Disabling & removing Windows Copilot..." "Disable_Copilot"

            # Also remove the app package for copilot
            $appsList = @('Microsoft.Copilot')
            RemoveApps $appsList
            continue
        }
        'DisableRecall' {
            RegImport "> Disabling Windows Recall snapshots..." "Disable_AI_Recall"
            continue
        }
        {$_ -in "HideWidgets", "DisableWidgets"} {
            RegImport "> Disabling the widget service and hiding the widget icon from the taskbar..." "Disable_Widgets_Taskbar"
            continue
        }
        {$_ -in "HideChat", "DisableChat"} {
            RegImport "> Hiding the chat icon from the taskbar..." "Disable_Chat_Taskbar"
            continue
        }
        'ShowHiddenFolders' {
            RegImport "> Unhiding hidden files, folders and drives..." "Show_Hidden_Folders"
            continue
        }
        'ShowKnownFileExt' {
            RegImport "> Enabling file extensions for known file types..." "Show_Extensions_For_Known_File_Types"
            continue
        }
        'HideHome' {
            RegImport "> Hiding the home section from the File Explorer navigation pane..." "Hide_Home_from_Explorer"
            continue
        }
        'HideGallery' {
            RegImport "> Hiding the gallery section from the File Explorer navigation pane..." "Hide_Gallery_from_Explorer"
            continue
        }
        'HideDupliDrive' {
            RegImport "> Hiding duplicate removable drive entries from the File Explorer navigation pane..." "Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer"
            continue
        }
        {$_ -in "HideOnedrive", "DisableOnedrive"} {
            RegImport "> Hiding the OneDrive folder from the File Explorer navigation pane..." "Hide_Onedrive_Folder"
            continue
        }
        {$_ -in "Hide3dObjects", "Disable3dObjects"} {
            RegImport "> Hiding the 3D objects folder from the File Explorer navigation pane..." "Hide_3D_Objects_Folder"
            continue
        }
        {$_ -in "HideMusic", "DisableMusic"} {
            RegImport "> Hiding the music folder from the File Explorer navigation pane..." "Hide_Music_Folder"
            continue
        }
        {$_ -in "HideIncludeInLibrary", "DisableIncludeInLibrary"} {
            RegImport "> Hiding 'Include in library' in the context menu..." "Disable_Include_in_library_from_context_menu"
            continue
        }
        {$_ -in "HideGiveAccessTo", "DisableGiveAccessTo"} {
            RegImport "> Hiding 'Give access to' in the context menu..." "Disable_Give_access_to_context_menu"
            continue
        }
        {$_ -in "HideShare", "DisableShare"} {
            RegImport "> Hiding 'Share' in the context menu..." "Disable_Share_from_context_menu"
            continue
        }
    }

    RestartExplorer

    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output "Script completed successfully!"

    AwaitKeyToExit
}

