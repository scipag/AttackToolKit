VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Attack Tool Kit"
   ClientHeight    =   6900
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8835
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8160
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":290D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3047
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3763
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   1535
      ButtonWidth     =   1429
      ButtonHeight    =   1376
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            Object.ToolTipText     =   "Start the attack"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Stop"
            Object.ToolTipText     =   "Stop the running attack"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Config"
            Object.ToolTipText     =   "Open the configuration"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit the selected plugin"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reload"
            Object.ToolTipText     =   "Reload the selected plugin"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete the selected plugin"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Visualize"
            Object.ToolTipText     =   "Visualize the running attack"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Response"
            Object.ToolTipText     =   "Analyze the attack response"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logs"
            Object.ToolTipText     =   "Analyze the log files"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            Object.ToolTipText     =   "See and export the report"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSWinsockLib.Winsock wskTCPWinsock 
      Index           =   0
      Left            =   8280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timTimeout 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7800
      Top             =   1080
   End
   Begin VB.Frame fraPluginOverview 
      Caption         =   "Plugin Overview"
      Height          =   4935
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   4935
      Begin VB.TextBox txtPluginContent 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   4575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin MSComctlLib.ProgressBar pbrProgress 
      Height          =   120
      Left            =   7305
      TabIndex        =   5
      Top             =   6720
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fraPlugins 
      Caption         =   "Plugins"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3615
      Begin VB.FileListBox filNASLPlugins 
         Height          =   870
         Left            =   360
         Pattern         =   "*.nasl"
         TabIndex        =   9
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.FileListBox filATKPlugins 
         Height          =   870
         Left            =   360
         Pattern         =   "*.plugin"
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComctlLib.TreeView tvwPlugins 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8070
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   6
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   6585
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11078
            MinWidth        =   1765
            Text            =   "Ready"
            TextSave        =   "Ready"
            Object.ToolTipText     =   "Status message"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "100 %"
            TextSave        =   "100 %"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblVulnerabilityState 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "There was no vulnerability tested yet. Please run the selected plugin to determine the existence of the flaw."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmMain.frx":5080
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Click here to open the response analysis"
      Top             =   1080
      Width           =   8595
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExitItem 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuScan 
      Caption         =   "&Scan"
      Begin VB.Menu mnuScanStartItem 
         Caption         =   "&Start"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu mnuScanStopItem 
         Caption         =   "Sto&p"
         Enabled         =   0   'False
         Shortcut        =   +^{F2}
      End
   End
   Begin VB.Menu mnuConfiguration 
      Caption         =   "&Configuration"
      Begin VB.Menu mnuConfigurationPreferencesItem 
         Caption         =   "&Preferences..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuConfigurationToolbarItem 
         Caption         =   "&Toolbar..."
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPluginsRun 
         Caption         =   "&Run"
         Begin VB.Menu mnuPluginsRunDetectionItem 
            Caption         =   "&Detection"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPluginsRunExploitItem 
            Caption         =   "&Exploit"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuPluginsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsReloadAllItem 
         Caption         =   "Reload &all"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPluginsSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsEditItem 
         Caption         =   "&Edit..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuPluginsExternalEditorItem 
         Caption         =   "Edit with e&xternal editor..."
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuPluginsReloadItem 
         Caption         =   "Re&load"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuPluginsDeleteItem 
         Caption         =   "&Delete"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuPluginsSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsSearchPluginItem 
         Caption         =   "&Search plugin..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuPluginsFindNextItem 
         Caption         =   "Find &next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuPluginsSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsReportConfigurationItem 
         Caption         =   "Report &configuration..."
         Shortcut        =   +^{F4}
      End
      Begin VB.Menu mnuPluginsSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsShowOnlinePluginItem 
         Caption         =   "Show online plugin in the web &browser ..."
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu mnuPluginsDownloadTheLatestPluginsItem 
         Caption         =   "D&ownload the latest plugins..."
         Shortcut        =   +^{F5}
      End
      Begin VB.Menu mnuPluginsExportLoadedPluginListItem 
         Caption         =   "Ex&port loaded plugin list..."
         Shortcut        =   +^{F6}
      End
   End
   Begin VB.Menu mnuAnalysis 
      Caption         =   "&Analysis"
      Begin VB.Menu mnuAnalysisAttackVisualizingItem 
         Caption         =   "Attack &visualizing..."
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuAnalysisAttackResponseItem 
         Caption         =   "Attack &response..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAnalysisSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnalysisLogsItem 
         Caption         =   "&Logs..."
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuReporting 
      Caption         =   "&Reporting"
      Begin VB.Menu mnuReportingShowReportItem 
         Caption         =   "&Show report..."
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuReportingSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportingConfigurationItem 
         Caption         =   "&Configuration..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuNslookupItem 
         Caption         =   "&Nslookup..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuICMPPingItem 
         Caption         =   "&ICMP ping..."
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuPortscannerItem 
         Caption         =   "&Portscanner..."
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIndexItem 
         Caption         =   "&Index"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectWebSiteItem 
         Caption         =   "Project &web site"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuAboutItem 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.1 2005-01-23                                                           *
' * - Fixed the resizing bug with the progressbar if Windows XP design is given.     *
' * Version 4.1 2005-01-16                                                           *
' * - Replaced, whenever possible, the Default and Cancel buttons with a form key    *
' *   press preview function.                                                        *
' * Version 4.0 2005-01-02                                                           *
' * - Replaced the old Nessus plugins URLs with the new ones.                        *
' * Version 4.0 2004-12-27                                                           *
' * - Added the feature to load and show the latest nasl plugin if nasl plugins are  *
' *   available.                                                                     *
' * Version 4.0 2004-12-15                                                           *
' * - Added the Reports menu item and icon by Pascal Widmer.                         *
' * Version 4.0 2004-12-10                                                           *
' * - Replaced all vbCrLf with vbNewLine - Because these are a bit faster.           *
' * - Optimized some of the string functions (e.g. Mid and LCase).                   *
' * Version 3.1 2004-11-17                                                           *
' * - Added a routine to show also the plugin loading progress in the splash screen. *
' * - Fixed an error with the progress bar value during loading of the plugins.      *
' * Version 3.0 2004-11-13                                                           *
' * - Changed the context menu popup to a mouseup event.                             *
' * Version 3.0 2004-11-12                                                           *
' * - Added the plugin search.                                                       *
' * Version 3.0 2004-11-07                                                           *
' * - Fixed a bug in the sending command. It was not possible to use several new     *
' *   lines in one single send command. Also increased the speed of the send command.*
' * Version 3.0 2004-11-06                                                           *
' * - Corrected and enhanced the procedure type detection.                           *
' * - Added the possibility of opening web URLs by double clicking a http link in    *
' *   the plugin overview.                                                           *
' * Version 3.0 2004-11-05                                                           *
' * - Fixed the vbModeless bug if the Attack Editor is opened via the treeview.      *
' * Version 3.0 2004-11-04                                                           *
' * - Added default cancel buttons in the whole project. Most sub-frames can be      *
' *   closed by clicking the esc button now.                                         *
' * - Added a routine for the plugin autoupdate which detects new plugins. Only in   *
' *   this case the new available plugins are loaded.                                *
' * Version 3.0 2004-11-03                                                           *
' * - Added a better freeze frame handling for more resource intensive procedures.   *
' * Version 3.0 2004-11-01                                                           *
' * - Corrected the last errors with treeview actions if no element nor node is      *
' *   selected.                                                                      *
' * Version 3.0 2004-10-17                                                           *
' * - Changed the reload behaviour. If a plugin is selected, just the plugin is      *
' *   reloaded. But if no plugins are loaded, the whole plugin repository will be    *
' *   reloaded.                                                                      *
' * - Added a procedure to handle the check if there should the triggr not be found. *
' * Version 3.0 2004-10-16                                                           *
' * - Changed the whole new trigger behaviour. The trigger is not generally saved in *
' *   plugin_trigger anymore. Instead the triggers are the parameter of the pattern  *
' *   matching commands.                                                             *
' * - Optimized the form resize procedure to be a bit faster.                        *
' * Version 3.0 2004-10-15                                                           *
' * - Deleted all the Nessus stuff because it was not working at the current time.   *
' *   Sorry, I know, this is very sad but I will try to implement the whole Nessus   *
' *   stuff in ATK 4.0 or 5.0 - Please be patient!                                   *
' * Version 3.0 2004-09-30                                                           *
' * - Added the run command to let the ATK run shell based commands.                 *
' * - Put the increasing of the status bar after a command has been run. This        *
' *   prevents the software from beeing showing 100 % has reached but the last       *
' *   command is running.                                                            *
' * Version 3.0 2004-09-25                                                           *
' * - Fixed some errors if a special mouse click sequence is sent and if no treeview *
' *   element is selected. The whole checking should also be faster than the old.    *
' * Version 2.1 2004-09-09                                                           *
' * - Fixed an error for the progress bar if more than 199 plugins are loaded.       *
' * Version 2.1 2004-09-08                                                           *
' * - Added a checking routine for unsaved data in the attack editor if a new plugin *
' *   is loaded.                                                                     *
' * - Added a better error checking routine for CVE names.                           *
' * Version 2.1 2004-09-05                                                           *
' * - Changed the frame menu for configuration. Added the two points preferences and *
' *   toolbar.                                                                       *
' * - Also changed the click behavior of the toolbar so a customization works.       *
' * - For faster config access added a context menu for the toolbar menu.            *
' * Version 2.1 2004-09-04                                                           *
' * - Added the progress bar status 100 % if scan is aborded.                        *
' * - Corrected the progress bar during full audit.                                  *
' * Version 2.1 2004-09-03                                                           *
' * - Fixed a runtime error if the user is clicking the right mouse button in the    *
' *   plugin TreeView but there is no node selected.                                 *
' * Version 2.0 2004-08-24                                                           *
' * - Modified the form resize handling to put the progress bar on the right place.  *
' ************************************************************************************

Dim strPluginSearchText As String 'Used for the plugin search

Private Sub filATKPlugins_Click()
    'Read the selected plugin file as fast as possible
    Call ParseATKPlugin(ReadPluginFromFile(filATKPlugins.Filename, application_plugin_directory))
End Sub

Private Sub filNASLPlugins_Click()
    'Read the selected plugin file
    Call ParseNASLPlugin(ReadPluginFromFile(filNASLPlugins.Filename, application_plugin_directory))
End Sub

Private Sub Form_Activate()
    If Me.WindowState = vbMinimized Then
        Me.WindowState = vbNormal
    End If
End Sub

' ********************************************************************
' * Here is all the things that happen when the main form is loaded. *
' * I want to keep this par as small as possible and call external   *
' * routines if possible. Just the most important stuff is done here.*
' ********************************************************************

Private Sub Form_Load()
    Me.Caption = application_name
End Sub

Private Sub ValidatePluginInput()
    If LenB(plugin_port) = 0 Then
        WriteLogEntry "Checking the existence of the plugin_port data ...", 6
        Call errPluginDataMissing("plugin_port", plugin_filename, plugin_id)
'    ElseIf LenB(session_procedure_commands) < 8 Then
'        WriteLogEntry "Checking the existence and compatibelity of the plugin_request data ...", 6
'        Call errPluginDataMissing("plugin_request", plugin_filename, plugin_id)
    Else
        'Initiate the attack if everything is okay
        Call InitiateAttack
    End If
End Sub

' ***************************************************
' * This routine prepares everything for the check. *
' ***************************************************

Private Sub InitiateAttack()
    'Reset the progress
    SetProgress 0
    
    'Write the log entry
    WriteLogEntry "Starting the attack ...", 6
    
    'Read the actual status
    Call ReadText("Starting the attack. Please wait until the attempt is finished...")
    
    'Freeze the frame
    Call FreezeWindows
    
    'Close the last used socket - Just to be sure
    Call wskTCPWinsock(0).Close
    
    'Do ICMP mapping of wanted
    If application_icmp_mapping_enable Then
        Call ICMPMapping
    Else
        WriteLogEntry "No mapping wanted. Starting attack ...", 6
        Call InitiateCheckOrAudit
    End If
End Sub

Private Sub ICMPMapping()
    Dim ECHO As ICMP_ECHO_REPLY
    
    'ping an ip address, passing the
    'address and the ECHO structure
    Call Ping(GetIPFromHostName(Target), ECHO)
    WriteLogEntry "Sending ICMP echo request ...", 6
    
    'display the results from the ECHO structure
    If GetStatusCode(ECHO.status) = 0 Then
        WriteLogEntry "ICMP echo reply received in " & ECHO.RoundTripTime & " ms. Starting attack ...", 6
        Call InitiateCheckOrAudit
    ElseIf application_icmp_mapping_ignore_enable = True Then
        WriteLogEntry "No ICMP echo reply received. Starting attack ...", 5
        Call InitiateCheckOrAudit
    Else
        WriteLogEntry "No ICMP echo reply reveiced because " & msg & ". Ready.", 4
        'Enable the form to allow further input
        Call FreeWindows
    End If
End Sub

' *******************************************************************
' * This routine decides if a single check or audit should be done. *
' *******************************************************************

Private Sub InitiateCheckOrAudit()
    If application_attack_mode = "SingleCheck" Then
        'Delete the last attack response
        LastResponse = vbNullString
        
        'Start a single check
        Call AttackProcedure
    Else
        Dim i As Integer                    'This i is used for the counters
        Dim LoadedPlugins As Integer        'How many plugins are loaded

        LoadedPlugins = filATKPlugins.ListCount

        'Initiate a full security audit
        For i = 1 To LoadedPlugins
            'Check if the scan was stopped
            If Me.tlbMenu.Buttons.Item(2).Enabled = True Then
                'Delete the last attack response
                LastResponse = vbNullString

                'Everytime select the new plugin and do the check until finish
                filATKPlugins.ListIndex = i - 1
                SetProgress (100 / LoadedPlugins) * i
                Call AttackProcedure
            Else
                Exit For
            End If
        Next i
        SetProgress 100
    End If
    Call FreeWindows
End Sub

' *********************************************************
' * This routine starts and manages the attack procedure. *
' * It is the heart or the brain of the software.         *
' *********************************************************

Private Sub AttackProcedure()
    Dim i As Integer            'The counter
    Dim intFreeFile As Integer  'The free file integer
    Dim Command() As String     'The array with all commands of a plugin
    Dim CommandCount As Integer 'The number of commands in a row
    
    'Detect DoS and abord if needed
    If InStr(1, bug_vulnerability_class, "Denial of Service") Then
        If application_no_dos_enable = True Then
            'Message if the vulnerability was found
            WriteLogEntry "No denial of service checks activated. Abording check.", 3
            Call FreeWindows
            Exit Sub
        End If
    End If
    
    Call FreezeWindows
    
    'Define the selected request for the attack
    If session_procedure_type = "detection" Then
        session_procedure_commands = plugin_procedure_detection
    ElseIf session_procedure_type = "exploit" Then
        session_procedure_commands = plugin_procedure_exploit
    Else
        Call SetPluginSessionProcedure
    End If
    
    'Replace the ATK scripting language variants
    If InStrB(1, session_procedure_commands, "$DHOST", vbBinaryCompare) Then
        session_procedure_commands = Replace(session_procedure_commands, "$DHOST", Target, , , vbBinaryCompare)
    End If
    
    If InStrB(1, session_procedure_commands, "$DPORT", vbBinaryCompare) Then
        session_procedure_commands = Replace(session_procedure_commands, "$DPORT", plugin_port, , , vbBinaryCompare)
    End If
    
    'Split the commands in the request apart
    Command = Split(session_procedure_commands, "|")
    
    'Count the commands of this check
    CommandCount = UBound(Command)
   
    'Start the attack timeout timer
    timTimeout.Interval = application_attack_timeout
    timTimeout.Enabled = False
    timTimeout.Enabled = True

    For i = 0 To CommandCount
        'We need this if the timeout comes before a send command; I have to check this
        On Error Resume Next
        
        If Mid$(Command(i), 1, 4) = "open" Then
            Dim Try As Integer
            Dim OpenTarget As String
            
            'Check the target host
            If Len(Command(i)) > 4 Then
                OpenTarget = Mid$(Command(i), 6, Len(Command(i)))
            Else
                OpenTarget = Target
            End If
            
            'Open a new connection using the target data
            WriteLogEntry "Opening socket to " & OpenTarget & ":" & plugin_port, 6
            wskTCPWinsock(0).Connect OpenTarget, plugin_port
            
            If IsFormVisible("frmAttackVisualizing") = True Then
                Call frmAttackVisualizing.VisualizeOpenConnection
            End If
            
            'Wait a few moments for a successful connection
            Do While wskTCPWinsock(0).State <> sckConnected
                If Try < application_attack_timeout * 0.5 Then
                    Pause 1
                    Try = Try + 1000
                Else
                    Exit Do
                End If
            Loop
        
        ElseIf Mid$(Command(i), 1, 5) = "close" Then
            If timTimeout.Enabled = True Then
                'Call to close the socket
                Call wskTCPWinsock(0).Close
            End If
            
            If IsFormVisible("frmAttackVisualizing") = True Then
                Call frmAttackVisualizing.VisualizeCloseConnection
            End If
        
        ElseIf Mid$(Command(i), 1, 4) = "send" Then
            Dim DataToSend As String
            
            If wskTCPWinsock(0).State = 7 Then
                If Len(Command(i)) > 5 Then
                    DataToSend = Replace(Mid$(Command(i), 6, Len(Command(i))), "\n", vbNewLine, , , vbBinaryCompare)
    
                    'Send the request with its needed command and linefeeds
                    wskTCPWinsock(0).SendData DataToSend
                Else
                    'Send a "blank" request if the param1 is empty
                    DataToSend = vbNewLine
                    wskTCPWinsock(0).SendData DataToSend
                End If
                
                WriteLogEntry "Sending data """ & Mid$(DataToSend, 1, 64) & """ ...", 6
            
                If IsFormVisible("frmAttackVisualizing") = True Then
                    Call frmAttackVisualizing.VisualizeSendData(DataToSend)
                End If
            End If
        
        ElseIf Mid$(Command(i), 1, 5) = "sleep" Then
            If timTimeout.Enabled = True Then
                Dim SleepTime As Integer    'Save the time wanted to sleep
            
                If Len(Command(i)) > 5 Then
                    'Sleep as long as requested
                    SleepTime = (Mid$(Command(i), 7, Len(Command(i))))
                Else
                    'Sleep default seconds if parameter is missing
                    SleepTime = application_sleep_time_default / 1000
                End If
            
                If IsFormVisible("frmAttackVisualizing") = True Then
                    Call frmAttackVisualizing.VisualizeSleep(SleepTime)
                End If
                
                WriteLogEntry "Sleeping for " & SleepTime & " seconds ...", 6
                Pause (SleepTime)
            End If
        ElseIf Mid$(Command(i), 1, 8) = "pattern_" Then
            'Dev note: We have to visualize the search for the pattern before we run the
            'routines for found or not found. This is because we want to keep the order of
            'the visualizing.
            
            If Mid$(Command(i), 1, 14) = "pattern_exists" Then
                If Len(Command(i)) > 15 Then
                    session_triggers = Mid$(Command(i), 16, Len(Command(i)))
                    
                    If IsFormVisible("frmAttackVisualizing") = True Then
                        Call frmAttackVisualizing.VisualizePatternExists(session_triggers)
                    End If
                    
                    Call PatternExists(session_triggers)
                End If
            ElseIf Mid$(Command(i), 1, 18) = "pattern_not_exists" Then
                If Len(Command(i)) > 19 Then
                    session_triggers = Mid$(Command(i), 20, Len(Command(i)))
                    
                    If IsFormVisible("frmAttackVisualizing") = True Then
                        Call frmAttackVisualizing.VisualizePatternExists(session_triggers)
                    End If
                    
                    Call PatternNotExists(session_triggers)
                End If
            End If
        ElseIf Mid$(Command(i), 1, 10) = "icmp_alive" Then
            'Send ICMP ping
            Dim ECHO As ICMP_ECHO_REPLY
            
            'ping an ip address, passing the
            'address and the ECHO structure
            Call Ping(GetIPFromHostName(Target), ECHO)
              
            'display the results from the ECHO structure
            If GetStatusCode(ECHO.status) = 0 Then
                Call VulnerabilityNotFound
            Else
                Call VulnerabilityFound
            End If
        
        ElseIf Mid$(Command(i), 1, 3) = "run" Then
            Dim strRunCommand As String
            Dim strRunCommandFileName As String
        
            'get the selected command to run
            strRunCommand = (Mid$(Command(i), 5, Len(Command(i))))
            
            strRunCommandFileName = application_response_directory & Target & "-runcommandresponse.txt"
            
            'run the selected command
            Shell Environ("Comspec") + " /C " & strRunCommand & " > " & strRunCommandFileName, vbMinimizedNoFocus
            
            'wait until the command is finished
            Pause (application_sleep_time_default / 1000)
            
            'put the last response of the command run in the last response variant
            intFreeFile = FreeFile
            Open strRunCommandFileName For Input As #intFreeFile
                LastResponse = Input(LOF(intFreeFile), #intFreeFile)
            Close
            
            Call LoadLatestResponse
        End If
        
        'Add for every command the progress bar
        If application_attack_mode = "SingleCheck" Then
            SetProgress pbrProgress.Value + 100 / (CommandCount + 1)
        End If
        
    Next i
    
    'Finish the progress bar
    If application_attack_mode = "SingleCheck" Then
        SetProgress 100
    End If
End Sub

' *********************************************************************
' * This routine is the "brain" of a pattern-based check. Here is the *
' * decision made, if the pattern can be found in the server response.*
' *********************************************************************

Private Sub PatternExists(ByRef strPattern As String)
    Dim i As Integer            'The integer for the OR counter
    Dim Patterns() As String    'The array for multiple patterns
    Dim PatternCount As Integer 'The count of the patterns
    
    'Split the multiple OR patterns
    Patterns = Split(strPattern, " OR ")
    
    PatternCount = UBound(Patterns)
    
    'Check for the existence of one of the patterns
    For i = 0 To PatternCount
        'Check if the string DOES exists in the response; also do a
        'regulary expression check. One of them should recognize the flaw.
        If InStr(1, LastResponse, Patterns(i)) <> 0 Or _
            LastResponse Like Patterns(i) Then
            
            'Call the VulnFound procedure if the pattern was found
            Call VulnerabilityFound
            
            'Write the new pattern. This is needed to check the pattern
            'in the response window and to show the found pattern in
            'the scan report.
            session_trigger_match = Patterns(i)
            
            'Exit the sub if the vulnerability was found
            Exit Sub
        End If
    Next i
    
    'Call the VulnNotFound procedure if the pattern was not found
    Call VulnerabilityNotFound
End Sub

Private Sub PatternNotExists(ByRef strPattern As String)
    Dim i As Integer            'The integer for the OR counter
    Dim Patterns() As String    'The array for multiple patterns
    Dim PatternCount As Integer 'The count of the patterns
    
    'Split the multiple OR patterns
    Patterns = Split(strPattern, " OR ")
    
    PatternCount = UBound(Patterns)
    
    'Check for the existence of one of the patterns
    For i = 0 To PatternCount
        'Check if the string DOES exists in the response; also do a
        'regulary expression check. One of them should recognize the flaw.
        
        If InStr(1, LastResponse, Patterns(i)) <> 0 Or _
            LastResponse Like Patterns(i) Then
            
            'Call the VulnFound procedure if the pattern was found
            Call VulnerabilityNotFound
            
            'Write the new pattern. This is needed to check the pattern
            'in the response window and to show the found pattern in
            'the scan report.
            session_trigger_match = Patterns(i)
            
            'Exit the sub if the vulnerability was found
            Exit Sub
        End If
    Next i
    
    'Call the VulnNotFound procedure if the pattern was not found
    Call VulnerabilityFound
End Sub

' **********************************************************************
' * This routine calls everything that is needed, if the vulnerability *
' * could be found with the used check.                                *
' **********************************************************************

Private Sub VulnerabilityFound()
    Dim strAlertingText As String
    
    strAlertingText = "The vulnerability " & plugin_name & _
        " was found on port " & plugin_protocol & "/" & plugin_port & _
        " of the host " & Target & "."
    
    'Message if the vulnerability was found
    lblVulnerabilityState.Caption = strAlertingText
    lblVulnerabilityState.BackColor = &HC0C0FF
    WriteLogEntry "Vulnerability found! Ready.", 5
    
    'Write the pluginname into the report
    Call WritePluginNameToReportFile(plugin_filename & ";1;" & GetTodaysDate("/") & ";" & GetActualTime(":"))

    If IsFormVisible("frmAttackVisualizing") = True Then
        Call frmAttackVisualizing.VisualizeVulnerabilityFound
    End If

    'Show the alert message
    If application_vulnerability_found_alert_enable = True Then
        MsgBox strAlertingText, _
            vbExclamation, "Attack Tool Kit vulnerability found"
    End If

    'Speak the status that the vulnerability seems to be found
    Call ReadText("Check is finished. The vulnerability was found.")
End Sub

' **********************************************************************
' * This routine calls everything that is needed, if the vulnerability *
' * could not be found with the used check.                            *
' **********************************************************************
Private Sub VulnerabilityNotFound()
    Dim strAlertingText As String
    
    strAlertingText = "The vulnerability " & plugin_name & _
        " was not found on port " & plugin_protocol & "/" & plugin_port & _
        " of the host " & Target & "."
    
    'Message if the vulnerability was found
    lblVulnerabilityState.Caption = strAlertingText
    lblVulnerabilityState.BackColor = &HC0FFC0
    WriteLogEntry "Vulnerability not found. Ready.", 5

    'Write the pluginname into the report
    Call WritePluginNameToReportFile(plugin_filename & ";0;" & GetTodaysDate("/") & ";" & GetActualTime(":"))
    
    If IsFormVisible("frmAttackVisualizing") = True Then
        Call frmAttackVisualizing.VisualizeVulnerabilityNotFound
    End If

    'Show the alert message
    If application_vulnerability_not_found_alert_enable = True Then
        MsgBox "The vulnerability " & plugin_name & vbNewLine & _
        " was not found on port " & plugin_protocol & "/" & plugin_port & " of the host " & Target & ".", _
            vbInformation, "Attack Tool Kit vulnerability not found"
    End If
    
    Call ReadText("Check is finished. The vulnerability was not found.")
End Sub

' ******************************************************************
' * This routine freezes the window, so the user can't give input. *
' * The main reason is to prevent unexpected behaviour during      *
' * checks or other long-term procedures.                          *
' ******************************************************************

Private Sub FreezeWindows()
    'Show the hourglass cursor as cursor during checking
    Screen.MousePointer = 13
    
    'Freeze the window to disallow inputs during checking
    mnuTools.Enabled = False
    tlbMenu.Buttons.Item(1).Enabled = False
    tlbMenu.Buttons.Item(2).Enabled = True
    tlbMenu.Buttons.Item(4).Enabled = False
    tlbMenu.Buttons.Item(6).Enabled = False
    tlbMenu.Buttons.Item(7).Enabled = False
    tlbMenu.Buttons.Item(8).Enabled = False
    tvwPlugins.Enabled = False
    mnuScanStartItem.Enabled = False
    mnuScanStopItem.Enabled = True
    mnuConfiguration.Enabled = False
    mnuPlugins.Enabled = False
    
    If IsFormVisible("frmAttackEditor") = True Then
        frmAttackEditor.Enabled = False
    End If
End Sub

' **************************************************************
' * This routine frees the window, so the user can give input. *
' * This is always then done, when a long-term procedure (e.g. *
' * checking for a vulnerability) is finished.                 *
' **************************************************************

Private Sub FreeWindows()
    timTimeout.Enabled = False
    
    'Enable the form to allow further input
    mnuTools.Enabled = True
    tlbMenu.Buttons.Item(1).Enabled = True
    tlbMenu.Buttons.Item(2).Enabled = False
    tlbMenu.Buttons.Item(4).Enabled = True
    tlbMenu.Buttons.Item(6).Enabled = True
    tlbMenu.Buttons.Item(7).Enabled = True
    tlbMenu.Buttons.Item(8).Enabled = True
    tvwPlugins.Enabled = True
    mnuScanStartItem.Enabled = True
    mnuScanStopItem.Enabled = False
    mnuConfiguration.Enabled = True
    mnuPlugins.Enabled = True
    
    If IsFormVisible("frmAttackEditor") = True Then
        frmAttackEditor.Enabled = True
    End If
           
    'Show the normal cursor
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************************
' * Loading an ATK plugin into the list.  The procedure is public because *
' * there may a refresh needed after a plugin was edited in the attack    *
' * editor. This may be "fixed" in a further release.                     *
' *************************************************************************

Public Sub LoadATKPlugins()
    Dim i As Integer                        'Our counter
    Dim ListCountOfATKPlugins As Integer    'A listcount of ATK plugins to increase speed
    Dim intPercentageValue As Integer       'HEre we save the progress of loading
    Dim sPadding As String                  'Padding for some data
    Dim sCVEorCAN As String                 'Needed to display the CVE data with detailed infos
    
    'Check the existence of the plugin directory
    If (Dir$(application_plugin_directory, 16) <> "") Then
        'Set the right plugin directory
        filATKPlugins.Path = application_plugin_directory
        
        'Count the loadable plugins
        ListCountOfATKPlugins = filATKPlugins.ListCount
        
        'Error message if no plugins are available
        If ListCountOfATKPlugins Then
            'Reset the progress bar
            SetProgress 0
            
            'Load the procedure to load the plugins
            WriteLogEntry "Loading the plugins from " & application_plugin_directory & ".", 6
            
            On Error Resume Next    'Prevent errors with plugins with missing needed fields
            tvwPlugins.Nodes.Add , , "ATK plugins", "ATK plugins"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK ID", "ID"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK Name", "Name"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK Port", "Port"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK Severity", "Severity"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK Family", "Family"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK Class", "Class"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK CVE", "CVE"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK Nessus", "Nessus ID"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK SecurityFocus", "SecurityFocus BID"
                tvwPlugins.Nodes.Add "ATK plugins", tvwChild, "ATK OSVDB", "OSVDB ID"
                        
            'load the data into the TreeView
            For i = 1 To ListCountOfATKPlugins
                filATKPlugins.ListIndex = i - 1
                        
                intPercentageValue = (100 / ListCountOfATKPlugins) * i
                If intPercentageValue <= 100 Then
                    If IsFormVisible("frmSplashScreen") = True Then
                        frmSplashScreen.pbrStatus = intPercentageValue
                        frmSplashScreen.lblStatusInformation.Caption = "loading plugin " & i & " of " & ListCountOfATKPlugins & " (" & intPercentageValue & " %) ..."
                    End If
                            
                    SetProgress intPercentageValue
                End If
        
                'Add the name sub tree
                tvwPlugins.Nodes.Add "ATK Name", tvwChild, "n" & filATKPlugins.Filename, plugin_name
                
                'Add the id sub tree
                If DoesNodeExsist(plugin_id) = False Then
                    sPadding = vbNullString
                    If Len(plugin_id) = 1 Then
                        sPadding = "    "
                    ElseIf Len(plugin_id) = 2 Then
                        sPadding = "   "
                    ElseIf Len(plugin_id) = 3 Then
                        sPadding = "  "
                    ElseIf Len(plugin_id) = 4 Then
                        sPadding = " "
                    End If
                    tvwPlugins.Nodes.Add "ATK ID", tvwChild, "i" & filATKPlugins.Filename, sPadding & plugin_id
                End If
                
                'Add the severity sub tree
                If DoesNodeExsist("p" & plugin_port) = False Then
                    sPadding = vbNullString
                    If Len(plugin_port) = 1 Then
                        sPadding = "    "
                    ElseIf Len(plugin_port) = 2 Then
                        sPadding = "   "
                    ElseIf Len(plugin_port) = 3 Then
                        sPadding = "  "
                    ElseIf Len(plugin_port) = 4 Then
                        sPadding = " "
                    End If
                    tvwPlugins.Nodes.Add "ATK Port", tvwChild, "p" & plugin_port, sPadding & plugin_port
                End If
                tvwPlugins.Nodes.Add "p" & plugin_port, tvwChild, "p" & filATKPlugins.Filename, plugin_name
                
                'Add the severity sub tree
                If DoesNodeExsist(bug_severity) = False Then
                    tvwPlugins.Nodes.Add "ATK Severity", tvwChild, bug_severity, bug_severity
                End If
                tvwPlugins.Nodes.Add bug_severity, tvwChild, "s" & filATKPlugins.Filename, plugin_name
                
                'Add the family sub tree
                If DoesNodeExsist("f" & plugin_family) = False Then
                    tvwPlugins.Nodes.Add "ATK Family", tvwChild, "f" & plugin_family, plugin_family
                End If
                tvwPlugins.Nodes.Add "f" & plugin_family, tvwChild, "f" & filATKPlugins.Filename, plugin_name
                
                'Add the class sub tree
                If DoesNodeExsist("c" & bug_vulnerability_class) = False Then
                    tvwPlugins.Nodes.Add "ATK Class", tvwChild, "c" & bug_vulnerability_class, bug_vulnerability_class
                End If
                tvwPlugins.Nodes.Add "c" & bug_vulnerability_class, tvwChild, "c" & filATKPlugins.Filename, plugin_name
                
                'Add the CVE sub tree
                If LenB(source_cve) <> 0 Then
                    If LenB(source_cve) = 26 Then
                        If InStr(1, source_cve, "CVE") Then
                            sCVEorCAN = "CVE"
                        ElseIf InStr(1, source_cve, "CAN") Then
                            sCVEorCAN = "CAN"
                        Else
                            sCVEorCAN = "unknown"
                        End If
                        
                        If DoesNodeExsist("v" & source_cve) = False Then
                            tvwPlugins.Nodes.Add "ATK CVE", tvwChild, "v" & filATKPlugins.Filename, Mid$(source_cve, 5, Len(source_cve)) & " (" & sCVEorCAN & ")"
                        End If
                        tvwPlugins.Nodes.Add "v" & source_cve, tvwChild, "v" & filATKPlugins.Filename, Mid$(source_cve, 5, Len(source_cve)) & " (" & sCVEorCAN & ")"
                    Else
                        If DoesNodeExsist("v" & source_cve) = False Then
                            tvwPlugins.Nodes.Add "ATK CVE", tvwChild, "v" & filATKPlugins.Filename, source_cve & " (undefined)"
                        End If
                        tvwPlugins.Nodes.Add "v" & source_cve, tvwChild, "v" & filATKPlugins.Filename, source_cve & " (undefined)"
                    End If
                End If
            
                'Add the Nessus sub tree
                If LenB(source_nessus_id) <> 0 Then
                    If DoesNodeExsist("u" & source_nessus_id) = False Then
                        tvwPlugins.Nodes.Add "ATK Nessus", tvwChild, "u" & filATKPlugins.Filename, source_nessus_id
                    End If
                    tvwPlugins.Nodes.Add "u" & source_nessus_id, tvwChild, "u" & filATKPlugins.Filename, source_nessus_id
                End If
            
                'Add the SecurityFocus sub tree
                If LenB(source_securityfocus_bid) <> 0 Then
                    sPadding = vbNullString
                    If Len(source_securityfocus_bid) = 1 Then
                        sPadding = "     "
                    ElseIf Len(source_securityfocus_bid) = 2 Then
                        sPadding = "    "
                    ElseIf Len(source_securityfocus_bid) = 3 Then
                        sPadding = "   "
                    ElseIf Len(source_securityfocus_bid) = 4 Then
                        sPadding = "  "
                    ElseIf Len(source_securityfocus_bid) = 5 Then
                        sPadding = " "
                    End If

                    If DoesNodeExsist("b" & source_securityfocus_bid) = False Then
                        tvwPlugins.Nodes.Add "ATK SecurityFocus", tvwChild, "b" & filATKPlugins.Filename, sPadding & source_securityfocus_bid
                    End If
                    tvwPlugins.Nodes.Add "b" & source_securityfocus_bid, tvwChild, "b" & filATKPlugins.Filename, sPadding & source_securityfocus_bid
                End If
            
                'Add the OSVDB sub tree
                If LenB(source_osvdb_id) <> 0 Then
                    sPadding = vbNullString
                    If Len(source_osvdb_id) = 1 Then
                        sPadding = "     "
                    ElseIf Len(source_osvdb_id) = 2 Then
                        sPadding = "    "
                    ElseIf Len(source_osvdb_id) = 3 Then
                        sPadding = "   "
                    ElseIf Len(source_osvdb_id) = 4 Then
                        sPadding = "  "
                    ElseIf Len(source_osvdb_id) = 5 Then
                        sPadding = " "
                    End If

                    If DoesNodeExsist("o" & source_osvdb_id) = False Then
                        tvwPlugins.Nodes.Add "ATK OSVDB", tvwChild, "o" & filATKPlugins.Filename, sPadding & source_osvdb_id
                    End If
                    tvwPlugins.Nodes.Add "o" & source_osvdb_id, tvwChild, "o" & filATKPlugins.Filename, sPadding & source_osvdb_id
                End If
            
            DoEvents
            
            Next i
        
            fraPlugins.Caption = "Plugins (" & HowManyLoadedPlugins & " loaded)"
            WriteLogEntry HowManyLoadedPlugins & " plugins loaded. Ready.", 6
            
            'Sort the loaded data
            Call SortTreeViewNodes
            
            'Expand the first node
            tvwPlugins.Nodes(1).Expanded = True
            
            'Load the first plugin
            Call ParseATKPlugin(ReadPluginFromFile(plugin_filename, plugin_filepath))
            Call PrepareTheNewPluginData
        Else
            Call errPluginsDirectoryEmpty
        End If
    Else
        Call errPluginsDirectoryNotExist
        Call errPluginsDirectoryEmpty
    End If

    SetProgress 100
End Sub

Public Sub LoadNASLPlugins()
    Dim i As Integer                        'Our counter
    Dim ListCountOfNASLPlugins As Integer   'A listcount of NASL plugins to increase speed
    Dim intPercentageValue As Integer       'HEre we save the progress of loading
    Dim sPadding As String                  'Padding for some data
    Dim sCVEorCAN As String                 'Needed to display the CVE data with detailed infos
    
    'Check the existence of the plugin directory
    If (Dir$(application_plugin_directory, 16) <> "") Then
        'Set the right plugin directory
        filNASLPlugins.Path = application_plugin_directory
        
        'Count the loadable plugins
        ListCountOfNASLPlugins = filNASLPlugins.ListCount
        
        'Error message if no plugins are available
        If ListCountOfNASLPlugins Then
            'Reset the progress bar
            SetProgress 0
            
            'Load the procedure to load the plugins
            WriteLogEntry "Loading the plugins from " & application_plugin_directory & ".", 6
            
            On Error Resume Next    'Prevent errors with plugins with missing needed fields
            tvwPlugins.Nodes.Add , , "NASL plugins", "NASL plugins"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL ID", "ID"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL Name", "Name"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL Port", "Port"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL Severity", "Severity"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL Family", "Family"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL Class", "Class"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL CVE", "CVE"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL Nessus", "Nessus ID"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL SecurityFocus", "SecurityFocus BID"
                tvwPlugins.Nodes.Add "NASL plugins", tvwChild, "NASL OSVDB", "OSVDB ID"
                        
            'load the data into the TreeView
            For i = 1 To ListCountOfNASLPlugins
                filNASLPlugins.ListIndex = i - 1
                
                intPercentageValue = (100 / ListCountOfNASLPlugins) * i
                If intPercentageValue <= 100 Then
                    If IsFormVisible("frmSplashScreen") = True Then
                        frmSplashScreen.pbrStatus = intPercentageValue
                        frmSplashScreen.lblStatusInformation.Caption = "loading plugin " & i & " of " & ListCountOfNASLPlugins & " (" & intPercentageValue & " %) ..."
                    End If
                            
                    SetProgress intPercentageValue
                End If
        
                'Add the name sub tree
                tvwPlugins.Nodes.Add "NASL Name", tvwChild, "n" & filNASLPlugins.Filename, plugin_name
                
                'Add the id sub tree
                If DoesNodeExsist(plugin_id) = False Then
                    sPadding = vbNullString
                    If Len(plugin_id) = 1 Then
                        sPadding = "    "
                    ElseIf Len(plugin_id) = 2 Then
                        sPadding = "   "
                    ElseIf Len(plugin_id) = 3 Then
                        sPadding = "  "
                    ElseIf Len(plugin_id) = 4 Then
                        sPadding = " "
                    End If
                    tvwPlugins.Nodes.Add "NASL ID", tvwChild, "i" & filNASLPlugins.Filename, sPadding & plugin_id
                End If
                
                'Add the severity sub tree
                If DoesNodeExsist("p" & plugin_port) = False Then
                    sPadding = vbNullString
                    If Len(plugin_port) = 1 Then
                        sPadding = "    "
                    ElseIf Len(plugin_port) = 2 Then
                        sPadding = "   "
                    ElseIf Len(plugin_port) = 3 Then
                        sPadding = "  "
                    ElseIf Len(plugin_port) = 4 Then
                        sPadding = " "
                    End If
                    tvwPlugins.Nodes.Add "NASL Port", tvwChild, "p" & plugin_port, sPadding & plugin_port
                End If
                tvwPlugins.Nodes.Add "p" & plugin_port, tvwChild, "p" & filNASLPlugins.Filename, plugin_name
                
                'Add the severity sub tree
                If DoesNodeExsist(bug_severity) = False Then
                    tvwPlugins.Nodes.Add "NASL Severity", tvwChild, bug_severity, bug_severity
                End If
                tvwPlugins.Nodes.Add bug_severity, tvwChild, "s" & filNASLPlugins.Filename, plugin_name
                
                'Add the family sub tree
                If DoesNodeExsist("f" & plugin_family) = False Then
                    tvwPlugins.Nodes.Add "NASL Family", tvwChild, "f" & plugin_family, plugin_family
                End If
                tvwPlugins.Nodes.Add "f" & plugin_family, tvwChild, "f" & filNASLPlugins.Filename, plugin_name
                
                'Add the class sub tree
                If DoesNodeExsist("c" & bug_vulnerability_class) = False Then
                    tvwPlugins.Nodes.Add "NASL Class", tvwChild, "c" & bug_vulnerability_class, bug_vulnerability_class
                End If
                tvwPlugins.Nodes.Add "c" & bug_vulnerability_class, tvwChild, "c" & filNASLPlugins.Filename, plugin_name
                
                'Add the CVE sub tree
                If LenB(source_cve) <> 0 Then
                    If LenB(source_cve) = 26 Then
                        If InStr(1, source_cve, "CVE") Then
                            sCVEorCAN = "CVE"
                        ElseIf InStr(1, source_cve, "CAN") Then
                            sCVEorCAN = "CAN"
                        Else
                            sCVEorCAN = "unknown"
                        End If
                        
                        If DoesNodeExsist("v" & source_cve) = False Then
                            tvwPlugins.Nodes.Add "NASL CVE", tvwChild, "v" & filNASLPlugins.Filename, Mid$(source_cve, 5, Len(source_cve)) & " (" & sCVEorCAN & ")"
                        End If
                        tvwPlugins.Nodes.Add "v" & source_cve, tvwChild, "v" & filNASLPlugins.Filename, Mid$(source_cve, 5, Len(source_cve)) & " (" & sCVEorCAN & ")"
                    Else
                        If DoesNodeExsist("v" & source_cve) = False Then
                            tvwPlugins.Nodes.Add "NASL CVE", tvwChild, "v" & filNASLPlugins.Filename, source_cve & " (undefined)"
                        End If
                        tvwPlugins.Nodes.Add "v" & source_cve, tvwChild, "v" & filNASLPlugins.Filename, source_cve & " (undefined)"
                    End If
                End If
            
                'Add the Nessus sub tree
                If LenB(source_nessus_id) <> 0 Then
                    If DoesNodeExsist("u" & source_nessus_id) = False Then
                        tvwPlugins.Nodes.Add "NASL Nessus", tvwChild, "u" & filNASLPlugins.Filename, source_nessus_id
                    End If
                    tvwPlugins.Nodes.Add "u" & source_nessus_id, tvwChild, "u" & filNASLPlugins.Filename, source_nessus_id
                End If
            
                'Add the SecurityFocus sub tree
                If LenB(source_securityfocus_bid) <> 0 Then
                    sPadding = vbNullString
                    If Len(source_securityfocus_bid) = 1 Then
                        sPadding = "     "
                    ElseIf Len(source_securityfocus_bid) = 2 Then
                        sPadding = "    "
                    ElseIf Len(source_securityfocus_bid) = 3 Then
                        sPadding = "   "
                    ElseIf Len(source_securityfocus_bid) = 4 Then
                        sPadding = "  "
                    ElseIf Len(source_securityfocus_bid) = 5 Then
                        sPadding = " "
                    End If

                    If DoesNodeExsist("b" & source_securityfocus_bid) = False Then
                        tvwPlugins.Nodes.Add "NASL SecurityFocus", tvwChild, "b" & filNASLPlugins.Filename, sPadding & source_securityfocus_bid
                    End If
                    tvwPlugins.Nodes.Add "b" & source_securityfocus_bid, tvwChild, "b" & filNASLPlugins.Filename, sPadding & source_securityfocus_bid
                End If
            
                'Add the OSVDB sub tree
                If LenB(source_osvdb_id) <> 0 Then
                    sPadding = vbNullString
                    If Len(source_osvdb_id) = 1 Then
                        sPadding = "     "
                    ElseIf Len(source_osvdb_id) = 2 Then
                        sPadding = "    "
                    ElseIf Len(source_osvdb_id) = 3 Then
                        sPadding = "   "
                    ElseIf Len(source_osvdb_id) = 4 Then
                        sPadding = "  "
                    ElseIf Len(source_osvdb_id) = 5 Then
                        sPadding = " "
                    End If

                    If DoesNodeExsist("o" & source_osvdb_id) = False Then
                        tvwPlugins.Nodes.Add "NASL OSVDB", tvwChild, "o" & filNASLPlugins.Filename, sPadding & source_osvdb_id
                    End If
                    tvwPlugins.Nodes.Add "o" & source_osvdb_id, tvwChild, "o" & filNASLPlugins.Filename, sPadding & source_osvdb_id
                End If
            
            DoEvents
            
            Next i
        
            fraPlugins.Caption = "Plugins (" & HowManyLoadedPlugins & " loaded)"
            WriteLogEntry HowManyLoadedPlugins & " plugins loaded. Ready.", 6
            
            'Sort the loaded data
            Call SortTreeViewNodes
                        
            'Expand the first node
            tvwPlugins.Nodes(1).Expanded = True
            
            'Load the first plugin
            Call ParseNASLPlugin(ReadPluginFromFile(plugin_filename, plugin_filepath))
            Call PrepareTheNewPluginData
'        Else
'            Call errPluginsDirectoryEmpty
        End If
    Else
        Call errPluginsDirectoryNotExist
        Call errPluginsDirectoryEmpty
    End If

    SetProgress 100
End Sub

Private Sub SortTreeViewNodes()
    Dim i As Integer
    Dim intNodesCount As Integer
    
    intNodesCount = tvwPlugins.Nodes.Count
    
    For i = 2 To intNodesCount
        'Sort the loaded data if childrens are given
        If tvwPlugins.Nodes(i).Children Then
            'Do not the severity children
            If InStr(1, tvwPlugins.Nodes(i).Key, "Severity", vbBinaryCompare) = 0 Then
                'Do not sort the NASL plugins root
                If InStr(1, tvwPlugins.Nodes(i).Key, "NASL plugins", vbBinaryCompare) = 0 Then
                    tvwPlugins.Nodes(i).Sorted = True
                End If
            End If
        End If
    Next i
End Sub

Function DoesNodeExsist(Key As String) As Boolean
    DoesNodeExsist = False
    On Local Error GoTo errhand
    Call TypeName(tvwPlugins.Nodes(Key))

    DoesNodeExsist = True
    Exit Function
errhand:
End Function

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Form_Resize()
    'Check the window state. Do not resize if the window is minimized
    If frmMain.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If frmMain.Height < 6780 Then
            frmMain.Height = 6780
        End If
        
        'Prevent zu small windows in width
        If frmMain.Width < 8055 Then
            frmMain.Width = 8055
        End If
        
        'Do the resizing for the plugins frame
        fraPlugins.Width = frmMain.Width / 2.5

        fraPlugins.Height = frmMain.Height - 2800
        tvwPlugins.Height = fraPlugins.Height - 360
        
        'The listview of the plugins
        tvwPlugins.Width = fraPlugins.Width - 260
        
        'The plugin overview frame
        fraPluginOverview.Left = fraPlugins.Width + 260
        
        fraPluginOverview.Width = frmMain.Width - fraPlugins.Width - 460
        fraPluginOverview.Height = fraPlugins.Height
        
        txtPluginContent.Width = fraPluginOverview.Width - 260
        txtPluginContent.Height = fraPluginOverview.Height - 360
           
        lblVulnerabilityState.Width = fraPluginOverview.Width + fraPlugins.Width + 140
        
        'The progress bar
        pbrProgress.Top = frmMain.ScaleHeight - 210
        pbrProgress.Left = frmMain.Width - (pbrProgress.Width + 520)
    End If
End Sub

Private Sub lblVulnerabilityState_Click()
    frmAttackResponse.Show vbModeless, Me
End Sub

Private Sub mnuAnalysisAttackResponseItem_Click()
    frmAttackResponse.Show vbModeless, Me
End Sub

Private Sub mnuAnalysisAttackVisualizingItem_Click()
    frmAttackVisualizing.Show vbModeless, Me
End Sub

Private Sub mnuAnalysisLogsItem_Click()
    frmLog.Show vbModeless, frmMain
End Sub

Private Sub mnuConfigurationPreferencesItem_Click()
    frmConfiguration.Show vbModal
End Sub

Private Sub mnuConfigurationToolbarItem_Click()
    tlbMenu.Customize
End Sub

Private Sub mnuHelpIndexItem_Click()
    Call OpenOnlineHelp
End Sub

Private Sub mnuPluginsDeleteItem_Click()
    'Delete the selected plugin if there is one available
    If tvwPlugins.Nodes.Count Then
        On Error Resume Next
        If tvwPlugins.SelectedItem.Selected = True Then
            tvwPlugins.Nodes.Remove (tvwPlugins.SelectedItem.Index)
        End If
    End If
    
    'Actualisize the new plugin count
    fraPlugins.Caption = "Plugins (" & HowManyLoadedPlugins & " loaded)"
End Sub

Public Sub mnuPluginsDownloadTheLatestPluginsItem_Click()
    Dim intAlreadyAvailablePlugins As Integer
    Dim intNewAvailablePlugins As Integer
    
    Call FreezeWindows
    
    intAlreadyAvailablePlugins = filATKPlugins.ListCount
    'Load the latest plugin repository from the project web site
    frmPluginAutoUpdate.Show vbModal
    filATKPlugins.Refresh
    intNewAvailablePlugins = filATKPlugins.ListCount
    
    If intAlreadyAvailablePlugins <> intNewAvailablePlugins Then
        Call mnuPluginsReloadAllItem_Click
    End If
    
    Call FreeWindows
End Sub

Private Sub mnuPluginsEditItem_Click()
    frmAttackEditor.Show vbModeless, Me
End Sub

Private Sub mnuPluginsExportLoadedPluginListItem_Click()
    Call FreezeWindows
    Call ExportPluginsToHTMLFile
    Call FreeWindows
End Sub

Private Sub mnuPluginsExternalEditorItem_Click()
    Dim strPluginsExternalEditor As String  'Here we save the name of the external editor
    
    'Presave the given value for an external editor
    strPluginsExternalEditor = application_plugin_external_editor
    
    'Define the external editor
    If LenB(application_plugin_external_editor) Then
        If application_plugin_external_editor = "notepad.exe" Then
            strPluginsExternalEditor = application_plugin_external_editor
        ElseIf application_plugin_external_editor = "wordpad.exe" Then
            strPluginsExternalEditor = application_plugin_external_editor
        ElseIf (Dir$(application_plugin_external_editor, 16) <> "") Then
            strPluginsExternalEditor = application_plugin_external_editor
        Else
            Call errPluginExternalEditorMissing
        End If
    Else
        If InStr(1, plugin_filename, ".nasl", vbBinaryCompare) Then
            strPluginsExternalEditor = "wordpad.exe"
        Else
            strPluginsExternalEditor = "notepad.exe"
        End If
    End If

    'Open the editor
    Call ShellExecute(Me.hwnd, "Open", strPluginsExternalEditor, ChrW$(34) & plugin_filename & Chr(34), application_plugin_directory, 1)
End Sub

Private Sub mnuPluginsFindNextItem_Click()
    Call SearchPlugin
End Sub

Public Sub mnuPluginsReloadAllItem_Click()
    Screen.MousePointer = 13
    'Delete the whole plugins list
    WriteLogEntry "Unload the loaded plugins ...", 6
    Call NotExpanded
    tvwPlugins.Nodes.Clear
    
    'Refresh the plugins directory listing. So we can surely detect new files
    'in the plugin directory. And the reload the plugins
    filATKPlugins.Refresh
    Call LoadATKPlugins
    
    filNASLPlugins.Refresh
    Call LoadNASLPlugins
    
    Screen.MousePointer = 0
End Sub

Private Sub mnuPluginsReloadItem_Click()
    If tvwPlugins.Nodes.Count Then
        On Error Resume Next
        If tvwPlugins.SelectedItem.Child.Selected = True Then
            Call tvwPlugins_NodeClick(tvwPlugins.SelectedItem)
        End If
    Else
        Call mnuPluginsReloadAllItem_Click
    End If
End Sub

Private Sub mnuPluginsReportConfigurationItem_Click()
    frmReportConfiguration.Show vbModal
End Sub

Private Sub mnuPluginsRunDetectionItem_Click()
    session_procedure_type = "detection"
    Call ValidatePluginInput
End Sub

Private Sub mnuPluginsRunExploitItem_Click()
    session_procedure_type = "exploit"
    Call ValidatePluginInput
End Sub

Private Sub mnuPluginsSearchPluginItem_Click()
    'Define a default search string if this is the first search
    If LenB(strPluginSearchText) = 0 Then
        strPluginSearchText = "Microsoft"
    End If
    
    'Ask for the search string
    strPluginSearchText = InputBox("Please enter string you are searching for. " & _
        "(e.g. Microsoft, Apache, Sendmail).", _
        "Attack Tool Kit plugin search", strPluginSearchText)
    
    'Start the search
    Call SearchPlugin
End Sub

Private Sub mnuPluginsShowOnlinePluginItem_Click()
    If LenB(plugin_filename) Then
        Dim WebSiteURL As String
        
        WebSiteURL = application_plugin_download_url & plugin_filename & ".html"
        
        'Load the project web site
        WriteLogEntry "Loading the plugin website " & WebSiteURL, 6
        Call ShellExecute(Me.hwnd, "Open", WebSiteURL, "", App.Path, 1)
    End If
End Sub

Private Sub mnuReportingConfigurationItem_Click()
    frmReportConfiguration.Show vbModeless, Me
End Sub

Private Sub mnuReportingShowReportItem_Click()
    frmReport.Show vbModeless, Me
End Sub

Private Sub mnuScanStartItem_Click()
    Call ValidatePluginInput
End Sub

Private Sub mnuScanStopItem_Click()
    Call StopAttack
End Sub

Public Sub SetProgress(ByRef iValue As Integer)
    'Prevent too large values (this is just a nasty workaround!)
    If iValue > 100 Then
        iValue = 100
    End If
    
    frmMain.StatusBar.Panels(2).Text = iValue & " %"
    pbrProgress.Value = iValue
End Sub

Private Sub timTimeout_Timer()
    WriteLogEntry "Attack timed out after " & _
        timTimeout.Interval & " milliseconds. Ready.", 5
    
    'Close the socket
    Call wskTCPWinsock(0).Close

    'Reset the progress bar if there is a single check
    If application_attack_mode = "SingleCheck" Then
        SetProgress (100)
    End If

    Call FreeWindows
End Sub

Private Sub StopAttack()
    'Abord the check; well, just closing the open socket.
    WriteLogEntry "Abording check ...", 6
    Call wskTCPWinsock(0).Close
    Call FreeWindows
    SetProgress 100
    WriteLogEntry "Check aborded. Ready.", 6
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Select the toolbar button and call the needed sub routine
    Select Case Button.Caption
        Case "Start" 'Start
            Call mnuScanStartItem_Click
        Case "Stop" 'Stop
            Call StopAttack
        Case "Config" 'Config
            Call mnuConfigurationPreferencesItem_Click
        Case "Edit" 'Edit
            Call mnuPluginsEditItem_Click
        Case "Reload" 'Reload
            Call mnuPluginsReloadItem_Click
        Case "Delete" 'Delete
            Call mnuPluginsDeleteItem_Click
        Case "Visualize" 'Visualize
            Call mnuAnalysisAttackVisualizingItem_Click
        Case "Response" 'Response
            Call mnuAnalysisAttackResponseItem_Click
        Case "Logs" 'Logs
            Call mnuAnalysisLogsItem_Click
        Case "Report" 'Logs
            Call mnuReportingShowReportItem_Click
    End Select
End Sub

Private Sub tlbMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuConfiguration
    End If
End Sub

Private Sub tvwPlugins_DblClick()
    Call OpenAttackEditor
End Sub

Private Sub tvwPlugins_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OpenAttackEditor
    End If
End Sub

Private Sub OpenAttackEditor()
    On Error Resume Next 'Prevent errors if no node is selected
    If tvwPlugins.SelectedItem.Child.Selected = True Then
        Dim strSelectedItemKey As String    'Here we save the key name of the selected node

        'Prepare the key to grant faster access
        strSelectedItemKey = tvwPlugins.SelectedItem.Key
        
        'Check if the selected node is really a plugin
        If InStr(Len(strSelectedItemKey) - 6, strSelectedItemKey, ".plugin", vbBinaryCompare) Then
            'Load the attack editor for small modifications
            frmAttackEditor.Show vbModeless, Me
        ElseIf InStr(Len(strSelectedItemKey) - 4, strSelectedItemKey, ".nasl", vbBinaryCompare) Then
            'Load the attack editor for small modifications
            frmAttackEditor.Show vbModeless, Me
        End If
    End If
End Sub

Private Sub tvwPlugins_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim oNode As Node
    
    Set oNode = tvwPlugins.HitTest(x, y)
    
    If Not oNode Is Nothing Then
        If Button = 2 Then
            On Error Resume Next 'Prevent errors if no node is selected
            If tvwPlugins.SelectedItem.Child.Selected = True Then
                Dim strSelectedItemKey As String    'Here we save the key name of the selected node
        
                'Prepare the key to grant faster access
                strSelectedItemKey = tvwPlugins.SelectedItem.Key
                
                'Check if the selected node is really a plugin
                If InStr(Len(strSelectedItemKey) - 6, strSelectedItemKey, ".plugin", vbBinaryCompare) Then
                    'Show context menu if 2nd mouse button is pressed
                    PopupMenu mnuPlugins
                ElseIf InStr(Len(strSelectedItemKey) - 5, strSelectedItemKey, ".nasl", vbBinaryCompare) Then
                    'Show context menu if 2nd mouse button is pressed
                    PopupMenu mnuPlugins
                End If
            End If
        End If
    End If
End Sub

Private Sub tvwPlugins_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim SelectedKey As String   'Here we save the selected key

    'Detect unsaved data in the attack editor
    If IsFormVisible("frmAttackEditor") = True Then
        Call frmAttackEditor.CheckIfPluginIsEdited
    End If
    
    If HowManyLoadedPlugins <> 0 Then
        'Center the view
        tvwPlugins.SelectedItem.EnsureVisible
        
        'Save the selected key
        SelectedKey = tvwPlugins.SelectedItem.Key
        
        'Check if the selected key is a plugin filename
        If InStr(1, SelectedKey, ".plugin") Then
            'Strip and cache the filename
            plugin_filename = Mid$(SelectedKey, 2, Len(SelectedKey))
            
            'Read the selected plugin
            Call ParseATKPlugin(ReadPluginFromFile(plugin_filename, application_plugin_directory))
            Call PrepareTheNewPluginData
        ElseIf InStr(1, SelectedKey, ".nasl") Then
            'Strip and cache the filename
            plugin_filename = Mid$(SelectedKey, 2, Len(SelectedKey))
            
            'Read the selected plugin
            Call ParseNASLPlugin(ReadPluginFromFile(plugin_filename, application_plugin_directory))
            Call PrepareTheNewPluginData
        End If
    End If
End Sub

Private Sub PrepareTheNewPluginData()
    WriteLogEntry "Reading plugin " & plugin_id & " (" & plugin_filename & ")...", 6

    'Reset the last response
    LastResponse = vbNullString
    
    'Show the plugin content
    txtPluginContent.Text = Replace(GenerateTXTReportPluginEntry(False, vbNullString), "     ", vbNullString, , , vbBinaryCompare)
    'txtPluginContent.Text = GenerateTXTReportPluginEntry(False, vbNullString)
    
    'Write the data into the attack editor if he is visible
    'If it is not visible he'll do it self on load
    If IsFormVisible("frmAttackEditor") Then
        'load the actual values
        Call frmAttackEditor.LoadActualValues
    End If

    If IsFormVisible("frmReportConfiguration") Then
        'load the actual values
        Call frmReportConfiguration.RefreshReportStructure
    End If
End Sub

Private Sub txtPluginContent_DblClick()
    Call OpenSelectedTextIfItIsURL(txtPluginContent.SelText)
End Sub

Private Sub wskTCPWinsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Here is the incoming data cached
    Dim DataStr As String
            
    'Read the incoming data and write it to DataStr$
    Call wskTCPWinsock(0).GetData(DataStr$, vbString)
    
    'Update the status bar
    WriteLogEntry "Receiving data """ & Mid$(DataStr, 1, 64) & """ from the target ...", 6
    
    If LenB(LastResponse) < 16000 Then
        LastResponse = LastResponse & DataStr
        LastResponseTime = GetActualTime(":")
    Else
        wskTCPWinsock(0).Close
    End If

    Call LoadLatestResponse

    If IsFormVisible("frmAttackVisualizing") = True Then
        frmAttackVisualizing.VisualizeDataArrival
    End If
End Sub

Private Sub wskTCPWinsock_Close(Index As Integer)
    'Write the response to a file
    Call WriteLastResponseToFile
    
    'Update the status bar
    WriteLogEntry "Closing socket ...", 6
    
    'Disable the timer because a time out makes no sense anymore
    timTimeout.Enabled = False
    
    'Close and free the socket
    wskTCPWinsock(0).Close
End Sub

Private Sub wskTCPWinsock_Error(Index As Integer, ByVal Number As Integer, _
    Description As String, ByVal Scode As Long, ByVal Source As String, _
    ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLogEntry "WinSock Error: [" & Number & "] " & Description, 1
    
    Call wskTCPWinsock_Close(0)
End Sub

Public Sub Pause(lngDuration As Long)
    'Sleep function for connection attempts and other stuff.
    Dim lngCurrent As Long
    
    lngCurrent = Timer
    Do Until Timer - lngCurrent >= lngDuration
        DoEvents
    Loop
End Sub

Private Sub mnuExitItem_Click()
    End
End Sub

Private Sub mnuAboutItem_Click()
    frmAbout.Show vbModeless, Me
End Sub
Private Sub mnuICMPPingItem_Click()
    frmICMPPing.Visible = True
End Sub

Private Sub mnuNslookupItem_Click()
    frmNslookup.Visible = True
End Sub

Private Sub mnuPortscannerItem_Click()
    frmPortscanner.Visible = True
End Sub

Private Sub mnuProjectWebSiteItem_Click()
    Call OpenProjectWebsite
End Sub

Private Sub NotExpanded()
    Dim mNode   As Node
    
    With tvwPlugins
        For Each mNode In .Nodes
            If mNode.Expanded Then mNode.Expanded = False
        Next
    End With
End Sub

Private Sub SearchPlugin()
    Dim intListItemStartPosition As Integer
    Dim intListItemCount As Integer
    Dim i As Integer
    
    WriteLogEntry "Starting the local plugin search for the string """ & strPluginSearchText & """ ...", 6
    
    intListItemCount = Me.tvwPlugins.Nodes.Count
    
    If tvwPlugins.SelectedItem.Index < intListItemCount Then
        intListItemStartPosition = tvwPlugins.SelectedItem.Index + 1
    Else
        intListItemStartPosition = 0
    End If
    
    For i = intListItemStartPosition To intListItemCount
        If InStr(1, _
            LCase$(tvwPlugins.Nodes.Item(i).Text), _
            LCase$(strPluginSearchText), vbBinaryCompare) Then
            
            Set tvwPlugins.SelectedItem = tvwPlugins.Nodes.Item(i)
            Call tvwPlugins_NodeClick(tvwPlugins.SelectedItem)
            Exit For
        End If
    Next i
End Sub
