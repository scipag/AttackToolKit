VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   4260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7125
   Icon            =   "frmConfiguration.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7125
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLogs 
      Caption         =   "Logs"
      Height          =   3375
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtLogsDirectory 
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   2040
         Width           =   5175
      End
      Begin VB.CommandButton cmdBrowseLogsDirectory 
         Caption         =   "Browse"
         Height          =   255
         Left            =   4920
         TabIndex        =   34
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmdDefaultLogsDirectory 
         Caption         =   "Default"
         Height          =   255
         Left            =   5760
         TabIndex        =   35
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox cmbLogsSecurityLevel 
         Height          =   315
         ItemData        =   "frmConfiguration.frx":0CCA
         Left            =   1320
         List            =   "frmConfiguration.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1560
         Width           =   5175
      End
      Begin VB.CheckBox chkActivateLogs 
         Caption         =   "Activate lo&gs"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Value           =   1  'Checked
         Width           =   6255
      End
      Begin VB.Label lblLabel 
         Caption         =   "Security Level"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   "Logs directory"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":0CCE
         Height          =   615
         Index           =   15
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   6015
      End
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reporting"
      Height          =   3375
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkReportOpenAfterGeneration 
         Caption         =   "Open a report with the default application after report generation"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.TextBox txtReportsDirectory 
         Height          =   285
         Left            =   1560
         TabIndex        =   59
         Top             =   360
         Width           =   4935
      End
      Begin VB.CommandButton cmdBrowseReportsDirectory 
         Caption         =   "Browse"
         Height          =   255
         Left            =   4920
         TabIndex        =   60
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdDefaultReportsDirectory 
         Caption         =   "Default"
         Height          =   255
         Left            =   5760
         TabIndex        =   61
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblReportTemplateNote 
         Caption         =   "Editing of the report templates can be done in the report configuration."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         MouseIcon       =   "frmConfiguration.frx":0DA9
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Label lblLabel 
         Caption         =   "Reports Directory"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraSearchengine 
      Caption         =   "Searchengine"
      Height          =   3375
      Left            =   240
      TabIndex        =   46
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbSearchEngineURL 
         Height          =   315
         Left            =   240
         TabIndex        =   47
         Text            =   "http://www.google.com/search?q="
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblSearchEngineTest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test the selected search engine query url by searching for the ATK project."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   645
         MouseIcon       =   "frmConfiguration.frx":10B3
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   1440
         Width           =   5325
      End
      Begin VB.Label lblLabel 
         Caption         =   "Search engine default query string for online searches"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame fraSuggestions 
      Caption         =   "Suggestions"
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtSuggestionsDirectory 
         Height          =   285
         Left            =   2040
         TabIndex        =   74
         Top             =   840
         Width           =   4455
      End
      Begin VB.CommandButton cmdBrowseSuggestionsDirectory 
         Caption         =   "Browse"
         Height          =   255
         Left            =   4920
         TabIndex        =   75
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdDefaultSuggestionsDirectory 
         Caption         =   "Default"
         Height          =   255
         Left            =   5760
         TabIndex        =   76
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkSuggestions 
         Caption         =   "&Activate suggestions"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         Caption         =   "Suggestions Directory"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame fraResponses 
      Caption         =   "Responses"
      Height          =   3375
      Left            =   240
      TabIndex        =   68
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtResponsesDirectory 
         Height          =   285
         Left            =   2040
         TabIndex        =   69
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmdBrowseResponsesDirectory 
         Caption         =   "Browse"
         Height          =   255
         Left            =   4920
         TabIndex        =   71
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdDefaultResponsesDirectory 
         Caption         =   "Default"
         Height          =   255
         Left            =   5760
         TabIndex        =   72
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblLabel 
         Caption         =   "Responses Directory"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraPlugins 
      Caption         =   "Plugins"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbPluginsExternalEditor 
         Height          =   315
         Left            =   1560
         TabIndex        =   58
         Text            =   "notepad.exe"
         Top             =   2760
         Width           =   4935
      End
      Begin VB.CommandButton cmdPluginsDirectoryDefault 
         Caption         =   "Default"
         Height          =   255
         Left            =   5760
         TabIndex        =   54
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdPluginsDirectoryBrowse 
         Caption         =   "Browse"
         Height          =   255
         Left            =   4920
         TabIndex        =   53
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtPluginsDirectory 
         Height          =   285
         Left            =   1560
         TabIndex        =   52
         Top             =   360
         Width           =   4935
      End
      Begin VB.ComboBox cmbPluginsDownloadURL 
         Height          =   315
         Left            =   1560
         TabIndex        =   55
         Text            =   "http://www.computec.ch/projekte/atk/plugins/pluginslist/"
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtDefaultSleep 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   57
         Text            =   "3000"
         ToolTipText     =   "Default wait time for sleep command"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtTimeout 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   56
         Text            =   "30000"
         ToolTipText     =   "Timeout for the plugins"
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLabel 
         Caption         =   "External editor for plugins"
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   64
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "Plugins Download"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "(Default: 3000 = 3 seconds)"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   28
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblLabel 
         Caption         =   "Default wait value (ms) for sleep"
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "(Default: 30000 = 30 seconds)"
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label lblLabel 
         Caption         =   "Plugins Directory"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "Timeout (ms) for stucked plugins"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame fraPreferences 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame fraSafety 
         Caption         =   "Safety"
         Height          =   2175
         Left            =   0
         TabIndex        =   65
         Top             =   1200
         Width           =   6615
         Begin VB.CheckBox chkDoNoDoSChecks 
            Caption         =   "Do no Denial of Service checks"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkDoSilentChecks 
            Caption         =   "&Do silent checks"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.Label lblLabel 
            Caption         =   $"frmConfiguration.frx":13BD
            Height          =   495
            Index           =   4
            Left            =   480
            TabIndex        =   67
            Top             =   1560
            Width           =   6015
         End
         Begin VB.Label lblLabel 
            Caption         =   $"frmConfiguration.frx":1463
            Height          =   495
            Index           =   3
            Left            =   480
            TabIndex        =   66
            Top             =   720
            Width           =   6015
         End
      End
      Begin VB.Frame fraMode 
         Caption         =   "Mode"
         Height          =   1095
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   6615
         Begin VB.OptionButton optSingleCheck 
            Caption         =   "&Single Check"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Only check specific potential flaws on demand."
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optFullAudit 
            Caption         =   "&Full Audit"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Check the target for all possible potential flaws."
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Only check specific potential flaws on demand."
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   21
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label lblLabel 
            Caption         =   "Check the target for all possible potential flaws."
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   20
            Top             =   720
            Width           =   4455
         End
      End
   End
   Begin VB.Frame fraAlerting 
      Caption         =   "Alerting"
      Height          =   3375
      Left            =   240
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkAlertingVulnerabilityNotFound 
         Caption         =   "Produce alert when vulnerbility is not found."
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   3495
      End
      Begin VB.CheckBox chkAlertingVulnerabilityFound 
         Caption         =   "Produce alert when vulnerability is found."
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame fraHelp 
      Caption         =   "Help"
      Height          =   3375
      Left            =   240
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.ComboBox cmbHelpURL 
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         Text            =   "http://www.computec.ch/projekte/atk/documentation/help/"
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label lblOnlineHelpURLTest 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test the selected online help url."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2145
         MouseIcon       =   "frmConfiguration.frx":1509
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   1920
         Width           =   2325
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":1813
         Height          =   615
         Index           =   19
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label lblLabel 
         Caption         =   "Help URL"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame fraSpeech 
      Caption         =   "Speech"
      Height          =   3375
      Left            =   240
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CheckBox chkActivateSpeech 
         Caption         =   "Activate Speech"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label lblTestSpeechFeature 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test the speech feature."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         MouseIcon       =   "frmConfiguration.frx":18CE
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label lblLabel 
         Caption         =   $"frmConfiguration.frx":1BD8
         Height          =   615
         Index           =   18
         Left            =   480
         TabIndex        =   30
         Top             =   720
         Width           =   5895
      End
   End
   Begin VB.Frame fraMapping 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame fraICMPMapping 
         Caption         =   "ICMP Mapping"
         Height          =   3375
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   6615
         Begin VB.CheckBox chkDoICMPMapping 
            Caption         =   "Do &ICMP mapping (ICMP echo request)"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.CheckBox chkScanifICMPfails 
            Caption         =   "Scan if ICMP mapping fails"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   720
            Width           =   2295
         End
      End
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   6615
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   120
         MaxLength       =   200
         TabIndex        =   1
         Text            =   "localhost"
         ToolTipText     =   "Host name or IP address of the target"
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   "Warning: You should never scan a network ressource without permission."
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Top             =   1560
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog cdgSaveAs 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save file"
      FileName        =   "newdefault.config"
      Filter          =   "ATK Configuration (*.config)|*.config|INI Files (*.ini)|*.ini|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open file"
      FileName        =   "default.config"
      Filter          =   "ATK Configuration (*.config)|*.config|INI Files (*.ini)|*.ini|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.TabStrip tspConfiguration 
      Height          =   3975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   12
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Target"
            Object.ToolTipText     =   "Selection of the target"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "P&references"
            Object.ToolTipText     =   "Preferences for program behavior"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mapping"
            Object.ToolTipText     =   "Settings for mapping (e.g. target deteciton)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Plugins"
            Object.ToolTipText     =   "Plugin directory and plugin behavior"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Alerting"
            Object.ToolTipText     =   "Settings for the alerting"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Responses"
            Object.ToolTipText     =   "Responses directory and responses behavior"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Suggestions"
            Object.ToolTipText     =   "Suggestions directory and suggestions behavior"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search&engine"
            Object.ToolTipText     =   "Searchengine settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reporting"
            Object.ToolTipText     =   "Reports directory and report behavior"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Logs"
            Object.ToolTipText     =   "Logs directory and logs behavior"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Speech"
            Object.ToolTipText     =   "Speech settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            Object.ToolTipText     =   "Help settings"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSpererator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveItem 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAsItem 
         Caption         =   "Save As ..."
      End
      Begin VB.Menu mnuFileSpererator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpConfigurationHelpItem 
         Caption         =   "&Configuration Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.1 2005-01-16                                                           *
' * - Fixed the tab select order.                                                    *
' * Version 4.0 2004-12-12                                                           *
' * - Added a tagging routine and editing check as like in the Attack Editor.        *
' * - Added the ToolTipText for the TabStrip sheets.                                 *
' * Version 4.0 2004-12-07                                                           *
' * - Added the elements and procedures for handling the external plugin editor.     *
' * Version 4.0 2004-12-06                                                           *
' * - Replaced the DirListBoxes with simple TextBoxes.                               *
' * - Added an error checking routine for the existence of selected directories.     *
' * - Added a browse and default button for all the directory TextBoxes.             *
' * - Re-activated the plugin reload procedure if a new plugin directory has been    *
' *   specified.                                                                     *
' * Version 3.0 2004-11-05                                                           *
' * - Added an error routine withing the save as function if the cancel button is    *
' *   pressed.                                                                       *
' * Version 3.0 2004-11-03                                                           *
' * - Fixed a bug if opening the report configuration.                               *
' * Version 3.0 2004-11-01                                                           *
' * - Fixed the File/New function. It should work now.                               *
' * - Deleted all not needed nor supported elements.                                 *
' * Version 3.0 2004-10-30                                                           *
' * - Fully enhanced and re-sorted the configuration file output. We are now using a *
' *   Unix/Linux conf file format that allows commenting out lines by using the #    *
' * Version 3.0 2004-10-28                                                           *
' * - Added the menu save as function to save specific/other configuration files.    *
' * Version 3.0 2004-10-23                                                           *
' * - Added the menu file open function to open specific/other configuration files.  *
' * Version 3.0 2004-10-22                                                           *
' * - Added the error message behavior if the target specifying is wrong.            *
' * Version 3.0 2004-10-20                                                           *
' * - Added the tab and routines for the online help configuration.                  *
' * Version 3.0 2004-10-08                                                           *
' * - Enhanced and bugfixed the whole logging.                                       *
' * - Added the update features for the AutoUpdate.                                  *
' * Version 2.1 2004-09-08                                                           *
' * - Corrected and enhanced the full audit mode warning.                            *
' * Version 2.0 2004-04-08                                                           *
' * - Added the actualizing of the target data in frmAttackVisualizing after         *
' *   clicking accept.                                                               *
' * Version 1.1 2004-03-20                                                           *
' * - Added the configuration file name in the frame caption for more verbosity.     *
' * - Added a warning message if the full audit mode is selected.                    *
' ************************************************************************************

Private bolConfigurationIsEdited As Boolean

Private Sub chkActivateLogs_Click()
    If chkActivateLogs.Value = 1 Then
        cmbLogsSecurityLevel.Enabled = True
        txtLogsDirectory.Enabled = True
        cmdBrowseLogsDirectory.Enabled = True
        cmdDefaultLogsDirectory.Enabled = True
    Else
        cmbLogsSecurityLevel.Enabled = False
        txtLogsDirectory.Enabled = False
        cmdBrowseLogsDirectory.Enabled = False
        cmdDefaultLogsDirectory.Enabled = False
    End If
End Sub

Private Sub chkActivateLogs_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkActivateSpeech_Click()
    Call TagConfigAsEdited
End Sub

Private Sub chkAlertingVulnerabilityFound_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkAlertingVulnerabilityNotFound_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkDoICMPMapping_Click()
    If chkDoICMPMapping.Value = 0 Then
        chkScanifICMPfails.Enabled = False
    Else
        chkScanifICMPfails.Enabled = True
    End If
End Sub

Private Sub chkDoICMPMapping_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkDoNoDoSChecks_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkDoSilentChecks_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkReportOpenAfterGeneration_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkScanifICMPfails_Click()
    If chkScanifICMPfails.Value = 0 Then
        chkScanifICMPfails.Enabled = False
    Else
        chkScanifICMPfails.Enabled = True
    End If
End Sub

Private Sub chkScanifICMPfails_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub chkSuggestions_Click()
    If chkSuggestions.Value <> 1 Then
        txtSuggestionsDirectory.Enabled = False
        cmdBrowseSuggestionsDirectory.Enabled = False
        cmdDefaultSuggestionsDirectory.Enabled = False
    Else
        txtSuggestionsDirectory.Enabled = True
        cmdBrowseSuggestionsDirectory.Enabled = True
        cmdDefaultSuggestionsDirectory.Enabled = True
    End If
End Sub

Private Sub chkSuggestions_LostFocus()
    Call TagConfigAsEdited
End Sub

Private Sub cmbHelpURL_LostFocus()
    Call DetectConfigAltering("help url", application_help_url, cmbHelpURL.Text)
End Sub

Private Sub cmbLogsSecurityLevel_LostFocus()
    Call DetectConfigAltering("logs security level", CStr(application_log_security_level), cmbLogsSecurityLevel.ListIndex)
End Sub

Private Sub cmbPluginsDownloadURL_LostFocus()
    Call DetectConfigAltering("plugins download url", application_plugin_download_url, cmbPluginsDownloadURL.Text)
End Sub

Private Sub cmbPluginsExternalEditor_LostFocus()
    Call DetectConfigAltering("default external plugins editor", application_plugin_external_editor, cmbPluginsExternalEditor.Text)
End Sub

Private Sub cmbSearchEngineURL_KeyPress(KeyAscii As Integer)
    'Complete a combobox writing
    Static iLeftOff As Long
    ComboAutoComplete cmbSearchEngineURL, KeyAscii, iLeftOff
End Sub

Private Sub cmbSearchEngineURL_LostFocus()
    Call DetectConfigAltering("default search engine url", application_searchengine_url, cmbSearchEngineURL.Text)
End Sub

Private Sub cmdBrowseLogsDirectory_Click()
    Dim strDirectory As String
  
    strDirectory = BrowseForFolder(Me)
  
    If LenB(strDirectory) Then
        txtLogsDirectory.Text = strDirectory
    Else
        txtLogsDirectory.Text = application_log_directory
    End If

    Call DetectConfigAltering("logs directory", application_log_directory, txtLogsDirectory.Text)
End Sub

Private Sub cmdBrowseReportsDirectory_Click()
    Dim strDirectory As String
  
    strDirectory = BrowseForFolder(Me)
  
    If LenB(strDirectory) Then
        txtReportsDirectory.Text = strDirectory
    Else
        txtReportsDirectory.Text = application_report_directory
    End If
    
    Call DetectConfigAltering("reports directory", application_report_directory, txtReportsDirectory.Text)
End Sub

Private Sub cmdBrowseResponsesDirectory_Click()
    Dim strDirectory As String
  
    strDirectory = BrowseForFolder(Me)
  
    If LenB(strDirectory) Then
        txtResponsesDirectory.Text = strDirectory
    Else
        txtResponsesDirectory.Text = application_response_directory
    End If
    
    Call DetectConfigAltering("responses directory", application_response_directory, txtResponsesDirectory.Text)
End Sub

Private Sub cmdBrowseSuggestionsDirectory_Click()
    Dim strDirectory As String
  
    strDirectory = BrowseForFolder(Me)
  
    If LenB(strDirectory) Then
        txtSuggestionsDirectory.Text = strDirectory
    Else
        txtSuggestionsDirectory.Text = application_suggestion_directory
    End If
    
    Call DetectConfigAltering("suggestions directory", application_suggestion_directory, txtSuggestionsDirectory.Text)
End Sub

Private Sub cmdDefaultLogsDirectory_Click()
    Dim strDefaultLogsDirectory As String
        
    strDefaultLogsDirectory = App.Path & "\logs"
        
    txtLogsDirectory.Text = strDefaultLogsDirectory

    Call DetectConfigAltering("logs directory", application_log_directory, strDefaultLogsDirectory)
End Sub

Private Sub cmdDefaultReportsDirectory_Click()
    Dim strDefaultReportsDirectory As String
    
    strDefaultReportsDirectory = App.Path & "\reports"
    
    txtReportsDirectory.Text = strDefaultReportsDirectory
    
    Call DetectConfigAltering("reports directory", application_report_directory, strDefaultReportsDirectory)
End Sub

Private Sub cmdDefaultResponsesDirectory_Click()
    Dim strDefaultResponsesDirectory As String
    
    strDefaultResponsesDirectory = App.Path & "\responses"
    
    txtResponsesDirectory.Text = strDefaultResponsesDirectory
    
    Call DetectConfigAltering("responses directory", application_response_directory, strDefaultResponsesDirectory)
End Sub

Private Sub cmdDefaultSuggestionsDirectory_Click()
    Dim strDefaultSuggestionsDirectory As String
    
    strDefaultSuggestionsDirectory = App.Path & "\suggestions"
    
    txtSuggestionsDirectory.Text = strDefaultSuggestionsDirectory
    
    Call DetectConfigAltering("suggestions directory", application_suggestion_directory, strDefaultSuggestionsDirectory)
End Sub

Private Sub cmdPluginsDirectoryBrowse_Click()
    Dim strDirectory As String
  
    strDirectory = BrowseForFolder(Me)
  
    If LenB(strDirectory) Then
        txtPluginsDirectory.Text = strDirectory
    Else
        txtPluginsDirectory.Text = application_plugin_directory
    End If
    
    Call DetectConfigAltering("plugins directory", application_plugin_directory, txtPluginsDirectory.Text)
End Sub

Private Sub cmdPluginsDirectoryDefault_Click()
    Dim strDefaultPluginPath As String
    
    strDefaultPluginPath = App.Path & "\plugins"
    
    txtPluginsDirectory.Text = strDefaultPluginPath
    
    Call DetectConfigAltering("plugins directory", application_plugin_directory, strDefaultPluginPath)
End Sub

Private Sub Form_Activate()
    If Me.WindowState = vbMinimized Then
        Me.WindowState = vbNormal
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 27
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    'Show the frame title with configuration file name
    frmConfiguration.Caption = "Configuration - " & application_configuration_filename
      
    'Add some default values
    cmbLogsSecurityLevel.AddItem "0 emergencies (A panic condition if the system is unusable.)", 0
    cmbLogsSecurityLevel.AddItem "1 alerts (A condition that should be corrected immediately.)", 1
    cmbLogsSecurityLevel.AddItem "2 critical (Critical conditions, e.g. hard device errors.)", 2
    cmbLogsSecurityLevel.AddItem "3 error (Errors)", 3
    cmbLogsSecurityLevel.AddItem "4 warnings (Warning messages)", 4
    cmbLogsSecurityLevel.AddItem "5 notifications (Conditions that are not error conditions, but should possibly be handled specially.)", 5
    cmbLogsSecurityLevel.AddItem "6 informational (Informational messages)", 6
    cmbLogsSecurityLevel.AddItem "7 debugging (Messages that contain information normally of use only when debugging a program.)", 7
    
    cmbPluginsDownloadURL.AddItem "http://www.computec.ch/projekte/atk/plugins/pluginslist/"
    
    cmbPluginsExternalEditor.AddItem "notepad.exe"
    cmbPluginsExternalEditor.AddItem "wordpad.exe"
    
    'Add the search engine query urls
    cmbSearchEngineURL.AddItem "http://www.google.com/search?q="
    cmbSearchEngineURL.AddItem "http://search.yahoo.com/search?p="
    cmbSearchEngineURL.AddItem "http://www.hotbot.com/default.asp?query="
    cmbSearchEngineURL.AddItem "http://www.altavista.com/web/results?q="
    cmbSearchEngineURL.AddItem "http://www.alltheweb.com/search?q="
    cmbSearchEngineURL.AddItem "http://search.netscape.com/ns/search?query="
    cmbSearchEngineURL.AddItem "http://a9.com/"
    cmbSearchEngineURL.AddItem "http://search.msn.com/results.aspx?q="
    cmbSearchEngineURL.AddItem "http://msxml.excite.com/info.xcite/search/web/"
    cmbSearchEngineURL.AddItem "http://suche.fireball.de/cgi-bin/pursuit?query="
    cmbSearchEngineURL.AddItem "http://suche.lycos.de/cgi-bin/pursuit?query="
    cmbSearchEngineURL.AddItem "http://search.megaspider.com/XP.html?"
    cmbSearchEngineURL.AddItem "http://web.ask.com/web?q="
    cmbSearchEngineURL.AddItem "http://search.dmoz.org/cgi-bin/search?search="
    cmbSearchEngineURL.AddItem "http://astalavista.box.sk/cgi-bin/robot?srch="
    cmbSearchEngineURL.AddItem "http://www2.packetstormsecurity.org/cgi-bin/search/search.cgi?searchvalue="
    cmbSearchEngineURL.AddItem "http://astalavista.box.sk/cgi-bin/robot?srch="
    cmbSearchEngineURL.AddItem "http://search.gulli.com/"
    cmbSearchEngineURL.AddItem "http://www.gurunet.com/query?s="
    cmbSearchEngineURL.AddItem "http://froogle.google.com/froogle?q="
    cmbSearchEngineURL.AddItem "http://anon.free.anonymizer.com/http://www.google.com/search?q="

    cmbHelpURL.AddItem "http://www.computec.ch/projekte/atk/documentation/help/"

    'Show the configuration
    Call LoadActualConfigurationValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If CheckIfConfigIsEdited = True Then
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    'Check the window state. Do not resize if the window is minimized
    If Me.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If Me.Height <> 4920 Then
            Me.Height = 4920
        End If
        
        'Prevent zu small windows in width
        If Me.Width <> 7215 Then
            Me.Width = 7215
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConfiguration = Nothing
End Sub

Private Sub lblOnlineHelpURLTest_Click()
    Dim strWebSiteURL As String
    
    strWebSiteURL = cmbHelpURL.Text
    
    'Load the online help
    WriteLogEntry "Loading the online help " & strWebSiteURL & " as a test ...", 6
    Call ShellExecute(Me.hwnd, "Open", strWebSiteURL, "", App.Path, 1)
End Sub

Private Sub lblReportTemplateNote_Click()
    'We are loading the form modal. This is not very nice but it is better than
    'another run-time error.
    frmReportConfiguration.Show vbModal
End Sub

Private Sub lblSearchEngineTest_Click()
    Dim strSearchEngineTestURL As String
    
    strSearchEngineTestURL = cmbSearchEngineURL.Text & "Attack Tool Kit"
    
    WriteLogEntry "Opening the search engine URL " & strSearchEngineTestURL & " for testing ...", 6
    Call ShellExecute(frmMain.hwnd, "Open", strSearchEngineTestURL, "", App.Path, 1)
End Sub

Private Sub lblTestSpeechFeature_Click()
    Dim bolActivateSpeechCache As Boolean
    
    If Not application_speech_enable Then
        application_speech_enable = True
    Else
        bolActivateSpeechCache = True
    End If
    
    Call ReadText("Test, check, check, one, two, ...")
    
    If bolActivateSpeechCache <> application_speech_enable Then
        application_speech_enable = bolActivateSpeechCache
    End If
End Sub

Private Sub mnuFileCloseItem_Click()
    Unload Me
End Sub

Private Sub mnuFileNewItem_Click()
    If CheckIfConfigIsEdited = False Then
        Call LoadConfigFromFile(App.Path & "\configs\127.0.0.1-" & GetTodaysDate(".") & ".config")
    
        'Show the "new" configuration
        Call LoadActualConfigurationValues
    End If
End Sub

Private Sub mnuFileOpenItem_Click()
    Dim strConfigurationFileName As String  'The name of the configuration file
    Dim strDefaultConfigurationPath As String
    
    If CheckIfConfigIsEdited = False Then
        strDefaultConfigurationPath = App.Path & "\configs"
        
        'Define the initial directory of the plugins
        On Error Resume Next
        If Not (Dir$(strDefaultConfigurationPath, 16) <> "") Then
            strDefaultConfigurationPath = App.Path
        End If
        
        cdgOpen.Filename = application_configuration_filename
        
        'Define the initial directory of the plugins
        cdgOpen.InitDir = strDefaultConfigurationPath
        
        'Ask the user for the desired filename
        cdgOpen.ShowOpen 'Opens the save dialog
        
        'Cache the filename into a variant to increase the speed
        strConfigurationFileName = cdgOpen.Filename
        
        'Check if a file was selected
        If LenB(strConfigurationFileName) Then
            'Check if the file exists
            If (Dir$(strConfigurationFileName, 16) <> "") Then
                'Load the configuration file
                Call LoadConfigFromFile(strConfigurationFileName)
                Call LoadActualConfigurationValues
            End If
        End If
    End If
End Sub

Private Sub mnuFileSaveAsItem_Click()
    Dim strDefaultConfigurationPath As String
    Dim strConfigurationFileName As String
    
    strDefaultConfigurationPath = App.Path & "\configs"
    
    'Define the initial directory of the plugins
    On Error Resume Next
    If Not (Dir$(strDefaultConfigurationPath, 16) <> "") Then
        strDefaultConfigurationPath = App.Path
    End If
    
    cdgSaveAs.InitDir = strDefaultConfigurationPath
    
    'Ask the user for the desired filename
    cdgSaveAs.Filename = application_configuration_filename
    On Error GoTo ErrSub
    cdgSaveAs.ShowSave 'Opens the save dialog
    strConfigurationFileName = cdgSaveAs.Filename 'Get the filename
    
    'Cut the plugin extension if there is one given
    If LenB(strConfigurationFileName) Then
        application_configuration_filename = strConfigurationFileName
        Call SaveConfigurationData
        Call WriteConfigurationToFile(application_configuration_filename)
    End If
ErrSub:
End Sub

Private Sub mnuFileSaveItem_Click()
    Call SaveConfigurationData
    Call WriteConfigurationToFile(application_configuration_filename)
End Sub

Private Sub mnuHelpConfigurationHelpItem_Click()
    Call OpenOnlineHelp("configuration")
End Sub

Private Sub optFullAudit_Click()
    MsgBox "The full audit mode is not the main feature of the ATK." & vbNewLine & _
        "The mode is not very efficient and a general audit task" & vbNewLine & _
        "can much better be done by other well-known security scanners." & vbNewLine & _
        "Please use the single check mode for checking dedicated" & vbNewLine & _
        "vulnerabilities (perhaps already identified by other scanners)" & vbNewLine & _
        "instead.", _
        vbInformation, "Attack Tool Kit full audit information"
End Sub

Private Sub optFullAudit_LostFocus()
    If optSingleCheck.Value = True Then
        Call DetectConfigAltering("Attack Mode", application_attack_mode, "SingleCheck")
    ElseIf optFullAudit.Value = True Then
        Call DetectConfigAltering("Attack Mode", application_attack_mode, "FullAudit")
    End If
End Sub

Private Sub optSingleCheck_LostFocus()
    If optSingleCheck.Value = True Then
        Call DetectConfigAltering("Attack Mode", application_attack_mode, "SingleCheck")
    ElseIf optFullAudit.Value = True Then
        Call DetectConfigAltering("Attack Mode", application_attack_mode, "FullAudit")
    End If
End Sub

Private Sub tspConfiguration_Click()
    'Target
    If tspConfiguration.SelectedItem.Index = 1 Then
        fraTarget.Visible = True
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Preferences
    ElseIf tspConfiguration.SelectedItem.Index = 2 Then
        fraPreferences.Visible = True
        fraTarget.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Mapping
    ElseIf tspConfiguration.SelectedItem.Index = 3 Then
        fraMapping.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Plugins
    ElseIf tspConfiguration.SelectedItem.Index = 4 Then
        fraPlugins.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Alerting
    ElseIf tspConfiguration.SelectedItem.Index = 5 Then
        fraAlerting.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Responses
    ElseIf tspConfiguration.SelectedItem.Index = 6 Then
        fraResponses.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraSpeech.Visible = False
        fraLogs.Visible = False
        fraHelp.Visible = False
    
    'Suggestions
    ElseIf tspConfiguration.SelectedItem.Index = 7 Then
        fraSuggestions.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False

    'Searchengine
    ElseIf tspConfiguration.SelectedItem.Index = 8 Then
        fraSearchengine.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSuggestions.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Reports
    ElseIf tspConfiguration.SelectedItem.Index = 9 Then
        fraReports.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraResponses.Visible = False
        fraSuggestions.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Logs
    ElseIf tspConfiguration.SelectedItem.Index = 10 Then
        fraLogs.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraSpeech.Visible = False
        fraHelp.Visible = False
    
    'Speech
    ElseIf tspConfiguration.SelectedItem.Index = 11 Then
        fraSpeech.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraLogs.Visible = False
        fraHelp.Visible = False
    
    'Help
    ElseIf tspConfiguration.SelectedItem.Index = 12 Then
        fraHelp.Visible = True
        fraTarget.Visible = False
        fraPreferences.Visible = False
        fraMapping.Visible = False
        fraPlugins.Visible = False
        fraAlerting.Visible = False
        fraSuggestions.Visible = False
        fraResponses.Visible = False
        fraReports.Visible = False
        fraSearchengine.Visible = False
        fraSpeech.Visible = False
        fraLogs.Visible = False
    End If
End Sub

Private Sub txtDefaultSleep_Change()
    If Len(txtDefaultSleep.Text) < 1 Then
        txtDefaultSleep.Text = 1
    Else
        If txtDefaultSleep.Text < 1 Then
            txtDefaultSleep.Text = 1
        End If
    End If
End Sub

Private Sub txtDefaultSleep_DblClick()
    txtDefaultSleep.Text = 3000
End Sub

Private Sub txtDefaultSleep_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack
    Case Else
        KeyAscii = 0
  End Select
End Sub

Private Sub txtDefaultSleep_LostFocus()
    Call DetectConfigAltering("default sleep time", CStr(application_sleep_time_default), txtDefaultSleep.Text)
End Sub

Private Sub txtLogsDirectory_Change()
    Call ValidateDirectoryInTextBox(txtLogsDirectory)
End Sub

Private Sub txtLogsDirectory_LostFocus()
    Call DetectConfigAltering("logs directory", application_log_directory, txtLogsDirectory.Text)
End Sub

Private Sub txtPluginsDirectory_Change()
    Call ValidateDirectoryInTextBox(txtPluginsDirectory)
End Sub

Private Sub txtPluginsDirectory_LostFocus()
    Call DetectConfigAltering("plugins directory", application_plugin_directory, txtPluginsDirectory.Text)
End Sub

Private Sub txtReportsDirectory_Change()
    Call ValidateDirectoryInTextBox(txtReportsDirectory)
End Sub

Private Sub txtReportsDirectory_LostFocus()
    Call DetectConfigAltering("reports directory", application_report_directory, txtReportsDirectory.Text)
End Sub

Private Sub txtResponsesDirectory_Change()
    Call ValidateDirectoryInTextBox(txtResponsesDirectory)
End Sub

Private Sub txtResponsesDirectory_LostFocus()
    Call DetectConfigAltering("responses directory", application_response_directory, txtResponsesDirectory.Text)
End Sub

Private Sub txtSuggestionsDirectory_Change()
    Call ValidateDirectoryInTextBox(txtSuggestionsDirectory)
End Sub

Private Sub txtSuggestionsDirectory_LostFocus()
    Call DetectConfigAltering("suggestions directory", application_suggestion_directory, txtSuggestionsDirectory.Text)
End Sub

Private Sub txtTarget_LostFocus()
    Dim strNewTarget As String
    
    strNewTarget = txtTarget.Text
    
    If Mid$(strNewTarget, 1, 7) = "http://" Then
        strNewTarget = Mid$(strNewTarget, 8, Len(strNewTarget))
        Call errTargetWrongSpecification
    ElseIf Mid$(strNewTarget, 1, 6) = "ftp://" Then
        strNewTarget = Mid$(strNewTarget, 7, Len(strNewTarget))
        Call errTargetWrongSpecification
    ElseIf Mid$(strNewTarget, 1, 2) = "\\" Then
        strNewTarget = Mid$(strNewTarget, 3, Len(strNewTarget))
        Call errTargetWrongSpecification
    End If
    
    txtTarget.Text = strNewTarget
    
    If LenB(strNewTarget) = 0 Then
        MsgBox ("Target missing." & vbNewLine & vbNewLine & _
            "Please enter the host name or IP address of the target."), vbInformation, "Attack Tool Kit error"
        txtTarget.SetFocus
    End If
    
    Call DetectConfigAltering("Target", Target, txtTarget.Text)
End Sub

Private Sub txtTimeout_Change()
    If LenB(txtTimeout.Text) = 0 Then
        txtTimeout.Text = 10000
    Else
        If txtTimeout.Text < 10000 Then
            txtTimeout.Text = 10000
        End If
    End If
End Sub

Private Sub txtTimeout_DblClick()
    txtTimeout.Text = 30000
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub SaveConfigurationData()
    Dim strLogsSecurityLevel As String  'We presave the security level for further analysis
    
    'Write the new values
    Target = txtTarget.Text
    
    application_attack_timeout = Val(txtTimeout.Text)
    application_sleep_time_default = Val(txtDefaultSleep.Text)

    If application_plugin_directory <> txtPluginsDirectory.Text Then
        application_plugin_directory = txtPluginsDirectory.Text
        If (Dir$(application_plugin_directory, 16) <> "") Then
            frmMain.filATKPlugins.Path = application_plugin_directory
            Call frmMain.mnuPluginsReloadAllItem_Click
        End If
    End If
    
    If application_plugin_download_url <> cmbPluginsDownloadURL.Text Then
        application_plugin_download_url = cmbPluginsDownloadURL.Text
    End If
    
    If application_plugin_external_editor <> cmbPluginsExternalEditor.Text Then
        application_plugin_external_editor = cmbPluginsExternalEditor.Text
    End If
    
    If application_searchengine_url <> cmbSearchEngineURL.Text Then
        application_searchengine_url = cmbSearchEngineURL.Text
    End If
    
    If application_help_url <> cmbHelpURL.Text Then
        application_help_url = cmbHelpURL.Text
    End If
    
    If application_suggestion_directory <> txtSuggestionsDirectory.Text Then
        application_suggestion_directory = txtSuggestionsDirectory.Text
    End If
    
    If application_response_directory <> txtResponsesDirectory.Text Then
        application_response_directory = txtResponsesDirectory.Text
    End If
    
    If application_report_directory <> txtReportsDirectory.Text Then
        application_report_directory = txtReportsDirectory.Text
    End If

    If optSingleCheck.Value = True Then
        application_attack_mode = "SingleCheck"
    ElseIf optFullAudit.Value = True Then
        application_attack_mode = "FullAudit"
    End If

    If chkDoSilentChecks.Value = 1 Then
        application_silent_checks_enable = True
    Else
        application_silent_checks_enable = False
    End If
    
    If chkDoNoDoSChecks.Value = 1 Then
        application_no_dos_enable = True
    Else
        application_no_dos_enable = False
    End If

    If chkDoICMPMapping.Value = 1 Then
        application_icmp_mapping_enable = True
    Else
        application_icmp_mapping_enable = False
    End If

    If chkScanifICMPfails.Value = 1 Then
        application_icmp_mapping_ignore_enable = True
    Else
        application_icmp_mapping_ignore_enable = False
    End If
    
    If chkSuggestions.Value = 1 Then
        application_suggestion_enable = True
    Else
        application_suggestion_enable = False
    End If
        
    If chkAlertingVulnerabilityFound.Value = 1 Then
        application_vulnerability_found_alert_enable = True
    Else
        application_vulnerability_found_alert_enable = False
    End If
    
    If chkAlertingVulnerabilityNotFound.Value = 1 Then
        application_vulnerability_not_found_alert_enable = True
    Else
        application_vulnerability_not_found_alert_enable = False
    End If
    
    If chkReportOpenAfterGeneration.Value = 1 Then
        application_report_open_enable = True
    Else
        application_report_open_enable = False
    End If
    
    If chkActivateLogs.Value = 1 Then
        application_log_enable = True
    Else
        application_log_enable = False
    End If
    
    'We presave the data
    strLogsSecurityLevel = cmbLogsSecurityLevel.Text
    
    If InStrB(1, strLogsSecurityLevel, "0 ", vbBinaryCompare) Then
        application_log_security_level = 0
    ElseIf InStrB(1, strLogsSecurityLevel, "1 ", vbBinaryCompare) Then
        application_log_security_level = 1
    ElseIf InStrB(1, strLogsSecurityLevel, "2 ", vbBinaryCompare) Then
        application_log_security_level = 2
    ElseIf InStrB(1, strLogsSecurityLevel, "3 ", vbBinaryCompare) Then
        application_log_security_level = 3
    ElseIf InStrB(1, strLogsSecurityLevel, "4 ", vbBinaryCompare) Then
        application_log_security_level = 4
    ElseIf InStrB(1, strLogsSecurityLevel, "5 ", vbBinaryCompare) Then
        application_log_security_level = 5
    ElseIf InStrB(1, strLogsSecurityLevel, "6 ", vbBinaryCompare) Then
        application_log_security_level = 6
    Else
        application_log_security_level = 7
    End If
    
    If application_log_directory <> txtLogsDirectory.Text Then
        application_log_directory = txtLogsDirectory.Text
    End If
    
    If chkActivateSpeech.Value = 1 Then
        application_speech_enable = True
    Else
        application_speech_enable = False
    End If

    If IsFormVisible("frmAttackVisualizing") = True Then
        frmAttackVisualizing.txtTargetData.Text = Target
        
        If InStrB(1, Target, "192.", vbBinaryCompare) Then
            frmAttackVisualizing.lblNetworkName.Caption = "LAN"
        ElseIf InStrB(1, Target, "172.", vbBinaryCompare) Then
            frmAttackVisualizing.lblNetworkName.Caption = "LAN"
        ElseIf InStrB(1, Target, "10.", vbBinaryCompare) Then
            frmAttackVisualizing.lblNetworkName.Caption = "LAN"
        ElseIf InStrB(1, Target, "127.", vbBinaryCompare) Then
            frmAttackVisualizing.lblNetworkName.Caption = "Localhost"
        Else
            frmAttackVisualizing.lblNetworkName.Caption = "Internet"
        End If
    End If
End Sub

Private Sub ValidateDirectoryInTextBox(ByRef txbTextBoxName As TextBox)
    If (Dir$(txbTextBoxName.Text, 16) <> "") Then
        txbTextBoxName.BackColor = &H80000005
    Else
        txbTextBoxName.BackColor = &HC0C0FF
    End If
End Sub

Private Sub DetectConfigAltering(ByRef strElementName As String, _
                                ByRef strPublicVariable As String, _
                                ByRef strElement As String)
    
    'Write the new data in the public variable
    If strPublicVariable <> strElement Then
        strPublicVariable = strElement
        
        'Write the log entry
        WriteLogEntry "Changed the " & strElementName & ".", 6
        
        'Tag the config as edited
        Call TagConfigAsEdited
    End If
End Sub

Private Sub TagConfigAsEdited()
    'Tag the config as edited
    bolConfigurationIsEdited = True
End Sub

Private Function CheckIfConfigIsEdited() As Boolean
    Dim iMsgBoxResponse As Integer
    
    'Set the focus to prevent not seen changes. The DoEvents is needed!
    tspConfiguration.SetFocus
    DoEvents
    
    CheckIfConfigIsEdited = False
    
    If bolConfigurationIsEdited = True Then
    iMsgBoxResponse = MsgBox("You have changed the behavior of the software by" & vbNewLine & _
            "changing the configuration." & vbNewLine & vbNewLine & _
            "Would you like to save the existing configuration?", _
            vbYesNoCancel + vbInformation, "Attack Tool Kit configuration changed")
                
        If iMsgBoxResponse = vbYes Then
            Call mnuFileSaveItem_Click
            Unload Me
        ElseIf iMsgBoxResponse = vbNo Then
            Call LoadConfigFromFile(application_configuration_filename)
            Call LoadActualConfigurationValues
        ElseIf iMsgBoxResponse = vbCancel Then
            CheckIfConfigIsEdited = True
        End If
    End If
End Function

Private Sub LoadActualConfigurationValues()
    'Show the frame title with configuration file name
    frmConfiguration.Caption = "Configuration - " & application_configuration_filename
    
    'Display and activate the loaded config data
    txtTarget.Text = Target
    txtTimeout.Text = Val(application_attack_timeout)
    txtDefaultSleep.Text = Val(application_sleep_time_default)
    cmbPluginsDownloadURL.Text = application_plugin_download_url
    cmbSearchEngineURL.Text = application_searchengine_url
    cmbHelpURL.Text = application_help_url
    
    txtPluginsDirectory.Text = application_plugin_directory
    txtSuggestionsDirectory.Text = application_suggestion_directory
    txtResponsesDirectory.Text = application_response_directory
    txtReportsDirectory.Text = application_report_directory
    
    cmbPluginsExternalEditor.Text = application_plugin_external_editor
    
    If application_attack_mode = "FullAudit" Then
        optSingleCheck.Value = False
        optFullAudit.Value = True
    Else
        optSingleCheck.Value = True
        optFullAudit.Value = False
    End If

    If application_silent_checks_enable = True Then
        chkDoSilentChecks.Value = 1
    Else
        chkDoSilentChecks.Value = 0
    End If
    
    If application_report_open_enable = True Then
        chkReportOpenAfterGeneration.Value = 1
    Else
        chkReportOpenAfterGeneration.Value = 0
    End If
    
    If application_no_dos_enable = True Then
        chkDoNoDoSChecks.Value = 1
    Else
        chkDoNoDoSChecks.Value = 0
    End If

    If application_icmp_mapping_enable = True Then
        chkDoICMPMapping.Value = 1
    Else
        chkDoICMPMapping.Value = 0
    End If

    If application_icmp_mapping_ignore_enable = True Then
        chkScanifICMPfails.Value = 1
    Else
        chkScanifICMPfails.Value = 0
    End If
    
    If application_vulnerability_found_alert_enable = True Then
        chkAlertingVulnerabilityFound.Value = 1
    Else
        chkAlertingVulnerabilityFound.Value = 0
    End If
    
    If application_vulnerability_not_found_alert_enable = True Then
        chkAlertingVulnerabilityNotFound.Value = 1
    Else
        chkAlertingVulnerabilityNotFound.Value = 0
    End If
    
    If application_suggestion_enable = True Then
        chkSuggestions.Value = 1
    Else
        chkSuggestions.Value = 0
    End If
    
    If application_log_enable = True Then
        chkActivateLogs.Value = 1
    Else
        chkActivateLogs.Value = 0
    End If
    
    txtLogsDirectory.Text = application_log_directory
    
    If application_speech_enable = True Then
        chkActivateSpeech.Value = 1
    Else
        chkActivateSpeech.Value = 0
    End If

    If application_log_security_level = 0 Then
        cmbLogsSecurityLevel.ListIndex = 0
    ElseIf application_log_security_level = 1 Then
        cmbLogsSecurityLevel.ListIndex = 1
    ElseIf application_log_security_level = 2 Then
        cmbLogsSecurityLevel.ListIndex = 2
    ElseIf application_log_security_level = 3 Then
        cmbLogsSecurityLevel.ListIndex = 3
    ElseIf application_log_security_level = 4 Then
        cmbLogsSecurityLevel.ListIndex = 4
    ElseIf application_log_security_level = 5 Then
        cmbLogsSecurityLevel.ListIndex = 5
    ElseIf application_log_security_level = 6 Then
        cmbLogsSecurityLevel.ListIndex = 6
    Else
        cmbLogsSecurityLevel.ListIndex = 7
    End If
End Sub

Private Sub txtTimeout_LostFocus()
    Call DetectConfigAltering("plugins timeout", CStr(application_attack_timeout), txtTimeout.Text)
End Sub
