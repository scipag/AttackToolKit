VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplashScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Attack Tool Kit"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmSplashScreen.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   13  'Arrow and Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer timTimer 
      Interval        =   55
      Left            =   4080
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar pbrStatus 
      Height          =   135
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© 2003-2005 by Marc Ruef"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblStatusInformation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "loading the software into the memory"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack Tool Kit is starting ... Please wait!"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.Shape shpRedLine 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = application_name & " starting ..."
    lblStatus.Caption = application_name & " is starting ... Please wait!"
End Sub

Private Sub timTimer_Timer()
    timTimer.Enabled = False
    Call CheckForUpdate
    Call StartInitialisation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplashScreen = Nothing
End Sub

Private Sub StartInitialisation()
    pbrStatus.Value = 0
    
    'Write that the software hast started. This has to be done before any other
    'routine has the ability to write a log entry. If this is not the first entry,
    'the whole log will be unorderd.
    WriteLogEntry application_name & " started.", 6
    pbrStatus.Value = 5
    
    'Load the last configuration
    lblStatusInformation.Caption = "loading the default configuration"
    Call LoadConfigFromFile
    pbrStatus.Value = 15
    
    'Check the existence of the directories before loading the data.
    lblStatusInformation.Caption = "checking the local directories"
    Call CheckDirectoriesBeforeLoading
    pbrStatus.Value = 25
   
    'Load initially the default report structure
    lblStatusInformation.Caption = "preparing the report structure"
    Call LoadDefaultReportStructure
    pbrStatus.Value = 30
    
    'Load the plugins initially into the list
    lblStatusInformation.Caption = "loading the available plugins"
    Call frmMain.LoadATKPlugins
    pbrStatus.Value = 60
    
    Call frmMain.LoadNASLPlugins
    pbrStatus.Value = 90
    
    'Pre-show the main frame
    frmMain.Visible = True
    
    'Prepare system variables
    Call LoadUserName
    pbrStatus.Value = 100
    
    'Handle the splash screen
    lblStatusInformation.Caption = "finishing the loading sequence"
    
    Unload Me
End Sub

' *************************************************************************
' * Check the existence of the needed and wanted directories. If they are *
' * not available and really needed, show a message and create them. We   *
' * prevent error checking during runtime and unpredictable errors.       *
' * Note: The plugins directory is checked in the procedure for loading   *
' *       the plugins. That is done because the check may needed on every *
' *       refresh because new loggins were loaded.                        *
' *************************************************************************

Private Sub CheckDirectoriesBeforeLoading()
    'Check the existence of the suggestions directory
    WriteLogEntry "Checking the existence of the suggestions directory " & application_suggestion_directory & " ...", 6
    If Not (Dir$(application_suggestion_directory, 16) <> "") Then
        Call errSuggestionsDirectoryNotExist
    End If

    'Check the existence of the logs directory
    WriteLogEntry "Checking the existence of the logs directory " & application_log_directory & " ...", 6
    If Not (Dir$(application_log_directory, 16) <> "") Then
        Call errLogDirectoryNotExist
    End If
End Sub

Private Sub CheckForUpdate()
    Dim datDate As Date
    
    datDate = Date
    
    If DatePart("yyyy", datDate) > 2004 Then
        If DatePart("m", datDate) > 8 Then
            If MsgBox("This version of the ATK software has been published in december 2004." & vbNewLine & _
                "In the meanwhile a new and updated version of the software may be available." & vbNewLine & vbNewLine & _
                "You should use the latest release to gain the maximum profit." & vbNewLine & _
                "Would you like to open the ATK project web site to get the latest software release?", _
                vbYesNo + vbInformation, "Attack Tool Kit new version available") = vbYes Then
                
                'Open the WebSite
                Call OpenProjectWebsite
                End
            End If
        End If
    End If
End Sub
