VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLog 
   Caption         =   "Log"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7215
   Icon            =   "frmLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filLogs 
      Height          =   870
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdgFileOpen 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load log file"
      Filter          =   "ATK Log Files (*.log)|*.log|All Files (*.*)|*.*"
   End
   Begin VB.Frame fraLogData 
      Caption         =   "Log Data (no log file has been loaded)"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin MSComctlLib.ListView lsvLog 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Security Level"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpLogHelpItem 
         Caption         =   "&Log Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.0 2004-10-16                                                           *
' * - Corrected the procedure to show the last log file.                             *
' * - Added the possibility of resizing this sub form.                               *
' * Version 3.0 2004-10-11                                                           *
' * - Added the whole procedures for handling the new logging security levels.       *
' ************************************************************************************

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
    Dim intLoadableLogFiles As Integer  'How many log files are available
    
    WriteLogEntry "Loading the " & frmLog.Caption & " window.", 6
    
    'Load the default lof file
    If (Dir$(application_log_directory, 16) <> "") Then
        filLogs.Path = application_log_directory
        
        intLoadableLogFiles = filLogs.ListCount
        
        'Check if there are log files available
        If intLoadableLogFiles Then
            'Load the latest log file
            filLogs.ListIndex = intLoadableLogFiles - 1
            Call LoadLogEntries(application_log_directory & "\" & filLogs.Filename)
        Else
            Call errLogDirectoryEmpty
        End If
    Else
        Call errLogDirectoryNotExist
    End If
End Sub

Private Sub Form_Resize()
    If frmLog.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If frmLog.Height < 3000 Then
            frmLog.Height = 3000
        End If
        
        'Prevent zu small windows in width
        If frmLog.Width < 6000 Then
            frmLog.Width = 6000
        End If
        
        fraLogData.Height = frmLog.Height - 920
        lsvLog.Height = fraLogData.Height - 360
        
        fraLogData.Width = frmLog.Width - 360
        lsvLog.Width = fraLogData.Width - 240
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLogEntry "Unloading the " & frmLog.Caption & " window.", 6
    Set frmLog = Nothing
End Sub

Private Sub lsvLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewColumnReorder(frmLog.lsvLog, ColumnHeader)
End Sub

Private Sub lsvLog_DblClick()
    If lsvLog.ListItems.Count Then
        MsgBox "Date:" & vbTab & lsvLog.SelectedItem.Text & vbNewLine & _
            "Time:" & vbTab & lsvLog.SelectedItem.SubItems(1) & vbNewLine & _
            "Text:" & vbTab & lsvLog.SelectedItem.SubItems(2), _
            vbOKOnly, "Attack Tool Kit log entry detailed view"
    Else
        Call errLogDirectoryEmpty
    End If
End Sub

Private Sub mnuFileCloseItem_Click()
    Unload Me
End Sub

Public Sub mnuFileOpenItem_Click()
    Dim LogFileName As String    'Here we save the desired filename for the new plugin
    
    'Define the initial directory of the plugins
    cdgFileOpen.InitDir = application_log_directory
    
    'Ask the user for the desired filename
    cdgFileOpen.ShowOpen 'Opens the save dialog
    LogFileName = cdgFileOpen.Filename 'Get the filename
    
    'Check if a file was selected
    If LenB(LogFileName) Then
        'Check if the file exists
        If (Dir$(LogFileName, 16) <> "") Then
            'Load a new log entry
            WriteLogEntry "Opening the log file " & LogFileName, 6
            Call LoadLogEntries(LogFileName)
        End If
    End If
End Sub

Private Sub LoadLogEntries(Filename As String)
    Dim intFreeFile As Integer  'Free file integer
    Dim List As ListItem        'Needed for the listview handling
    Dim TempString As String    'Here we save the lines
    Dim TempArray() As String   'In this array we save the split result
    
    'Delete the old displayed log data
    lsvLog.ListItems.Clear
    
    'Open and read the plugin file
    If (Dir$(Filename, 16) <> "") Then
        intFreeFile = FreeFile
        Open Filename For Input As #intFreeFile
            Do While Not EOF(intFreeFile)
                Line Input #intFreeFile, TempString
                    
                'Split the log data to be written
                TempArray = Split(TempString, ";")
                
                'Write the log data into the log frame
                On Error Resume Next    'Just a workaround because I get strange errors
                Set List = lsvLog.ListItems.Add(, , TempArray(0))
                    List.SubItems(1) = TempArray(1)
                    List.SubItems(2) = TempArray(2)
                    List.SubItems(3) = TempArray(3)
            Loop
        Close
    
        'Set the right column width
        LVColumnWidth lsvLog
    End If

    'Edit the frame title
    Me.Caption = "Log - " & Filename
    fraLogData.Caption = "Log Data (" & Filename & ")"
End Sub

Private Sub mnuHelpLogHelpItem_Click()
    Call OpenOnlineHelp("log")
End Sub
