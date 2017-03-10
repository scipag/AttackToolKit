VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   6015
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8415
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraReportData 
      Caption         =   "Report (no report loaded)"
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      Begin MSComctlLib.ListView lsvLoadedReport 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
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
            Text            =   "Plugin Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraReportStatistics 
      Caption         =   "Statictics"
      Height          =   5175
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   7935
      Begin VB.Label lblNotFoundVulnerabilitiesCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label lblFoundVulnerabilitiesCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Not Found Vulnerabilities"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   5
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Found Vulnerabilities"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Shape shpNotFoundVulnerabilities 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   3960
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Shape shpFoundVulnerabilities 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   960
         Top             =   1200
         Width           =   3015
      End
   End
   Begin VB.Frame fraExampleReport 
      Caption         =   "Example Report"
      Height          =   5175
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox txtExampleReport 
         Height          =   4695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   360
         Width           =   7695
      End
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open report file"
      Filter          =   "ATK Report Data (*.report)|*.report|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgSaveAs 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save a file as"
      FileName        =   "newplugin.plugin"
      Filter          =   "HTML Report (*.html)|*.html|TXT Report (*.txt)|*.txt|Nessus NSR Report (*.nsr)|*.nsr"
   End
   Begin MSComctlLib.TabStrip tspReportTypes 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10186
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Report &Data"
            Object.ToolTipText     =   "Data of the loaded report"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Example Report"
            Object.ToolTipText     =   "Example of a possible text report"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Statistics"
            Object.ToolTipText     =   "Statistical data of the loaded report"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileReloadItem 
         Caption         =   "&Reload"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAsItem 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReportHelp 
         Caption         =   "&Report Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.1 2005-01-17                                                           *
' * - Added the sorting and re-sorting of the listview.                              *
' * - Fixed a run-time error due a double click in an empty listview.                *
' * Version 4.0 2004-12-05                                                           *
' * - Made the first preparations for the new reporting functionality in 4.0.        *
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
    Call LoadReportData(application_report_directory & "\" & Target & "\" & Target & ".report")
    Call PrepareFrameCaption
End Sub

Private Sub PrepareFrameCaption()
    If LenB(report_filecontent) Then
        Me.Caption = "Report - " & report_filename
        fraReportData.Caption = "Report (" & report_filename & ")"
    Else
        Me.Caption = "Report"
        fraReportData.Caption = "Report (no report loaded)"
    End If
End Sub

Private Sub LoadReportData(ByRef strReportFileName As String)
    Dim i As Integer
    Dim intFreeFile As Integer
    Dim List As ListItem        'Needed for the listview handling
    Dim strReportFileLinesArray() As String
    Dim intReportFileLinesCount As Integer
    Dim strReportFileLineDataArray() As String
    Dim intFoundVulnerabilitiesCount As Integer
    Dim intNotFoundVulnerabilitiesCount As Integer
    
    Call ReleaseWindow(False)

    report_filename = strReportFileName
    
    If (Dir$(report_filename, 16) <> "") Then
        WriteLogEntry "Loading the report file " & report_filename & " ...", 6
        intFreeFile = FreeFile
        Open report_filename For Input As #intFreeFile
            report_filecontent = Input(LOF(intFreeFile), #intFreeFile)
        Close

        report_filesize = Len(report_filecontent)
        
        strReportFileLinesArray = Split(report_filecontent, vbNewLine, , vbBinaryCompare)
        
        intReportFileLinesCount = UBound(strReportFileLinesArray)
        
        lsvLoadedReport.ListItems.Clear
        Call PrepareFrameCaption
        
        For i = 0 To intReportFileLinesCount
            strReportFileLineDataArray = Split(strReportFileLinesArray(i), ";", , vbBinaryCompare)
        
            If UBound(strReportFileLineDataArray) = 3 Then
                Set List = lsvLoadedReport.ListItems.Add(, , strReportFileLineDataArray(0))
                    If strReportFileLineDataArray(1) = "1" Then
                        List.SubItems(1) = "Found (1)"
                        lsvLoadedReport.ListItems(i + 1).ListSubItems(1).ForeColor = &HC0&
                        intFoundVulnerabilitiesCount = intFoundVulnerabilitiesCount + 1
                    Else
                        List.SubItems(1) = "Not Found (0)"
                        lsvLoadedReport.ListItems(i + 1).ListSubItems(1).ForeColor = &HC000&
                        intNotFoundVulnerabilitiesCount = intNotFoundVulnerabilitiesCount + 1
                    End If
                    List.SubItems(2) = strReportFileLineDataArray(2)
                    List.SubItems(3) = strReportFileLineDataArray(3)
            End If
        Next i
        
        LVColumnWidth lsvLoadedReport
        
        'Generate the report statistics
        lblFoundVulnerabilitiesCount.Caption = intFoundVulnerabilitiesCount
        lblNotFoundVulnerabilitiesCount.Caption = intNotFoundVulnerabilitiesCount
    Else
        WriteLogEntry "Could not read the report file " & report_filename & " ...", 3
    End If

    Call ReleaseWindow(True)
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If frmLog.Height < 3000 Then
            frmLog.Height = 3000
        End If
        
        'Prevent zu small windows in width
        If frmLog.Width < 6000 Then
            frmLog.Width = 6000
        End If
        
        tspReportTypes.Height = Me.Height - 920
        fraReportData.Height = tspReportTypes.Height - 620
        fraExampleReport.Height = fraReportData.Height
        txtExampleReport.Height = fraReportData.Height - 480
        fraReportStatistics.Height = fraReportData.Height
        lsvLoadedReport.Height = txtExampleReport.Height
        
        tspReportTypes.Width = Me.Width - 360
        fraReportData.Width = tspReportTypes.Width - 240
        fraExampleReport.Width = fraReportData.Width
        txtExampleReport.Width = fraReportData.Width - 240
        fraReportStatistics.Width = fraReportData.Width
        lsvLoadedReport.Width = txtExampleReport.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReport = Nothing
End Sub

Private Sub lsvLoadedReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewColumnReorder(frmReport.lsvLoadedReport, ColumnHeader)
End Sub

Private Sub lsvLoadedReport_DblClick()
    Dim strPluginFileName As String
    Dim strPluginContent As String
    
    If lsvLoadedReport.ListItems.Count Then
        'Save the plugin_filename
        strPluginFileName = lsvLoadedReport.SelectedItem.Text
        
        strPluginContent = ReadPluginFromFile(strPluginFileName, plugin_filepath)
        
        If LenB(strPluginContent) Then
            plugin_filename = strPluginFileName
            Call ParseATKPlugin(strPluginContent)
            MsgBox Replace$(GenerateTXTReportPluginEntry(True, ""), "     ", vbNullString, , , vbBinaryCompare)
        End If
    End If
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuFileOpenItem_Click()
    Dim strReportFileName As String  'The name of the configuration file
    Dim strDefaultReportPath As String
    
    Call ReleaseWindow(False)
    
    strDefaultReportPath = application_report_directory & "\" & Target
    
    'Define the initial directory of the reports
    If Not (Dir$(strDefaultReportPath, 16) <> "") Then
        strDefaultReportPath = App.Path
    End If
    
    cdgOpen.Filename = Target & ".report"
    
    'Define the initial directory of the plugins
    cdgOpen.InitDir = strDefaultReportPath
    
    'Ask the user for the desired filename
    cdgOpen.ShowOpen 'Opens the save dialog
    
    'Cache the filename into a variant to increase the speed
    strReportFileName = cdgOpen.Filename
    
    'Check if a file was selected
    If LenB(strReportFileName) Then
        'Check if the file exists
        If (Dir$(strReportFileName, 16) <> "") Then
            'Load the configuration file
            Call LoadReportData(strReportFileName)
        End If
    End If
    
    Call GenerateExampleReport
    
    Call ReleaseWindow(True)
End Sub

Private Sub mnuFileReloadItem_Click()
    Call LoadReportData(report_filename)
    Call GenerateExampleReport
End Sub

Private Sub mnuFileSaveAsItem_Click()
    Dim strReportDestinationFileName As String    'Here we save the desired filename for the new report
    
    Call ReleaseWindow(False)
    
    'Define the initial directory of the plugins
    If (Dir$(application_report_directory & "\" & Target, 16) <> "") Then
        cdgSaveAs.InitDir = application_report_directory & "\" & Target
    Else
        cdgSaveAs.InitDir = application_report_directory
    End If
    
    'Ask the user for the desired filename
    On Error GoTo ErrSub
    cdgSaveAs.Filename = Target
    cdgSaveAs.ShowSave 'Opens the save dialog
    strReportDestinationFileName = cdgSaveAs.Filename 'Get the filename
    
    'Cut the plugin extension if there is one given
    If LenB(strReportDestinationFileName) Then
        If cdgSaveAs.FilterIndex = 1 Then
            Call GenerateHTMLReport(report_filename, strReportDestinationFileName)
        ElseIf cdgSaveAs.FilterIndex = 2 Then
            Call WriteTXTReportToFile(report_filename, strReportDestinationFileName)
            'Call GenerateTXTReport(report_filename, strReportDestinationFileName)
        ElseIf cdgSaveAs.FilterIndex = 3 Then
            Call GenerateNSRReport(report_filename, strReportDestinationFileName)
        End If
    End If

ErrSub:

    Call ReleaseWindow(True)
End Sub

Private Sub mnuHelpReportHelp_Click()
    Call OpenOnlineHelp("report")
End Sub

Private Sub tspReportTypes_Click()
    Dim intSelectedItem As Integer
    intSelectedItem = tspReportTypes.SelectedItem.Index
    
    If intSelectedItem = 1 Then
        fraReportData.Visible = True
        fraReportStatistics.Visible = False
        fraExampleReport.Visible = False
    ElseIf intSelectedItem = 2 Then
        fraExampleReport.Visible = True
        fraReportStatistics.Visible = False
        fraReportData.Visible = False
        Call GenerateExampleReport
    ElseIf intSelectedItem = 3 Then
        fraReportStatistics.Visible = True
        fraReportData.Visible = False
        fraExampleReport.Visible = False
    End If
End Sub

Private Sub ReleaseWindow(ByRef bolFreeze As Boolean)
    mnuFile.Enabled = bolFreeze
End Sub

Private Sub GenerateExampleReport()
    If txtExampleReport.Visible = True Then
        txtExampleReport.Text = GenerateTXTReport(report_filename)
    End If
End Sub

Private Sub txtExampleReport_DblClick()
    Call OpenSelectedTextIfItIsURL(txtExampleReport.SelText)
End Sub
