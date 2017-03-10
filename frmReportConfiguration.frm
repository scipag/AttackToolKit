VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportConfiguration 
   Caption         =   "Report Configuration"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11070
   Icon            =   "frmReportConfiguration.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11070
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdgSaveAs 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save a file as"
      FileName        =   "new_report_template.reporttemplate"
      Filter          =   "ATK Report Templates (*.reporttemplate)|*.reporttemplate|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open file"
      FileName        =   "default.reporttemplate"
      Filter          =   "ATK Report Template (*.reporttemplate)|*.reporttemplate|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.Frame fraExample 
      Caption         =   "Example Report Item"
      Height          =   6735
      Left            =   6360
      TabIndex        =   8
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtReportExample 
         Height          =   6255
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame fraVulnerabilitiesCustomizing 
      Caption         =   "Vulnerabilities Customizing"
      Height          =   6735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdVulnerabilityDown 
         Caption         =   "dn"
         Height          =   255
         Left            =   5640
         TabIndex        =   5
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdVulnerabilityUp 
         Caption         =   "up"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   3240
         Width           =   375
      End
      Begin VB.ListBox lstVulnerabilityPositions 
         Height          =   5910
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   2415
      End
      Begin VB.ListBox lstVulnerabilityReport 
         Height          =   5910
         Left            =   3120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdVulnerabilityRemove 
         Caption         =   "-"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdVulnerabilityAdd 
         Caption         =   "+"
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblLabel 
         Caption         =   "Actual report structure"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "Available positions"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label lblDragAndDrop 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAsItem 
         Caption         =   "Save &As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReportConfigurationHelpItem 
         Caption         =   "Report Configuration Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmReportConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Frame Description                                                                *
' *                                                                                  *
' * In this frame the user is able to configure the whole reporting.                 *
' ************************************************************************************

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2004-12-27                                                           *
' * - Fixed a bug in the report listview if a command button is pressed and no index *
' *   is selected.                                                                   *
' * Version 4.0 2004-12-24                                                           *
' * - Re-ordered the available positions alphabetically.                             *
' * - Added a  function to focus the new line of the example report after adding a   *
' *   new tag.                                                                       *
' * Version 4.0 2004-12-23                                                           *
' * - Also added the possibility of double clicking urls to open them in the browser.*
' * - Changed the frame boarder style to sizeable and added the resize sub.          *
' * Version 4.0 2004-12-21                                                           *
' * - Completely replaced the reporting template routines.                           *
' * Version 3.0 2004-11-01                                                           *
' * - Replaced all useless functions with normal subs.                               *
' * - Corrected the tab order to be a more logical.                                  *
' * Version 2.0 2004-08-15                                                           *
' * - Added the whole new fields and some special fields fur further diagnostics.    *
' * - Fixed a nasty error when the up and down buttons were pushed too much.         *
' * - Added and corrected the whole tab stops.                                       *
' ************************************************************************************

Private Sub LoadReportTemplate(ByRef strReportTemplateContent As String)
    Dim i As Integer
    Dim strReportTemplateArray() As String
    Dim intReportTemplateArrayItems As Integer
    
    If LenB(strReportTemplateContent) Then
        strReportTemplateArray = Split(strReportTemplateContent, vbCrLf, , vbBinaryCompare)
        
        intReportTemplateArrayItems = UBound(strReportTemplateArray)
        
        For i = 0 To intReportTemplateArrayItems
            If LenB(strReportTemplateArray(i)) Then
                lstVulnerabilityReport.AddItem strReportTemplateArray(i)
            End If
        Next i
    End If
End Sub

Private Sub cmdVulnerabilityAdd_Click()
    lstVulnerabilityReport.AddItem lstVulnerabilityPositions.Text

    txtReportExample.SelStart = Len(txtReportExample.Text)
End Sub

Private Sub cmdVulnerabilityDown_Click()
    Dim strTemp1 As String    'Hold the selected index data temporarily for move
    Dim iCnt    As Integer    'Holds the index of the item to be moved
        
    If lstVulnerabilityReport.ListCount Then
        'Assign the first index
        If lstVulnerabilityReport.ListIndex >= 0 Then
            iCnt = lstVulnerabilityReport.ListIndex
            
            If iCnt < lstVulnerabilityReport.ListCount - 1 Then
                 
                strTemp1 = lstVulnerabilityReport.List(iCnt)
                
                'Add the item selected to below the current position
                lstVulnerabilityReport.AddItem strTemp1, (iCnt + 2)
                
                lstVulnerabilityReport.RemoveItem (iCnt)
                
                'Reselect the item that was moved.
                lstVulnerabilityReport.Selected(iCnt + 1) = True
            End If
        End If
    End If
End Sub

Public Sub RefreshReportStructure()
    'Recompute the report structure
    Call PrepareReportStructure
    
    fraExample.Caption = "Example Report Item (" & plugin_filename & ")"
    
    txtReportExample.Text = GenerateTXTReportPluginEntry(False, vbNullString)
End Sub

Private Sub cmdVulnerabilityRemove_Click()
    'Prevent errors if no item is selected
    If lstVulnerabilityReport.ListIndex < 0 Then
        lstVulnerabilityReport.ListIndex = lstVulnerabilityReport.ListCount - 1
    End If
        
    'Delete the selected item
    If lstVulnerabilityReport.ListCount > 0 Then
        lstVulnerabilityReport.RemoveItem lstVulnerabilityReport.ListIndex
    End If
End Sub

Private Sub cmdVulnerabilityUp_Click()
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    iCnt = lstVulnerabilityReport.ListIndex
        
        If iCnt > 0 Then
         
        strTemp1 = lstVulnerabilityReport.List(iCnt)
        
        'Add the item selected to one position above the current position
        lstVulnerabilityReport.AddItem strTemp1, (iCnt - 1)
        
        'remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstVulnerabilityReport.RemoveItem (iCnt + 1)
        
        'Reselect the item that was moved.
        lstVulnerabilityReport.Selected(iCnt - 1) = True
    End If
End Sub

Private Sub Form_Activate()
    If Me.WindowState = vbMinimized Then
        Me.WindowState = vbNormal
    End If
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If State = 0 Then Source.MousePointer = 12
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 27
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    'Add the items for the header positions
    lstVulnerabilityPositions.AddItem "<br>"
    
    lstVulnerabilityPositions.AddItem "application_attack_mode"
    lstVulnerabilityPositions.AddItem "application_attack_timeout"
    lstVulnerabilityPositions.AddItem "application_configuration_filename"
    lstVulnerabilityPositions.AddItem "application_help_url"
    lstVulnerabilityPositions.AddItem "application_icmp_mapping_enable"
    lstVulnerabilityPositions.AddItem "application_icmp_mapping_ignore_enable"
    lstVulnerabilityPositions.AddItem "application_log_directory"
    lstVulnerabilityPositions.AddItem "application_log_enable"
    lstVulnerabilityPositions.AddItem "application_log_security_level"
    lstVulnerabilityPositions.AddItem "application_name"
    lstVulnerabilityPositions.AddItem "application_no_dos_enable"
    lstVulnerabilityPositions.AddItem "application_plugin_count"
    lstVulnerabilityPositions.AddItem "application_plugin_directory"
    lstVulnerabilityPositions.AddItem "application_plugin_download_url"
    lstVulnerabilityPositions.AddItem "application_plugin_external_editor"
    lstVulnerabilityPositions.AddItem "application_report_directory"
    lstVulnerabilityPositions.AddItem "application_report_open_enable"
    lstVulnerabilityPositions.AddItem "application_response_directory"
    lstVulnerabilityPositions.AddItem "application_searchengine_url"
    lstVulnerabilityPositions.AddItem "application_silent_checks_enable"
    lstVulnerabilityPositions.AddItem "application_sleep_time_default"
    lstVulnerabilityPositions.AddItem "application_speech_enable"
    lstVulnerabilityPositions.AddItem "application_suggestion_directory"
    lstVulnerabilityPositions.AddItem "application_suggestion_enable"
    lstVulnerabilityPositions.AddItem "application_vulnerability_found_alert_enable"
    lstVulnerabilityPositions.AddItem "application_vulnerability_not_found_alert_enable"
    lstVulnerabilityPositions.AddItem "application_website_url"
    
    lstVulnerabilityPositions.AddItem "bug_advisory"
    lstVulnerabilityPositions.AddItem "bug_affected"
    lstVulnerabilityPositions.AddItem "bug_checking_tool"
    lstVulnerabilityPositions.AddItem "bug_description"
    lstVulnerabilityPositions.AddItem "bug_exploit_availability"
    lstVulnerabilityPositions.AddItem "bug_exploit_url"
    lstVulnerabilityPositions.AddItem "bug_false_negatives"
    lstVulnerabilityPositions.AddItem "bug_false_positives"
    lstVulnerabilityPositions.AddItem "bug_fixing_time"
    lstVulnerabilityPositions.AddItem "bug_impact"
    lstVulnerabilityPositions.AddItem "bug_local"
    lstVulnerabilityPositions.AddItem "bug_iss_scanner_rating"
    lstVulnerabilityPositions.AddItem "bug_netrecon_rating"
    lstVulnerabilityPositions.AddItem "bug_nessus_risk"
    lstVulnerabilityPositions.AddItem "bug_not_affected"
    lstVulnerabilityPositions.AddItem "bug_popularity"
    lstVulnerabilityPositions.AddItem "bug_produced_email"
    lstVulnerabilityPositions.AddItem "bug_produced_name"
    lstVulnerabilityPositions.AddItem "bug_produced_web"
    lstVulnerabilityPositions.AddItem "bug_published_company"
    lstVulnerabilityPositions.AddItem "bug_published_date"
    lstVulnerabilityPositions.AddItem "bug_published_email"
    lstVulnerabilityPositions.AddItem "bug_published_name"
    lstVulnerabilityPositions.AddItem "bug_published_web"
    lstVulnerabilityPositions.AddItem "bug_remote"
    lstVulnerabilityPositions.AddItem "bug_response"
    lstVulnerabilityPositions.AddItem "bug_risk"
    lstVulnerabilityPositions.AddItem "bug_severity"
    lstVulnerabilityPositions.AddItem "bug_simplicity"
    lstVulnerabilityPositions.AddItem "bug_solution"
    lstVulnerabilityPositions.AddItem "bug_vulnerability_class"
    
    lstVulnerabilityPositions.AddItem "plugin_changelog"
    lstVulnerabilityPositions.AddItem "plugin_comment"
    lstVulnerabilityPositions.AddItem "plugin_created_company"
    lstVulnerabilityPositions.AddItem "plugin_created_date"
    lstVulnerabilityPositions.AddItem "plugin_created_email"
    lstVulnerabilityPositions.AddItem "plugin_created_name"
    lstVulnerabilityPositions.AddItem "plugin_created_web"
    lstVulnerabilityPositions.AddItem "plugin_detection_accuracy"
    lstVulnerabilityPositions.AddItem "plugin_exploit_accuracy"
    lstVulnerabilityPositions.AddItem "plugin_family"
    lstVulnerabilityPositions.AddItem "plugin_filename"
    lstVulnerabilityPositions.AddItem "plugin_filesize"
    lstVulnerabilityPositions.AddItem "plugin_id"
    lstVulnerabilityPositions.AddItem "plugin_name"
    lstVulnerabilityPositions.AddItem "plugin_port"
    lstVulnerabilityPositions.AddItem "plugin_procedure_detection"
    lstVulnerabilityPositions.AddItem "plugin_procedure_exploit"
    lstVulnerabilityPositions.AddItem "plugin_protocol"
    lstVulnerabilityPositions.AddItem "plugin_updated_company"
    lstVulnerabilityPositions.AddItem "plugin_updated_date"
    lstVulnerabilityPositions.AddItem "plugin_updated_email"
    lstVulnerabilityPositions.AddItem "plugin_updated_name"
    lstVulnerabilityPositions.AddItem "plugin_updated_web"
    lstVulnerabilityPositions.AddItem "plugin_version"
    
    lstVulnerabilityPositions.AddItem "report_structure"
    lstVulnerabilityPositions.AddItem "report_filecontent"
    lstVulnerabilityPositions.AddItem "report_filename"
    'lstVulnerabilityPositions.AddItem "report_filepath"
    lstVulnerabilityPositions.AddItem "report_filesize"
    
    lstVulnerabilityPositions.AddItem "report_template_filecontent"
    lstVulnerabilityPositions.AddItem "report_template_filename"
    lstVulnerabilityPositions.AddItem "report_template_filepath"
    lstVulnerabilityPositions.AddItem "report_template_filesize"
    
    lstVulnerabilityPositions.AddItem "scan_date"
    lstVulnerabilityPositions.AddItem "scan_mode"
    lstVulnerabilityPositions.AddItem "scan_target"
    lstVulnerabilityPositions.AddItem "scan_time"
    
    lstVulnerabilityPositions.AddItem "session_procedure_commands"
    lstVulnerabilityPositions.AddItem "session_procedure_type"

    lstVulnerabilityPositions.AddItem "source_aerasec_id"
    lstVulnerabilityPositions.AddItem "source_arachnids_id"
    lstVulnerabilityPositions.AddItem "source_cert_id"
    lstVulnerabilityPositions.AddItem "source_certvu_id"
    lstVulnerabilityPositions.AddItem "source_ciac_id"
    lstVulnerabilityPositions.AddItem "source_cve"
    lstVulnerabilityPositions.AddItem "source_heise_news"
    lstVulnerabilityPositions.AddItem "source_heise_security"
    lstVulnerabilityPositions.AddItem "source_issxforce_id"
    lstVulnerabilityPositions.AddItem "source_literature"
    lstVulnerabilityPositions.AddItem "source_misc"
    lstVulnerabilityPositions.AddItem "source_mskb_id"
    lstVulnerabilityPositions.AddItem "source_mssb_id"
    lstVulnerabilityPositions.AddItem "source_nessus_id"
    lstVulnerabilityPositions.AddItem "source_netbsdsa_id"
    lstVulnerabilityPositions.AddItem "source_osvdb_id"
    lstVulnerabilityPositions.AddItem "source_rhsa_id"
    lstVulnerabilityPositions.AddItem "source_scip_id"
    lstVulnerabilityPositions.AddItem "source_secunia_id"
    lstVulnerabilityPositions.AddItem "source_securiteam_url"
    lstVulnerabilityPositions.AddItem "source_securitytracker_id"
    lstVulnerabilityPositions.AddItem "source_securityfocus_bid"
    lstVulnerabilityPositions.AddItem "source_snort_id"
    lstVulnerabilityPositions.AddItem "source_tecchannel_id"
    lstVulnerabilityPositions.AddItem "source_uscertta_id"
    
    lstVulnerabilityPositions.AddItem "system_username"
    
    'Load the given report structure
    Call LoadReportTemplate(report_structure)
    
    Call PrepareFrameCaption
    
    'Select the first item in the positions
    lstVulnerabilityPositions.ListIndex = 0
    
    'Refresh the example report
    Call RefreshReportStructure
End Sub

Private Sub PrepareFrameCaption()
    If LenB(report_structure) Then
        Me.Caption = "Report Configuration - " & report_template_filepath & "\" & report_template_filename
    Else
        Me.Caption = "Report Configuration"
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If Me.Height < 4000 Then
            Me.Height = 4000
        End If
        
        'Prevent zu small windows in width
        If Me.Width < 8000 Then
            Me.Width = 8000
        End If
        
        'Customizing elements
        fraVulnerabilitiesCustomizing.Width = Me.Width / 1.8
        fraVulnerabilitiesCustomizing.Height = Me.Height - 920
        
        lblLabel(0).Width = (fraVulnerabilitiesCustomizing.Width / 2) - (cmdVulnerabilityAdd.Width * 2)
        lstVulnerabilityPositions.Width = lblLabel(0).Width
        lstVulnerabilityPositions.Height = fraVulnerabilitiesCustomizing.Height - 780
        
        lblLabel(1).Width = lblLabel(0).Width
        lstVulnerabilityReport.Width = lblLabel(1).Width
        lstVulnerabilityReport.Height = lstVulnerabilityPositions.Height
        
        lblDragAndDrop.Width = lblLabel(0).Width
        
        lblLabel(0).Left = 120
        lstVulnerabilityPositions.Left = lblLabel(0).Left
        cmdVulnerabilityAdd.Left = lstVulnerabilityPositions.Width + 260
        cmdVulnerabilityRemove.Left = cmdVulnerabilityAdd.Left
        cmdVulnerabilityAdd.Top = (lstVulnerabilityPositions.Height / 2)
        cmdVulnerabilityRemove.Top = cmdVulnerabilityAdd.Top + cmdVulnerabilityAdd.Height + 120
        
        lblLabel(1).Left = lstVulnerabilityPositions.Width + cmdVulnerabilityAdd.Width + 420
        lstVulnerabilityReport.Left = lblLabel(1).Left
        cmdVulnerabilityDown.Left = lstVulnerabilityPositions.Width + lstVulnerabilityReport.Width + 920
        cmdVulnerabilityUp.Left = cmdVulnerabilityDown.Left
        cmdVulnerabilityUp.Top = (lstVulnerabilityReport.Height / 2)
        cmdVulnerabilityDown.Top = cmdVulnerabilityUp.Top + cmdVulnerabilityUp.Height + 120
        
        'Example elements
        fraExample.Width = Me.Width - fraVulnerabilitiesCustomizing.Width - 460
        fraExample.Height = fraVulnerabilitiesCustomizing.Height
                
        fraExample.Left = fraVulnerabilitiesCustomizing.Width + 260
    
        txtReportExample.Width = fraExample.Width - 260
        txtReportExample.Height = fraExample.Height - 460
    End If
End Sub

Private Sub lstVulnerabilityPositions_Click()
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'Refresh the example report
    Call RefreshReportStructure
End Sub

Private Sub lstVulnerabilityPositions_DblClick()
    Call AddVulnerabilityPosition
End Sub

Private Sub AddVulnerabilityPosition()
    If lstVulnerabilityPositions.SelCount Then
        lstVulnerabilityReport.AddItem lstVulnerabilityPositions.Text
        
        Call RefreshReportStructure
    
        txtReportExample.SelStart = Len(txtReportExample.Text)
    End If
End Sub

Private Sub lstVulnerabilityPositions_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If State = 0 Then Source.MousePointer = 12
    If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub lstVulnerabilityPositions_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 39, 45, 107
            Call AddVulnerabilityPosition
    End Select
End Sub

Private Sub lstVulnerabilityPositions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Dim DY As Integer
        
        DY = TextHeight("A")
        lblDragAndDrop.Move fraVulnerabilitiesCustomizing.Left + lstVulnerabilityPositions.Left, _
            fraVulnerabilitiesCustomizing.Top + lstVulnerabilityPositions.Top + y - DY * 0.5, lstVulnerabilityPositions.Width, DY
        lblDragAndDrop.Drag
    End If
End Sub

Private Sub lstVulnerabilityReport_Click()
    'Clear the example report
    txtReportExample.Text = vbNullString
    
    'Refresh the example report
    Call RefreshReportStructure
End Sub

Private Sub lstVulnerabilityReport_DblClick()
    Call RemoveVulnerabilityPosition
End Sub

Private Sub RemoveVulnerabilityPosition()
    Dim iReportListIndex As Integer
    
    If lstVulnerabilityReport.SelCount Then
        iReportListIndex = lstVulnerabilityReport.ListIndex
        
        lstVulnerabilityReport.RemoveItem iReportListIndex
        
        If lstVulnerabilityReport.ListCount > iReportListIndex Then
            lstVulnerabilityReport.Selected(iReportListIndex) = True
        End If
        
        'Clear the example report
        txtReportExample.Text = vbNullString
        
        'Refresh the example report
        Call RefreshReportStructure
    End If
End Sub

Private Sub lstVulnerabilityReport_DragDrop(Source As Control, x As Single, y As Single)
    lstVulnerabilityReport.AddItem lstVulnerabilityPositions.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Recompute the report structure
    Call PrepareReportStructure
    
    'Show the new report structure also in the main frame
    frmMain.txtPluginContent.Text = GenerateTXTReportPluginEntry(False, vbNullString)
    
    WriteLogEntry "Unloading the " & Me.Caption & " window.", 6
    Set frmReportConfiguration = Nothing
End Sub

Private Sub lstVulnerabilityReport_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37, 46, 109
            Call RemoveVulnerabilityPosition
    End Select
End Sub

Private Sub mnuFileCloseItem_Click()
    Unload Me
End Sub

Private Sub mnuFileNewItem_Click()
    'Delete the report
    lstVulnerabilityReport.Clear

    'Refresh the example report
    txtReportExample.Text = vbNullString
End Sub

Private Sub mnuFileOpenItem_Click()
    Dim strReportTemplateFileName As String  'The name of the configuration file
    Dim strDefaultReportTemplatePath As String
    
    strDefaultReportTemplatePath = App.Path & "\reporttemplates"
    
    'Define the initial directory of the plugins
    On Error Resume Next
    If Not (Dir$(strDefaultReportTemplatePath, 16) <> "") Then
        strDefaultReportTemplatePath = App.Path
    End If
    
    cdgOpen.Filename = "default.reporttemplate"
    
    'Define the initial directory of the plugins
    cdgOpen.InitDir = strDefaultReportTemplatePath
    
    'Ask the user for the desired filename
    cdgOpen.ShowOpen 'Opens the save dialog
    
    'Cache the filename into a variant to increase the speed
    strReportTemplateFileName = cdgOpen.Filename
    
    'Check if a file was selected
    If LenB(strReportTemplateFileName) Then
        'Check if the file exists
        If (Dir$(strReportTemplateFileName, 16) <> "") Then
            'Delete the report
            lstVulnerabilityReport.Clear
            
            'Load the configuration file
            Call LoadReportTemplate(LoadReportTemplateFromFile(strReportTemplateFileName, vbNullString))
            Call RefreshReportStructure
            
            Me.Caption = "Report Configuration - " & strReportTemplateFileName
        End If
    End If
End Sub

Private Sub mnuFileSaveAsItem_Click()
    Dim strTemplateFileName As String    'Here we save the desired filename for the new template
    
    'Define the initial directory of the reporttemplates
    On Error Resume Next
    cdgSaveAs.InitDir = App.Path & "\reporttemplates"
    
    'Ask the user for the desired filename
    On Error GoTo ErrSub
    cdgSaveAs.Filename = report_template_filename
    cdgSaveAs.ShowSave 'Opens the save dialog
    strTemplateFileName = cdgSaveAs.Filename 'Get the filename
    
    'Cut the plugin extension if there is one given
    If LenB(strTemplateFileName) Then
        'Prepare the new structure
        Call PrepareReportStructure
        
        'Write the new plugin
        Call WriteReportTemplateToFile(strTemplateFileName)
    End If
    
    'Write the new title
    Call PrepareFrameCaption

ErrSub:
End Sub

Private Sub mnuHelpReportConfigurationHelpItem_Click()
    Call OpenOnlineHelp("report_configuration")
End Sub

Private Sub txtReportExample_DblClick()
    Call OpenSelectedTextIfItIsURL(txtReportExample.SelText)
End Sub
