VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAttackResponse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attack Response"
   ClientHeight    =   5910
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open file"
      Filter          =   "ATK Responses (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.TextBox txtLastResponse 
      Height          =   2775
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Frame fraSilentChecks 
      Caption         =   "Silent Checks"
      Height          =   4695
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   5895
      Begin MSComctlLib.ListView lsvSilentPlugins 
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Trigger"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraSuggestions 
      Caption         =   "Suggestions"
      Height          =   4695
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   5895
      Begin VB.FileListBox filSuggestions 
         Height          =   870
         Left            =   4080
         Pattern         =   "*.suggestion"
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.ListView lsvSuggestions 
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "iltSuggestionsIcons"
         SmallIcons      =   "iltSuggestionsIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Suggestion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Trigger"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList iltSuggestionsIcons 
         Left            =   3480
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   16
         ImageHeight     =   16
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAttackResponse.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSuggestion 
         Caption         =   "There is no suggestion selected. Please select a suggestion first to get a hint what to do next."
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   5655
      End
   End
   Begin VB.Frame fraLastResponse 
      Caption         =   "Last Response"
      Height          =   4695
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      Begin VB.Label lblPositionName 
         Caption         =   "Position"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblCursorPosition 
         Caption         =   "0 byte"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   4320
         Width           =   4935
      End
      Begin VB.Label lblHostName 
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTimeName 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblLengthName 
         Caption         =   "Length"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblHost 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblTime 
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label lblLength 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label lblPortName 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblPort 
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   4935
      End
   End
   Begin MSComctlLib.TabStrip tspAttackResponse 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Last Response"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraResponseAnalysis 
      Caption         =   "Response Analysis"
      Height          =   5655
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   6375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenItem 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuContextSuggestions 
      Caption         =   "&Suggestions"
      Begin VB.Menu mnuSuggestionsExplainSuggestionItem 
         Caption         =   "&Explain suggestion"
      End
      Begin VB.Menu mnuSuggestionsContextSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuggestionsDeleteItem 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAttackResponseHelpItem 
         Caption         =   "&Attack Response Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmAttackResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 27
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    'Load the latest response
    Call LoadLatestResponse
    
    Call PrepareTabs
    
    Call SelectPluginTrigger
End Sub

Public Sub PrepareTabs()
    'Add the additional tabs if a lastresponse is given
    If LenB(LastResponse) <> 0 Then
        If tspAttackResponse.Tabs.Count < 3 Then
            Call tspAttackResponse.Tabs.Add(, , "&Suggestions")
            Call tspAttackResponse.Tabs.Add(, , "Silent &Checks")
        End If
        
        'Check the existence of the suggestions directory
        If application_suggestion_enable = True Then
            If (Dir$(application_suggestion_directory, 16) <> "") Then
                'Define the suggestions directory
                filSuggestions.Path = application_suggestion_directory
                
                'Compute the suggestion for further checks
                If lsvSuggestions.ListItems.Count = 0 Then
                    Call ComputeSuggestions
                End If
            End If
        End If
    
        If application_silent_checks_enable = True Then
            If lsvSilentPlugins.ListItems.Count = 0 Then
                Call DoSilentChecksScan
            End If
        End If
    End If
End Sub

Private Sub SelectPluginTrigger()
    'Select automaticly the trigger
    If LenB(LastResponse) > LenB(session_trigger_match) Then
        Dim Response() As String

        'Split the response apart
        Response = Split(LastResponse, session_trigger_match)

        'Select the trigger text
        txtLastResponse.SelStart = Len(Response(0))
        txtLastResponse.SelLength = Len(session_trigger_match)

        'Show the selected text
        Call ComputeSelection
    End If
End Sub

Private Sub DoSilentChecksScan()
    Dim i As Integer
    Dim List As ListItem
    Dim OriginalPluginFilename As String
    
    lsvSilentPlugins.ListItems.Clear
    
    'Save the original data of the check
    OriginalPluginFilename = plugin_filename
    
    'load the data into the ListView
    For i = 1 To frmMain.filATKPlugins.ListCount
        frmMain.filATKPlugins.ListIndex = i - 1
        
        If LenB(session_triggers) <> 0 Then
            If InStrB(1, LastResponse, session_trigger_match) > 0 Then
                Set List = lsvSilentPlugins.ListItems.Add(, , plugin_id)
                    List.SubItems(1) = plugin_name
                    List.SubItems(2) = bug_description
                    List.SubItems(3) = session_trigger_match
            End If
        End If
    Next i

    'Call the correct column width procedure
    LVColumnWidth lsvSilentPlugins

    'Select the original plugin
    Call ParseATKPlugin(ReadPluginFromFile(OriginalPluginFilename, application_plugin_directory))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLogEntry "Unload the " & Me.Caption, 6
    Set frmAttackResponse = Nothing
End Sub

Private Sub lsvSilentPlugins_Click()
    'This if is to prevent errors if there is no entry
    If lsvSilentPlugins.ListItems.Count > 0 Then
        Dim Response() As String
        
        'Select the text field
        txtLastResponse.SetFocus
        
        'Split the response apart
        Response = Split(LastResponse, lsvSilentPlugins.SelectedItem.SubItems(3))
    
        'Select the trigger text
        txtLastResponse.SelStart = Len(Response(0))
        txtLastResponse.SelLength = Len(lsvSilentPlugins.SelectedItem.SubItems(3))
    
        'Show the selected text
        Call ComputeSelection
    End If
End Sub

Private Sub lsvSilentPlugins_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewColumnReorder(frmAttackResponse.lsvSilentPlugins, ColumnHeader)
End Sub

Private Sub lsvSuggestions_Click()
    Dim Response() As String
    
    If lsvSuggestions.ListItems.Count > 0 Then
        'This is just a workaround
        On Error Resume Next
        
        'Select the text field
        txtLastResponse.SetFocus
        
        'Split the response apart
        Response = Split(LastResponse, lsvSuggestions.SelectedItem.SubItems(3))
    
        'Select the trigger text
        txtLastResponse.SelStart = Len(Response(0))
        txtLastResponse.SelLength = Len(lsvSuggestions.SelectedItem.SubItems(3))
    
        'Show the selected text
        Call ComputeSelection
        
        lblSuggestion.Caption = lsvSuggestions.SelectedItem.SubItems(1) & _
            vbNewLine & vbNewLine & lsvSuggestions.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub lsvSuggestions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show context menu if 2nd mouse button is pressed
    If Button = 2 Then
        PopupMenu mnuContextSuggestions
    End If
End Sub

Private Sub mnuFileCloseItem_Click()
    Unload Me
End Sub

Private Sub mnuFileOpenItem_Click()
    Dim ResponseFileName As String  'The name of the response file
    
    'Define the initial directory of the plugins
    cdgOpen.InitDir = application_response_directory
    
    'Ask the user for the desired filename
    cdgOpen.ShowOpen 'Opens the save dialog
    
    'Cache the filename into a variant to increase the speed
    ResponseFileName = cdgOpen.Filename
    
    'Check if a file was selected
    If LenB(ResponseFileName) <> 0 Then
        'Check if the file exists
        If (Dir$(ResponseFileName, 16) <> "") Then
            'Load the response file
            Call OpenResponseFile(ResponseFileName)
        End If
    End If
End Sub

Private Sub OpenResponseFile(ResponseFileName As String)
    Dim Temp As String          'Here we get the input from the opened file
    
    'Change the frame title
    Me.Caption = "Attack Response - File " & ResponseFileName
    
    'Erase the last response variant
    LastResponse = vbNullString
    
    'Open and read the plugin file
    Open ResponseFileName For Input As 1
        Do While Not EOF(1)
            Line Input #1, Temp
                LastResponse = LastResponse & Temp & vbNewLine
        Loop
    Close
        
    'Erase the last response in the textbox
    txtLastResponse.Text = vbNullString
    
    'Load the new last response
    Call LoadLatestResponse
    
    'Prepare the tabs
    Call PrepareTabs
End Sub

Private Sub mnuHelpAttackResponseHelpItem_Click()
    Call OpenOnlineHelp("attack_response")
End Sub

Private Sub mnuSuggestionsExplainSuggestionItem_Click()
    'Check if there is a suggestion
    If lsvSuggestions.ListItems.Count <> 0 Then
        MsgBox lsvSuggestions.SelectedItem.SubItems(1) & vbNewLine & vbNewLine & _
            lsvSuggestions.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub mnuSuggestionsDeleteItem_Click()
    'Delete the selected suggestion
    If lsvSuggestions.ListItems.Count <> 0 Then
        lsvSuggestions.ListItems.Remove (lsvSuggestions.SelectedItem.Index)
    End If
End Sub

Private Sub tspAttackResponse_Click()
    'Last Response
    If tspAttackResponse.SelectedItem.Index = 1 Then
        fraLastResponse.Visible = True
        fraSuggestions.Visible = False
        fraSilentChecks.Visible = False
        
        txtLastResponse.Top = 2280
        txtLastResponse.Height = 2775
    
        'Select the plugin trigger
        Call SelectPluginTrigger
        txtLastResponse.SetFocus
    
    'Suggestions
    ElseIf tspAttackResponse.SelectedItem.Index = 2 Then
        fraLastResponse.Visible = False
        fraSuggestions.Visible = True
        fraSilentChecks.Visible = False
    
        txtLastResponse.Top = 2280
        txtLastResponse.Height = 2055
    
    'Silent checks
    ElseIf tspAttackResponse.SelectedItem.Index = 3 Then
        fraLastResponse.Visible = False
        fraSuggestions.Visible = False
        fraSilentChecks.Visible = True
    
        txtLastResponse.Top = 2640
        txtLastResponse.Height = 2775
    End If
End Sub

Private Sub txtLastResponse_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ComputeSelection
End Sub

Private Sub txtLastResponse_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ComputeSelection
End Sub

Private Sub txtLastResponse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ComputeSelection
End Sub

Private Sub ComputeSelection()
    If txtLastResponse.SelLength > 0 Then
        lblCursorPosition.Caption = txtLastResponse.SelStart & " byte to " & _
            txtLastResponse.SelStart + txtLastResponse.SelLength & _
            " byte (selection length: " & txtLastResponse.SelLength & ")"
    Else
        lblCursorPosition.Caption = txtLastResponse.SelStart & " byte"
    End If
End Sub

Private Sub ComputeSuggestions()
    Dim i As Integer                   'Counter to check existing suggestions
    Dim j As Integer                   'Couter for adding new suggestions
    Dim strPatterns() As String        'The array for multiple patterns
    Dim intPatternCount As Integer     'The number of elements in the pattern array
    Dim intSuggestionsCount As Integer 'The suggestions
    Dim List As ListItem               'The listitem
    
    intSuggestionsCount = filSuggestions.ListCount
    
    'Delete the old suggestions
    lsvSuggestions.ListItems.Clear
    
    For i = 1 To intSuggestionsCount
        filSuggestions.ListIndex = i - 1
        Call ReadSuggestionFromFile(filSuggestions.Filename)
        
        'Split the multiple OR patterns
        strPatterns = Split(suggestion_trigger, " OR ")
        intPatternCount = UBound(strPatterns)

        'Check for the existence of one of the patterns
        For j = 0 To intPatternCount
            If InStr(1, LastResponse, strPatterns(j)) > 0 Then
                Set List = lsvSuggestions.ListItems.Add(, , suggestion_name, , 1)
                    List.SubItems(1) = suggestion_description
                    List.SubItems(2) = suggestion_todo
                    List.SubItems(3) = strPatterns(j)
            End If
        Next j
    Next i
End Sub
