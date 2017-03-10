VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPluginAutoUpdate 
   Caption         =   "Plugin AutoUpdate"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7635
   Icon            =   "frmPluginAutoUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame fraPlugins 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.Frame fraDownloadMessage 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   4455
         Begin MSComctlLib.ProgressBar pbrStatus 
            Height          =   135
            Left            =   1680
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   238
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblDownloadText 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Downloading ... Please wait!"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   4215
         End
         Begin VB.Shape shpStatus 
            BackStyle       =   1  'Opaque
            Height          =   495
            Left            =   0
            Top             =   360
            Width           =   4455
         End
         Begin VB.Shape shpRedLine 
            BackColor       =   &H00000080&
            BackStyle       =   1  'Opaque
            Height          =   735
            Left            =   0
            Top             =   240
            Width           =   4455
         End
      End
      Begin MSWinsockLib.Winsock wskPluginDownload 
         Index           =   0
         Left            =   6240
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "www.computec.ch"
         RemotePort      =   80
      End
      Begin MSWinsockLib.Winsock wskDownload 
         Index           =   0
         Left            =   6720
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "www.computec.ch"
         RemotePort      =   80
      End
      Begin MSComctlLib.ListView lsvPlugins 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7858
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   5
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
            Text            =   "Installed"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Available"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Action"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPluginsRefreshItem 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPluginsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsSearchPluginItem 
         Caption         =   "Sear&ch plugin"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuPluginsFindNextItem 
         Caption         =   "&Find next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuPluginsSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsSelectAllItem 
         Caption         =   "Select &all"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPluginsDeselectAllItem 
         Caption         =   "Dese&lect all"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuPluginsDownloadItem 
         Caption         =   "&Download"
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuPluginsSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginsShowItem 
         Caption         =   "&Show plugin entry in web browser"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpPluginAutoUpdateHelp 
         Caption         =   "&Plugin AutoUpdate Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmPluginAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2005-01-04                                                           *
' * - Fixed a bug in the pluginslist routine. The first item was not checked because *
' *   the checking started with 1 instead of 0.                                      *
' ************************************************************************************

Dim strNewPluginFileName As String
Dim strSearchText As String

Private Sub cmdClose_Click()
    WriteLogEntry "Closing the " & Me.Caption & " frame.", 6
    Unload Me
End Sub

Private Sub cmdDownload_Click()
    Dim i As Integer
    Dim intNewAvailablePlugins As Integer
    Dim Try As Integer
    Dim intNewPluginID As Integer
    
    Call FreezeFrame
    Call ReadText("Downloading new plugins ... Please wait!")
    
    fraDownloadMessage.Visible = True
    
    intNewAvailablePlugins = lsvPlugins.ListItems.Count
    
    For i = 1 To intNewAvailablePlugins
        'Delete the last download response
        LastResponse = vbNullString
        
        'On Error Resume Next
        Set lsvPlugins.SelectedItem = lsvPlugins.ListItems(i)
        If lsvPlugins.SelectedItem.Checked = True Then
            Try = 0
            strNewPluginFileName = lsvPlugins.SelectedItem.SubItems(1)
            intNewPluginID = lsvPlugins.SelectedItem.Text

            lblDownloadText.Caption = "Downloading plugin id " & intNewPluginID & _
                " (" & i & "/" & intNewAvailablePlugins & ")... Please wait!"
            WriteLogEntry "Downloading plugin id " & intNewPluginID & ". Please wait!", 6
            Call DownloadNewPlugin
            
            'Wait a few moments for a successful connection
            Do While wskPluginDownload(0).State = sckConnected
                If Try < application_attack_timeout * 0.5 Then
                    frmMain.Pause 1
                    Try = Try + 1000
                Else
                    WriteLogEntry "Downloading plugin id " & intNewPluginID & _
                        " timeout after " & application_attack_timeout & " milliseconds.", 3
                    Exit Do
                End If
            Loop
        End If
        SetProgress (100 / intNewAvailablePlugins) * i
    Next i
    
    lsvPlugins.ListItems.Clear
        
    WriteLogEntry "Plugin AutoUpdate complete. Ready!", 6
    Call ReadText("New plugins has been installed and are now ready to use.")
    Call ReleaseFrame
    
    MsgBox "Your local ATK plugin repository has been updated." & vbNewLine & vbNewLine & _
        "You are now able to run the latest ATK checks.", _
        vbOKOnly, "Attack Tool Kit Plugin AutoUpdate finished"
    
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call FreezeFrame
    SetProgress 0
    WriteLogEntry "Plugin AutoUpdate generate available plugin list. Please wait!", 6
    Call ReadText("Refreshing the list of available plugins.")
    Call GenerateActualATKPluginsList
        
    'Delete the last response
    LastResponse = vbNullString
    
    SetProgress 50
    
    Me.Caption = "Plugin AutoUpdate - " & application_plugin_download_url
    Call DownloadNewPluginsList

    WriteLogEntry "Plugin AutoUpdate available list download complete. Ready!", 6
    SetProgress 100
    Call ReadText("The list of available plugins has been refreshed.")
    Call ReleaseFrame
End Sub

Private Sub Form_Activate()
    If Me.WindowState = vbMinimized Then
        Me.WindowState = vbNormal
    End If
End Sub

Private Sub Form_Load()
    Call cmdRefresh_Click
End Sub

Private Sub DownloadNewPlugin()
    Dim strPluginDownloadRequestFileName As String
    Dim Try As Integer

    strPluginDownloadRequestFileName = Replace(application_plugin_download_url & strNewPluginFileName, " ", "%20")

    wskPluginDownload(0).Close
    wskPluginDownload(0).Connect GetDownloadHostname(), 80

    'Wait a few moments for a successful connection
    Do While wskPluginDownload(0).State <> sckConnected
        If Try < application_attack_timeout * 0.5 Then
            frmMain.Pause 1
            Try = Try + 1000
        Else
            Exit Do
        End If
    Loop

    If wskPluginDownload(0).State = 7 Then
        'Send the request with its needed command and linefeeds
        WriteLogEntry "Sending request for downloading plugin ...", 6
        wskPluginDownload(0).SendData "GET " & strPluginDownloadRequestFileName & " HTTP/1.0" & vbNewLine & vbNewLine
    End If
End Sub

Private Sub DownloadNewPluginsList()
    Dim Try As Integer

    wskDownload(0).Close
    wskDownload(0).Connect GetDownloadHostname(), 80

    'Wait a few moments for a successful connection
    Do While wskDownload(0).State <> sckConnected
        If Try < application_attack_timeout * 0.5 Then
            frmMain.Pause 1
            Try = Try + 1000
        Else
            Exit Do
        End If
    Loop

    If wskDownload(0).State = 7 Then
        'Send the request with its needed command and linefeeds
        WriteLogEntry "Sending request for downloading new plugins list ...", 6
        wskDownload(0).SendData "GET " & application_plugin_download_url & "pluginslist.txt HTTP/1.0" & vbNewLine & vbNewLine
    Else
        WriteLogEntry "No connection to the plugin repository server possible. Abording.", 3
        MsgBox "There could no connection to the plugin repository server " & vbNewLine & _
            wskDownload(0).RemoteHost & _
            " be established." & vbNewLine & vbNewLine & _
            "Please check the network settings and try again.", _
            vbOKOnly, "Attack Tool Kit Plugin AutoUpdate connection error"
            
        Unload Me
    End If
End Sub

Private Sub LoadNewPlugins()
    Dim m As Integer
    Dim n As Integer
    Dim List As ListItem        'Needed for the listview handling
    Dim intFreeFile1 As Integer
    Dim intFreeFile2 As Integer
    Dim strTempStringNew As String
    Dim strTempStringAvailable As String
    Dim ArrayNew() As String
    Dim ArrayAvailable() As String
    Dim intArrayNewItems As Integer
    Dim intArrayAvailableItems As Integer
    Dim TempArrayNew() As String
    Dim TempArrayAvailable() As String
    Dim bolPluginAvailable As Boolean
    
    SetProgress 0
    lsvPlugins.ListItems.Clear
    
    'Put the file data into the arrays
    If (Dir$(application_plugin_directory & "/newpluginslist.txt", 16) <> "") Then
        WriteLogEntry "Generate the file containing the new plugins...", 6
        intFreeFile1 = FreeFile
        Open application_plugin_directory & "/newpluginslist.txt" For Input As #intFreeFile1
            strTempStringNew = Input(LOF(intFreeFile1), #intFreeFile1)
        Close
        
        ArrayNew() = Split(strTempStringNew, vbNewLine, , vbBinaryCompare)
        
        intArrayNewItems = UBound(ArrayNew)
    Else
        WriteLogEntry "Could not create the file containing the new plugins...", 2
    End If
    
    If (Dir$(application_plugin_directory & "/pluginslist.txt", 16) <> "") Then
        WriteLogEntry "Generate the file containing the locally installed plugins", 6
        intFreeFile2 = FreeFile
        Open application_plugin_directory & "/pluginslist.txt" For Input As #intFreeFile2
                strTempStringAvailable = Input(LOF(intFreeFile2), #intFreeFile2)
        Close
        
        ArrayAvailable() = Split(strTempStringAvailable, vbNewLine, , vbBinaryCompare)
    
        intArrayAvailableItems = UBound(ArrayAvailable)
        
        If intArrayAvailableItems < 2 Then
            ArrayAvailable(1) = vbNullString
        End If
    Else
        WriteLogEntry "Could not create the file containing the locally installed plugins.", 2
    End If
    
    WriteLogEntry "Comparing the local and remote plugin repository...", 6
    For m = 0 To intArrayNewItems
        If InStrB(1, ArrayNew(m), ";", vbBinaryCompare) Then
            
            'Split the data to be written
            TempArrayNew = Split(ArrayNew(m), ";")

            For n = 0 To intArrayAvailableItems
                If InStrB(1, ArrayAvailable(n), ";", vbBinaryCompare) Then
                    
                    'Split the data to be written
                    TempArrayAvailable = Split(ArrayAvailable(n), ";")
                        
                    bolPluginAvailable = False
    
                    If TempArrayAvailable(0) = TempArrayNew(0) Then
                        bolPluginAvailable = True
                        If TempArrayAvailable(2) <> TempArrayNew(2) Then
                            Set List = lsvPlugins.ListItems.Add(, , TempArrayAvailable(0))
                                List.SubItems(1) = TempArrayAvailable(1)
                                List.SubItems(2) = TempArrayAvailable(2) & " (" & TempArrayAvailable(3) & ")"
                                List.SubItems(3) = TempArrayNew(2) & " (" & TempArrayNew(3) & ")"
                                List.SubItems(4) = "Update (" & TempArrayNew(4) & " bytes)"
                        End If
                        Exit For
                    End If
                End If
            Next n
            
            'Write the plugin if the plugin is new
            If bolPluginAvailable = False Then
                Set List = lsvPlugins.ListItems.Add(, , TempArrayNew(0))
                    List.SubItems(1) = TempArrayNew(1)
                    List.SubItems(2) = "N/A"
                    List.SubItems(3) = TempArrayNew(2) & " (" & TempArrayNew(3) & ")"
                    List.SubItems(4) = "Install (" & TempArrayNew(4) & " bytes)"
            End If

        End If
        
        SetProgress (100 / intArrayNewItems) * m
    Next m
    
    fraPlugins.Caption = lsvPlugins.ListItems.Count & " new plugins are available"
    
    'Set the right column width
    LVColumnWidth lsvPlugins
    
    SetProgress 100
    
    If lsvPlugins.ListItems.Count = 0 Then
        WriteLogEntry "No new plugins are available for download.", 5
        MsgBox "No new plugins available.", _
            vbOKOnly, "Attack Tool Kit Plugin AutoUpdate information"
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        
        'Prevent zu small windows in width
        If Me.Width < 6000 Then
            Me.Width = 6000
        End If
        
        fraPlugins.Height = Me.Height - 1460
        lsvPlugins.Height = fraPlugins.Height - 360
        
        fraPlugins.Width = Me.Width - 360
        lsvPlugins.Width = fraPlugins.Width - 240
        
        cmdClose.Left = fraPlugins.Width - 980
        cmdDownload.Left = cmdClose.Left - cmdClose.Width - 120
        cmdRefresh.Left = cmdDownload.Left - cmdDownload.Width - 120
        
        cmdClose.Top = Me.Height - 1200
        cmdDownload.Top = cmdClose.Top
        cmdRefresh.Top = cmdClose.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLogEntry "Unload the " & Me.Caption, 6
    Set frmPluginAutoUpdate = Nothing
End Sub

Private Sub lsvPlugins_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewColumnReorder(frmPluginAutoUpdate.lsvPlugins, ColumnHeader)
End Sub

Private Sub lsvPlugins_DblClick()
    If lsvPlugins.ListItems.Count Then
        Dim WebSiteURL As String
        
        WebSiteURL = application_plugin_download_url & lsvPlugins.SelectedItem.SubItems(1) & ".html"
        
        'Load the project web site
        WriteLogEntry "Loading the plugin website " & WebSiteURL, 6
        Call ShellExecute(Me.hwnd, "Open", WebSiteURL, "", App.Path, 1)
    Else
        MsgBox "No new plugins available." & vbNewLine & vbNewLine & _
            "Please hit the reload button to refresh the list of loadable new plugins.", _
            vbOKOnly, "Attack Tool Kit Plugin AutoUpdate error"
        Call cmdRefresh_Click
    End If
End Sub

Private Sub lsvPlugins_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show context menu if 2nd mouse button is pressed
    If Button = vbRightButton Then
        PopupMenu mnuPlugins
    End If
End Sub

Private Sub mnuFileCloseItem_Click()
    Call cmdClose_Click
End Sub

Private Sub mnuHelpPluginAutoUpdateHelp_Click()
    Call OpenOnlineHelp("plugin_autoupdate")
End Sub

Private Sub mnuPluginsDeselectAllItem_Click()
    Dim i As Integer
    
    'Deselect all loadable plugins
    For i = 1 To lsvPlugins.ListItems.Count
        lsvPlugins.ListItems.Item(i).Checked = False
    Next i
End Sub

Private Sub mnuPluginsDownloadItem_Click()
    Call cmdDownload_Click
End Sub

Private Sub mnuPluginsFindNextItem_Click()
    Call SearchPlugin
End Sub

Private Sub mnuPluginsRefreshItem_Click()
    Call cmdRefresh_Click
End Sub

Private Sub mnuPluginsSearchPluginItem_Click()
    'Define a default search string if this is the first search
    If LenB(strSearchText) = 0 Then
        strSearchText = "Apache prior 2.0"
    End If
    
    'Ask for the search string
    strSearchText = InputBox("Please enter string you are searching for. " & _
        "(e.g. Microsoft, Apache, Sendmail).", _
        "Plugin AutoUpdate plugin search", strSearchText)
    
    'Start the search
    Call SearchPlugin
End Sub

Private Sub mnuPluginsSelectAllItem_Click()
    Dim i As Integer                    'This i is used for the counters
    Dim intShownPlugins As Integer        'How many plugins are loaded

    intShownPlugins = lsvPlugins.ListItems.Count
    
    'Check if there is one or more checks activated for the audit
    For i = 1 To intShownPlugins
        lsvPlugins.ListItems.Item(i).Checked = True
    Next i
End Sub

Private Sub SelectNewPlugins()
    Dim i As Integer                    'This i is used for the counters
    Dim intShownPlugins As Integer        'How many plugins are loaded

    intShownPlugins = lsvPlugins.ListItems.Count
    
    'Check if there is one or more checks activated for the audit
    For i = 1 To intShownPlugins
        If Not lsvPlugins.ListItems.Item(i).SubItems(2) = lsvPlugins.ListItems.Item(i).SubItems(3) Then
            lsvPlugins.ListItems.Item(i).Checked = True
        End If
    Next i
End Sub

Private Sub mnuPluginsShowItem_Click()
    Call lsvPlugins_DblClick
End Sub

Private Sub wskDownload_Close(Index As Integer)
    Call WriteNewPluginsListToFile
    Call LoadNewPlugins
    Call SelectNewPlugins

    wskDownload(0).Close
End Sub

Private Sub WriteNewPluginsListToFile()
    Dim intFreeFile As Integer
    
    intFreeFile = FreeFile
    
    'Strip the http header
    LastResponse = Mid$(LastResponse, InStr(1, LastResponse, vbNewLine & vbNewLine) + 4)
    
    'Replace the Linefeeds
    If InStrB(1, LastResponse, vbLf, vbBinaryCompare) Then
        LastResponse = Replace$(LastResponse, vbLf, vbNewLine)
    End If
    
    On Error Resume Next ' Needed if there are no write permissions
    Open application_plugin_directory & "\newpluginslist.txt" For Output As #intFreeFile
        Print #intFreeFile, LastResponse
    Close
End Sub

Private Sub WriteNewPluginToFile()
    Dim intFreeFile As Integer
    
    intFreeFile = FreeFile
    
    'Strip the http header
    LastResponse = Mid$(LastResponse, InStr(1, LastResponse, vbNewLine & vbNewLine) + 4)
    
    'Replace the Linefeeds
    If InStrB(1, LastResponse, vbLf, vbBinaryCompare) Then
        LastResponse = Replace$(LastResponse, vbLf, vbNewLine)
    End If
    
    On Error Resume Next ' Needed if there are no write permissions
    Open application_plugin_directory & "\" & strNewPluginFileName For Output As #intFreeFile
        Print #intFreeFile, LastResponse
    Close
End Sub

Private Sub wskDownload_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Here is the incoming data cached
    Dim DataStr As String
    
    'Read the incoming data and write it to DataStr$
    Call wskDownload(0).GetData(DataStr$, vbString)

    LastResponse = LastResponse & DataStr
End Sub

Private Sub wskDownload_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLogEntry "WinSock Error: [" & Number & "] " & Description, 1
    
    Call wskDownload_Close(0)
End Sub

Private Function GetDownloadHostname() As String
    Dim strTempArray() As String
    Dim strPluginsDownloadURLTemp As String
    
    strPluginsDownloadURLTemp = application_plugin_download_url
    
    'Strip http://
    If InStr(1, strPluginsDownloadURLTemp, "http://") Then
        strPluginsDownloadURLTemp = Mid$(strPluginsDownloadURLTemp, 8, Len(strPluginsDownloadURLTemp))
    End If
    
    'Strip all slashes from the URL
    If InStr(1, strPluginsDownloadURLTemp, "/") Then
        strTempArray = Split(strPluginsDownloadURLTemp, "/")
        strPluginsDownloadURLTemp = strTempArray(0)
    End If
    
    'Strip all back slashes from the URL
    If InStr(1, strPluginsDownloadURLTemp, "\") Then
        strTempArray = Split(strPluginsDownloadURLTemp, "\")
        strPluginsDownloadURLTemp = strTempArray(0)
    End If
    
    GetDownloadHostname = strPluginsDownloadURLTemp
End Function

Private Sub wskPluginDownload_Close(Index As Integer)
    Call WriteNewPluginToFile
    
    wskPluginDownload(0).Close
End Sub

Private Sub wskPluginDownload_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Here is the incoming data cached
    Dim DataStr As String
    
    'Read the incoming data and write it to DataStr$
    Call wskPluginDownload(0).GetData(DataStr$, vbString)

    LastResponse = LastResponse & DataStr
End Sub

Private Sub wskPluginDownload_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WriteLogEntry "WinSock Error: [" & Number & "] " & Description, 1
    
    Call wskPluginDownload_Close(0)
End Sub

Private Sub FreezeFrame()
    fraDownloadMessage.Visible = True
    Screen.MousePointer = 13
    lsvPlugins.Enabled = False
    cmdRefresh.Enabled = False
    cmdDownload.Enabled = False
    mnuPluginsSearchPluginItem.Enabled = False
    mnuPluginsFindNextItem.Enabled = False
    mnuPluginsRefreshItem.Enabled = False
    mnuPluginsDeselectAllItem.Enabled = False
    mnuPluginsSelectAllItem.Enabled = False
    mnuPluginsShowItem.Enabled = False
    mnuPluginsDownloadItem.Enabled = False
    DoEvents
End Sub

Private Sub ReleaseFrame()
    fraDownloadMessage.Visible = False
    Screen.MousePointer = 0
    lsvPlugins.Enabled = True
    cmdRefresh.Enabled = True
    cmdDownload.Enabled = True
    mnuPluginsSearchPluginItem.Enabled = True
    mnuPluginsFindNextItem.Enabled = True
    mnuPluginsRefreshItem.Enabled = True
    mnuPluginsDeselectAllItem.Enabled = True
    mnuPluginsSelectAllItem.Enabled = True
    mnuPluginsShowItem.Enabled = True
    mnuPluginsDownloadItem.Enabled = True
End Sub

Public Sub SetProgress(ByRef intValue As Integer)
    'Prevent too large values (this is just a nasty workaround!)
    If intValue > 100 Then
        intValue = 100
    End If
    
    pbrStatus.Value = intValue
    frmMain.StatusBar.Panels(2).Text = intValue & " %"
    frmMain.pbrProgress.Value = intValue
End Sub

Private Sub SearchPlugin()
    Dim intListItemStartPosition As Integer
    Dim intListItemCount As Integer
    Dim i As Integer
    
    WriteLogEntry "Starting the search for the string " & strSearchText, 1
    
    intListItemCount = lsvPlugins.ListItems.Count
    
    If lsvPlugins.SelectedItem.Index < intListItemCount Then
        intListItemStartPosition = lsvPlugins.SelectedItem.Index + 1
    Else
        intListItemStartPosition = 1
    End If
    
    For i = intListItemStartPosition To intListItemCount
        If InStrB(1, _
            LCase$(lsvPlugins.ListItems.Item(i).SubItems(1)), _
            LCase$(strSearchText), vbBinaryCompare) Then
            
            Set lsvPlugins.SelectedItem = lsvPlugins.ListItems(i)
            lsvPlugins.SetFocus
            lsvPlugins.SelectedItem.EnsureVisible
            Exit For
        End If
    Next i
End Sub

