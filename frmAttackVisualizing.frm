VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAttackVisualizing 
   Caption         =   "Attack Visualizing"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6270
   Icon            =   "frmAttackVisualizing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraIllustrated 
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   5775
      Begin VB.TextBox txtNetworkData 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmAttackVisualizing.frx":0CCA
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox txtAttackerData 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   735
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmAttackVisualizing.frx":0D4A
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtTargetData 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   735
         Left            =   3120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmAttackVisualizing.frx":0D56
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblAttackerComputing 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "!"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTargetName 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         MouseIcon       =   "frmAttackVisualizing.frx":0D58
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTargetComputing 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "!"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblNetworkName 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Network"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         MouseIcon       =   "frmAttackVisualizing.frx":1062
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblAttackerName 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Attacker"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Shape shpTarget 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         Height          =   735
         Left            =   4320
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape shpNetwork 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   2040
         Shape           =   2  'Oval
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape shpAttacker 
         BackColor       =   &H00000040&
         BackStyle       =   1  'Opaque
         Height          =   735
         Left            =   120
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line linNetwork 
         X1              =   2880
         X2              =   2880
         Y1              =   1200
         Y2              =   2160
      End
      Begin VB.Line Line2 
         X1              =   5400
         X2              =   5400
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   360
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Line linArrow4LineC 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1560
         X2              =   1440
         Y1              =   960
         Y2              =   840
      End
      Begin VB.Line linArrow4LineB 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1560
         X2              =   1440
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Line linArrow4LineA 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   2040
         X2              =   1440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line linArrow3LineB 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   3840
         X2              =   3720
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Line linArrow3LineC 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   3720
         X2              =   3840
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Line linArrow3LineA 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   4320
         X2              =   3720
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line linArrow2LineC 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   4200
         X2              =   4320
         Y1              =   720
         Y2              =   600
      End
      Begin VB.Line linArrow2LineB 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   4200
         X2              =   4320
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line linArrow1LineC 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1920
         X2              =   2040
         Y1              =   720
         Y2              =   600
      End
      Begin VB.Line linArrow1LineB 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1920
         X2              =   2040
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line linArrow1LineA 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1440
         X2              =   2040
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line linArrow2LineA 
         BorderColor     =   &H00000040&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   3720
         X2              =   4320
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame fraListing 
      Height          =   5055
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin MSComctlLib.ListView lsvListing 
         Height          =   4695
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
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
            Text            =   "Source"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Destination"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Data"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tspVisualizing 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9975
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Illustrated"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Listing"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCloseItem 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAttackVisualizingHelpItem 
         Caption         =   "&Attack Visualizing Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmAttackVisualizing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2004-12-28                                                           *
' * - Optimized the visualizing routines.                                            *
' * - Added the source and destination ports in the listing.                         *
' * - Added the printing of the pattern during pattern matching.                     *
' * - Fixed a bug with a listing misorder (pattern matching after vulnerability      *
' *   found.                                                                         *
' * - Fixed a source/destination misorder during data receiving.                     *
' * Version 3.0 2004-11-14                                                           *
' * - Added the resizing possibility of the frame.                                   *
' * Version 3.0 2004-11-04                                                           *
' * - Replaced the keypress events with the default/cancel properties.               *
' * Version 3.0 2004-11-01                                                           *
' * - Replaced all useless functions with normal subs.                               *
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
    Dim strLocalIP As String
    
    strLocalIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    
    txtAttackerData.Text = strLocalIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")"
    lblAttackerName.ToolTipText = strLocalIP
    
    txtTargetData.Text = Target
    lblTargetName.ToolTipText = Target

    If InStrB(1, Target, "192.168.", vbBinaryCompare) Then
        lblNetworkName.Caption = "LAN Class C"
        lblNetworkName.ToolTipText = "192.168.0.0 - 192.168.255.255"
    ElseIf InStrB(1, Target, "172.", vbBinaryCompare) Then
        lblNetworkName.Caption = "LAN Class B"
        lblNetworkName.ToolTipText = "172.16.0.0 - 172.31.255.255"
    ElseIf InStrB(1, Target, "10.", vbBinaryCompare) Then
        lblNetworkName.Caption = "LAN Class A"
        lblNetworkName.ToolTipText = "10.0.0.0 - 10.255.255.255"
    ElseIf InStrB(1, Target, "127.", vbBinaryCompare) Then
        lblNetworkName.Caption = "Localhost"
        lblNetworkName.ToolTipText = "127.0.0.0 - 127.255.255.255"
    Else
        lblNetworkName.Caption = "Internet"
        lblNetworkName.ToolTipText = "0.0.0.0 - 255.255.255.255"
    End If

End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        'Prevent zu small windows in height
        If Me.Height < 6585 Then
            Me.Height = 6585
        End If
        
        'Prevent zu small windows in width
        If Me.Width < 6390 Then
            Me.Width = 6390
        End If
        
        tspVisualizing.Height = Me.Height - 920
        tspVisualizing.Width = Me.Width - 360
        
        fraIllustrated.Width = tspVisualizing.Width - 240
        fraIllustrated.Height = tspVisualizing.Height - 600
        
        fraListing.Width = fraIllustrated.Width
        fraListing.Height = fraIllustrated.Height
        lsvListing.Width = fraListing.Width - 240
        lsvListing.Height = fraListing.Height - 360
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAttackVisualizing = Nothing
End Sub

Private Sub lblNetworkName_Click()
    Dim WebSiteURL As String
    
    WebSiteURL = "http://www.faqs.org/rfcs/rfc1918.html"
    
    'Load the project web site
    WriteLogEntry "Loading the website " & WebSiteURL, 6
    Call ShellExecute(Me.hwnd, "Open", WebSiteURL, "", App.Path, 1)
End Sub

Private Sub lblTargetName_Click()
    Shell Environ("Comspec") + " /C telnet " & Target & " " & plugin_port, vbNormalFocus
End Sub

Private Sub lsvListing_KeyPress(KeyAscii As Integer)
    If KeyAscii = "13" Then
        If lsvListing.ListItems.Count Then
            Clipboard.Clear
            Clipboard.SetText lsvListing.SelectedItem.SubItems(1) & ";" & _
            lsvListing.SelectedItem.SubItems(2) & ";" & _
            lsvListing.SelectedItem.SubItems(3) & ";" & _
            lsvListing.SelectedItem.SubItems(4) & ";" & _
            lsvListing.SelectedItem.SubItems(5), vbCFText
        End If
        Unload Me
    End If
End Sub

Private Sub mnuFileCloseItem_Click()
    Unload Me
End Sub

Private Sub mnuHelpAttackVisualizingHelpItem_Click()
    Call OpenOnlineHelp("attack_visualizing")
End Sub

Private Sub tspVisualizing_Click()
    Dim intSelectedItem As Integer
    
    intSelectedItem = tspVisualizing.SelectedItem.Index

    If intSelectedItem = 1 Then
        fraIllustrated.Visible = True
        fraListing.Visible = False
    ElseIf intSelectedItem = 2 Then
        fraListing.Visible = True
        fraIllustrated.Visible = False
    End If
End Sub

Private Sub txtAttackerData_KeyPress(KeyAscii As Integer)
    'If the user presses enter, the selected field is copied to the clipboard
    If KeyAscii = "13" Then
        Clipboard.Clear
        txtAttackerData.SelStart = 0
        txtAttackerData.SelLength = Len(txtAttackerData.Text)
        Clipboard.SetText txtAttackerData.SelText, vbCFText
        Unload Me
    Else
        TextBoxSetFocus KeyAscii, "1"
    End If
End Sub

Private Sub txtNetworkData_KeyPress(KeyAscii As Integer)
    'If the user presses enter, the selected field is copied to the clipboard
    If KeyAscii = "13" Then
        Clipboard.Clear
        txtNetworkData.SelStart = 0
        txtNetworkData.SelLength = Len(txtNetworkData.Text)
        Clipboard.SetText txtNetworkData.SelText, vbCFText
        Unload Me
    Else
        TextBoxSetFocus KeyAscii, "3"
    End If
End Sub

Private Sub txtTargetData_KeyPress(KeyAscii As Integer)
    'If the user presses enter, the selected field is copied to the clipboard
    If KeyAscii = "13" Then
        Clipboard.Clear
        txtTargetData.SelStart = 0
        txtTargetData.SelLength = Len(txtTargetData.Text)
        Clipboard.SetText txtTargetData.SelText, vbCFText
        Unload Me
    Else
        TextBoxSetFocus KeyAscii, "2"
    End If
End Sub

Private Sub TextBoxSetFocus(ByRef KeyAscii As Integer, ByRef BoxNumber As Integer)
    '1
    If KeyAscii = "49" Then
        txtAttackerData.SetFocus
    'a
    ElseIf KeyAscii = "97" Then
        txtAttackerData.SetFocus
    '2
    ElseIf KeyAscii = "50" Then
        txtTargetData.SetFocus
    't
    ElseIf KeyAscii = "116" Then
        txtTargetData.SetFocus
    '3
    ElseIf KeyAscii = "51" Then
        txtNetworkData.SetFocus
    'n
    ElseIf KeyAscii = "110" Then
        txtNetworkData.SetFocus
    '+
    ElseIf KeyAscii = "43" Then
        If BoxNumber = 1 Then
            txtTargetData.SetFocus
        ElseIf BoxNumber = 2 Then
            txtNetworkData.SetFocus
        ElseIf BoxNumber = 3 Then
            txtAttackerData.SetFocus
        End If
    '-
    ElseIf KeyAscii = "45" Then
        If BoxNumber = 1 Then
            txtNetworkData.SetFocus
        ElseIf BoxNumber = 2 Then
            txtAttackerData.SetFocus
        ElseIf BoxNumber = 3 Then
            txtTargetData.SetFocus
        End If
    End If
End Sub

Public Sub VisualizeOpenConnection()
    Dim List As ListItem        'Needed for the listview handling
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort
    
    lblAttackerComputing.Visible = True
    lblTargetComputing.Visible = False
    
    linArrow1LineA.Visible = True
    linArrow1LineB.Visible = True
    linArrow1LineC.Visible = True

    linArrow2LineA.Visible = True
    linArrow2LineB.Visible = True
    linArrow2LineC.Visible = True
    
    lblAttackerComputing.Visible = False
    lblTargetComputing.Visible = True
    
    linArrow3LineA.Visible = False
    linArrow3LineB.Visible = False
    linArrow3LineC.Visible = False
    
    linArrow4LineA.Visible = False
    linArrow4LineB.Visible = False
    linArrow4LineC.Visible = False
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "Opening socket ..."

    txtTargetData.Text = strDestinationIP & vbNewLine & _
        "Receiving connection request ..."
        
    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, "", "Opening connection")
End Sub

Public Sub VisualizeCloseConnection()
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort
    
    lblAttackerComputing.Visible = True
    lblTargetComputing.Visible = True

    linArrow1LineA.Visible = False
    linArrow1LineB.Visible = False
    linArrow1LineC.Visible = False
    
    linArrow2LineA.Visible = False
    linArrow2LineB.Visible = False
    linArrow2LineC.Visible = False
    
    linArrow3LineA.Visible = False
    linArrow3LineB.Visible = False
    linArrow3LineC.Visible = False
    
    linArrow4LineA.Visible = False
    linArrow4LineB.Visible = False
    linArrow4LineC.Visible = False
    
    lblAttackerComputing.Visible = False
    lblTargetComputing.Visible = False
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "Closing socket ..."
    
    txtTargetData.Text = strDestinationIP & " (" & Target & ")" & vbNewLine & _
        "Session terminated. Waiting for next connection."

    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, vbNullString, "Closing connection")
End Sub

Public Sub VisualizeSendData(ByRef strDataToSend As String)
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "Sending data ..."

    txtNetworkData.Text = GetActualTime(":") & " " & "Attacker (" & strSourceIP & _
        ") -> Target (" & strDestinationIP & ")" & vbNewLine & vbNewLine & _
        strDataToSend
        
    txtTargetData.Text = strDestinationIP & " (" & strDestinationIP & ")" & vbNewLine & _
        "Receiving data ..."
    
    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, strDataToSend, "Sending data")
End Sub

Public Sub VisualizeSleep(ByRef intSleepTime As Integer)
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort

    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "Waiting " & intSleepTime & " seconds ..."
        
    lblAttackerComputing.Visible = True
    lblTargetComputing.Visible = False
    
    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, vbNullString, "Sleep for " & intSleepTime & " seconds")
End Sub

Public Sub VisualizePatternExists(ByRef strTriggers As String)
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort
    
    lblAttackerComputing.Visible = True
    lblTargetComputing.Visible = False
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "Checking if the pattern exists ..."

    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, "" & strTriggers & "", "Check if the pattern exists")
End Sub

Public Sub VisualizeVulnerabilityFound()
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort
    
    lblAttackerComputing.Visible = False
    lblTargetComputing.Visible = False
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "The vulnerability was found. Waiting for input."

    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, vbNullString, "Vulnerability was found")
End Sub

Public Sub VisualizeVulnerabilityNotFound()
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intSourcePort = frmMain.wskTCPWinsock.Item(0).LocalPort
    strDestinationIP = Target
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).RemotePort
    
    lblAttackerComputing.Visible = False
    lblTargetComputing.Visible = False
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "The vulnerability was not found. Waiting for input."

    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, vbNullString, "Vulnerability was not found")
End Sub

Public Sub VisualizeDataArrival()
    Dim strSourceIP As String
    Dim intSourcePort As Integer
    Dim strDestinationIP As String
    Dim intDestinationPort As Integer
    
    strSourceIP = Target
    intSourcePort = frmMain.wskTCPWinsock.Item(0).RemotePort
    strDestinationIP = frmMain.wskTCPWinsock.Item(0).LocalIP
    intDestinationPort = frmMain.wskTCPWinsock.Item(0).LocalPort
    
    lblAttackerComputing.Visible = False
    lblTargetComputing.Visible = False
    
    linArrow1LineA.Visible = False
    linArrow1LineB.Visible = False
    linArrow1LineC.Visible = False
    
    linArrow2LineA.Visible = False
    linArrow2LineB.Visible = False
    linArrow2LineC.Visible = False
    
    linArrow3LineA.Visible = True
    linArrow3LineB.Visible = True
    linArrow3LineC.Visible = True
    
    linArrow4LineA.Visible = True
    linArrow4LineB.Visible = True
    linArrow4LineC.Visible = True
    
    txtTargetData.Text = strDestinationIP & " (" & strSourceIP & ")" & vbNewLine & _
        "Sending data back ..."
    
    txtAttackerData.Text = strSourceIP & " (" & frmMain.wskTCPWinsock.Item(0).LocalHostName & ")" & vbNewLine & _
        "Receiving data ..."
    
    txtNetworkData.Text = LastResponseTime & " " & "Target (" & strDestinationIP & ") -> Attacker (" & strSourceIP & ")" & _
        vbNewLine & vbNewLine & _
        LastResponse
    
    Call WriteDataToListView(GetTodaysDate("/"), GetActualTime(":"), strSourceIP & ":" & intSourcePort, _
        strDestinationIP & ":" & intDestinationPort, LastResponse, "Incoming data")
End Sub

Public Sub WriteDataToListView(ByRef strDate As String, _
                                ByRef strTime As String, _
                                ByRef strSource As String, _
                                ByRef strDestination As String, _
                                ByRef strData As String, _
                                ByRef strDescription As String)

    Dim List As ListItem        'Needed for the listview handling

    'Set the frame title
    Me.Caption = "Attack Visualizing - " & plugin_name & ": " & strDescription

    'Write the log data into the log frame
    Set List = lsvListing.ListItems.Add(, , strDate)
        List.SubItems(1) = strTime
        List.SubItems(2) = strSource
        List.SubItems(3) = strDestination
        List.SubItems(4) = strData
        List.SubItems(5) = strDescription
    
    'Set the right column width
    LVColumnWidth lsvListing
End Sub
