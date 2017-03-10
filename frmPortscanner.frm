VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPortscanner 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Portscanner"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   4440
      Width           =   735
   End
   Begin VB.Frame fraEnumerationOptions 
      Caption         =   "Enumeration Options"
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   3375
      Begin VB.CheckBox chkDoActiveBannerGrabbing 
         Caption         =   "Do &active banner grabbing"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox chkDoPassiveBannerGrabbing 
         Caption         =   "Do &passive banner grabbing"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Value           =   1  'Checked
         Width           =   3135
      End
   End
   Begin VB.Frame fraScanStatus 
      Caption         =   "Scan Status"
      Height          =   975
      Left            =   3600
      TabIndex        =   14
      Top             =   3360
      Width           =   3375
      Begin MSComctlLib.ProgressBar pbrScanStatus 
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblScanStatus 
         Caption         =   "Waiting for input"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblScanStatusStatusName 
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblProgressName 
         Caption         =   "Progress"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
   End
   Begin MSWinsockLib.Winsock wskTCPWinsock 
      Index           =   0
      Left            =   4080
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   4440
      Width           =   735
   End
   Begin VB.Frame fraResult 
      Caption         =   "Result"
      Height          =   3135
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   3375
      Begin MSComctlLib.ListView lsvResults 
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
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
            Text            =   "Time"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Port"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Banner"
            Object.Width           =   1235
         EndProperty
      End
   End
   Begin VB.Frame fraPortscanOptions 
      Caption         =   "Portscan Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3375
      Begin VB.TextBox txtEndPort 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Text            =   "1023"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtStartPort 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   12
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtMaxSockets 
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "200"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblMaximumSocketsDefault 
         Caption         =   "(Default: 200)"
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblPortscanStartAndEndPortName 
         Caption         =   "Start and end port"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblMaximumSocketsName 
         Caption         =   "Maximum sockets"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraMappingOptions 
      Caption         =   "Mapping Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
      Begin VB.CheckBox chkScanIfPingFails 
         Caption         =   "Scan if ping &fails"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkDoICMPMapping 
         Caption         =   "&Do ICMP mapping (ICMP echo request)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmPortscanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDoActiveBannerGrabbing_Click()
    If chkDoPassiveBannerGrabbing.Value = 1 Then
        chkDoPassiveBannerGrabbing.Value = 0
    End If
End Sub

Private Sub chkDoPassiveBannerGrabbing_Click()
    If chkDoActiveBannerGrabbing.Value = 1 Then
        chkDoActiveBannerGrabbing.Value = 0
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdScan_Click()
    Dim Socket As Variant           'How can I set here a String Array?
    Dim CurrentPort As Integer
    Dim i As Integer
    Dim MaxSockets As Integer

    Dim List As ListItem

    'Needed if during a scan the frame is unloaded
    On Error Resume Next

    'Set the command to stop
    cmdScan.Enabled = False
    cmdStop.Enabled = True

    'We need a way to Start / Stop, so we'll use
    'the command button's caption as a reference
    If cmdStop.Enabled = True Then
    
        'Clear the last result
        lsvResults.ListItems.Clear
    
        'Reset the progress bar to zero
        pbrScanStatus.Value = 0
        
        'Lock all text boxes
        txtTarget.Enabled = False
        txtMaxSockets.Enabled = False
        txtStartPort.Enabled = False
        txtEndPort.Enabled = False
        
        'Read the maximum sockets
        MaxSockets = txtMaxSockets.Text
        
        ' Lets load some sockets to use
        For i = 1 To MaxSockets
            'Load new sock instance i
            Load wskTCPWinsock(i)
        Next i
        
        CurrentPort = txtStartPort.Text
        
        ' Again using the command1.caption as a reference
        ' to start / stop
        While cmdStop.Enabled = True
            For Each Socket In wskTCPWinsock
                ' Definately Need this so the system doesn't freeze
                DoEvents
                ' check if the socket is still trying to connect
                ' or is connected
                If Socket.State <> sckClosed Then
                    ' skip the increment of the port
                    GoTo continue
                End If
                ' close the socket to make double sure
                Socket.Close
                ' if it got to here, it's ready to try
                ' the next port, only after checking
                ' if we've done all the ports and the user
                ' hasn't clicked on Stop
                
                If CurrentPort > Val(txtEndPort.Text) + 1 Then
                    lblScanStatus.Caption = "Portscan finished"
                    'Lock free text boxes
                    txtTarget.Enabled = True
                    txtMaxSockets.Enabled = True
                    txtStartPort.Enabled = True
                    txtEndPort.Enabled = True
                    
                    cmdScan.Enabled = True
                    cmdStop.Enabled = False
                    Exit For
                End If
                'set the host
                Socket.RemoteHost = txtTarget.Text
                ' set the port
                Socket.RemotePort = CurrentPort
                
                lblScanStatus.Caption = "Scanning port " & CurrentPort
                pbrScanStatus.Value = pbrScanStatus.Value + _
                    ((txtEndPort.Text - txtStartPort.Text) / 100)
                
                ' attempt connect
                Socket.Connect
                ' fromhere, the socket will do one of two things
                ' 1) Raise a Connect therefore the port is open
                ' 2) Raise an Error therefore the port is closed
                
                ' increment the current port
                CurrentPort = CurrentPort + 1
    ' if the socketisn't ready to be incremented, go here
continue:
            
            ' goto the next socket instance
            Next Socket
        Wend
    Else ' command1.caption is "Stop"
        lblScanStatus.Caption = "Scan aborded"
        
        'Lock free text boxes
        txtTarget.Enabled = True
        txtMaxSockets.Enabled = True
        txtStartPort.Enabled = True
        txtEndPort.Enabled = True
    End If

    ' close all the sockets to save memory
    For i = 1 To MaxSockets
        Unload wskTCPWinsock(i)
    Next i

End Sub

Private Function AddPortToList(Port As Integer, Optional Banner As String)
'**************************************************
'* This is a function to add the port to the list *
'**************************************************

Dim List As ListItem

Set List = lsvResults.ListItems.Add(, , GetActualTime(":"))
    List.SubItems(1) = Port
    List.SubItems(2) = "open"
    List.SubItems(3) = Banner

    LVColumnWidth lsvResults

End Function

Private Sub cmdStop_Click()
    cmdScan.Enabled = True
    cmdStop.Enabled = False
End Sub

Private Sub Form_Load()
    txtTarget.Text = Target
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmPortscanner = Nothing
End Sub

Private Sub wskTCPWinsock_Connect(Index As Integer)
    ' the port is open so inform the user
    AddPortToList wskTCPWinsock(Index).RemotePort
    
    'Close the port immidiatly if no banner grabbing is wanted
    If chkDoPassiveBannerGrabbing.Value = 0 Then
        wskTCPWinsock(Index).Close
    ElseIf chkDoActiveBannerGrabbing.Value = 0 Then
        wskTCPWinsock(Index).Close
    End If
End Sub

Private Sub wskTCPWinsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim DataStr As String
    Dim i As Integer
    
    If chkDoPassiveBannerGrabbing.Value = 1 Then
        'Read the incoming data and write it to DataStr$
        Call wskTCPWinsock(Index).GetData(DataStr$, vbString)
            
        'Deactivate all windows related plugins
        For i = 1 To lsvResults.ListItems.Count
            If lsvResults.ListItems.Item(i).SubItems(1) = _
                wskTCPWinsock(Index).RemotePort Then
                                               
                lsvResults.ListItems.Remove (lsvResults.ListItems(i).Index)
                
            End If
        Next i
        
        AddPortToList wskTCPWinsock(Index).RemotePort, DataStr
    ElseIf chkDoActiveBannerGrabbing.Value = 1 Then
        wskTCPWinsock(Index).SendData (vbNewLine & vbNewLine)
        
        'Read the incoming data and write it to DataStr$
        Call wskTCPWinsock(Index).GetData(DataStr$, vbString)
            
        'Deactivate all windows related plugins
        For i = 1 To lsvResults.ListItems.Count
            If lsvResults.ListItems.Item(i).SubItems(1) = _
                wskTCPWinsock(Index).RemotePort Then
                                               
                lsvResults.ListItems.Remove (lsvResults.ListItems(i).Index)
                
            End If
        Next i
        
        AddPortToList wskTCPWinsock(Index).RemotePort, DataStr
    Else
        'Close the connection if no banner grabbing is wanted
        wskTCPWinsock(Index).Close
    End If
End Sub

Private Sub wskTCPWinsock_Error(Index As Integer, ByVal Number As Integer, _
    Description As String, ByVal Scode As Long, ByVal Source As String, _
    ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    ' the port is closed so close the socket so it
    ' will be incremented
    wskTCPWinsock(Index).Close
End Sub

