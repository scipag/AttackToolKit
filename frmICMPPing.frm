VERSION 5.00
Begin VB.Form frmICMPPing 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ICMP Ping"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPing 
      Caption         =   "&Ping"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame fraResult 
      Caption         =   "Result"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.Label lblDatasize 
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblRoundTripTime 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblMessage 
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblPingStatus 
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblIPAddress 
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblDatasizeName 
         Caption         =   "Datasize"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblMessageName 
         Caption         =   "Message"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblPingStatusName 
         Caption         =   "Ping Status"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblRTTName 
         Caption         =   "Round Trip Time (rtt)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblIPAddressName 
         Caption         =   "IP address"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   120
         MaxLength       =   255
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmICMPPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPing_Click()
    Dim ECHO As ICMP_ECHO_REPLY
    Dim TargetIP As String
    
    'Reset the last result
    lblIPAddress.Caption = vbNullString
    lblPingStatus.Caption = vbNullString
    lblMessage.Caption = vbNullString
    lblRoundTripTime.Caption = vbNullString
    lblDatasize.Caption = vbNullString

    'Do a nslookup of the target
    TargetIP = GetIPFromHostName(txtTarget.Text)

    'If nslookup is possible do ping
    If Len(TargetIP) > 0 Then
        Call Ping(TargetIP, ECHO)
        lblIPAddress.Caption = TargetIP
        
        'Print status number
        lblPingStatus.Caption = ECHO.status
        If ECHO.status = 0 Then
            lblMessage.Caption = "Successful"
            lblRoundTripTime.Caption = ECHO.RoundTripTime & " ms"
        Else
            lblMessage.Caption = "Not successful"
        End If
        
        lblDatasize.Caption = ECHO.DataSize & " bytes"
    Else
        lblMessage.Caption = "Not successful"
    End If
               
End Sub

Private Sub Form_Load()
    txtTarget.Text = Target
End Sub
