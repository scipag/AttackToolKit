VERSION 5.00
Begin VB.Form frmNslookup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "nslookup"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLookup 
      Caption         =   "&Lookup"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame fraResult 
      Caption         =   "Result"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.CommandButton cmdReverseLookupApply 
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdHostnameApply 
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdIPAddressApply 
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtReverseLookup 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtHostname 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   9
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtIPaddress 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblReverseLookupName 
         Caption         =   "Reverse lookup"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblHostNameName 
         Caption         =   "Host name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblIPAddressName 
         Caption         =   "IP address"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraHost 
      Caption         =   "Host"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtHost 
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
Attribute VB_Name = "frmNslookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHostnameApply_Click()
    txtHost.Text = txtHostname.Text
    Call cmdLookup_Click
End Sub

Private Sub cmdIPAddressApply_Click()
    txtHost.Text = txtIPaddress.Text
    Call cmdLookup_Click
End Sub

Private Sub cmdLookup_Click()
    Dim IPAddress As String
    Dim Hostname As String
    
    'Show the hourglass as cursor during checking
    Me.MousePointer = vbHourglass
    
    'Reset the result fields
    txtIPaddress.Text = vbNullString
    txtHostname.Text = vbNullString
    txtReverseLookup.Text = vbNullString
    
    'Compute the data and write the result
    IPAddress = GetIPFromHostName(txtHost.Text)
    Hostname = GetHostNameFromIP(IPAddress)
    
    If LenB(Hostname) Then
        txtHostname.Text = Hostname
    Else
        txtHostname.Text = txtHost.Text
    End If

    txtReverseLookup.Text = GetIPFromHostName(txtHostname.Text)
    txtIPaddress.Text = IPAddress

    'Show the normal cursor
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdReverseLookupApply_Click()
    txtHost.Text = txtReverseLookup.Text
    Call cmdLookup_Click
End Sub

Private Sub Form_Load()
    txtHost.Text = Target
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmNslookup = Nothing
End Sub
