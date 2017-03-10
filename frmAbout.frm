VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attack Tool Kit About"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSymbol 
      BorderStyle     =   0  'None
      Caption         =   "Symbol"
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2655
      Begin VB.Label lblPaintingCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "painting by martin anner"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000060&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmAbout.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   5130
         Left            =   0
         MouseIcon       =   "frmAbout.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":0614
         Top             =   0
         Width           =   2700
      End
   End
   Begin VB.Frame fraFrame 
      Height          =   5175
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   3735
      Begin VB.Frame fraDonate 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Image imgPayPalDonation 
            Height          =   465
            Left            =   1200
            MouseIcon       =   "frmAbout.frx":2374
            MousePointer    =   99  'Custom
            Picture         =   "frmAbout.frx":267E
            ToolTipText     =   "Click here to visit the project web site to make a donation"
            Top             =   1800
            Width           =   930
         End
         Begin VB.Label lblDonateText 
            Caption         =   $"frmAbout.frx":29E3
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Frame fraAboutInfo 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3495
         Begin VB.Label lblDescription 
            Caption         =   "The Attack Tool Kit (ATK) is an open-source utility to realize vulnerability checks and enhance security audits. "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   3375
         End
         Begin VB.Label lblCopyrights 
            Caption         =   "The Attack Tool Kit (ATK) is developed and maintained (2003-2005) by"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   18
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            Caption         =   "Marc Ruef"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "marc.ruef@computec.ch"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   0
            MouseIcon       =   "frmAbout.frx":2AF4
            MousePointer    =   99  'Custom
            TabIndex        =   16
            ToolTipText     =   "Write me an email"
            Top             =   1800
            Width           =   2070
         End
         Begin VB.Label lblWeb 
            AutoSize        =   -1  'True
            Caption         =   "http://www.computec.ch"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   0
            MouseIcon       =   "frmAbout.frx":2DFE
            MousePointer    =   99  'Custom
            TabIndex        =   15
            ToolTipText     =   "Visit my web site"
            Top             =   2040
            Width           =   2100
         End
      End
      Begin VB.Frame fraProjectWebsite 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   3495
         Begin VB.Label lblProjectWebsite 
            AutoSize        =   -1  'True
            Caption         =   "http://www.computec.ch/projekte/atk/"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   0
            MouseIcon       =   "frmAbout.frx":3108
            MousePointer    =   99  'Custom
            TabIndex        =   13
            ToolTipText     =   "Visit the project web site"
            Top             =   600
            Width           =   3315
         End
         Begin VB.Label lblMoreInformation 
            Caption         =   "You will find more informations about the Attack Tool Kit (ATK) on the project web site"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Frame fraMerchandising 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Image imgShirt 
            Height          =   2115
            Left            =   720
            MouseIcon       =   "frmAbout.frx":3412
            MousePointer    =   99  'Custom
            Picture         =   "frmAbout.frx":371C
            ToolTipText     =   "Click here to buy an ATK shirt online!"
            Top             =   480
            Width           =   2025
         End
         Begin VB.Label lblShirtAdvertisement 
            Alignment       =   2  'Center
            Caption         =   "Go and get your ATK shirt online!"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmAbout.frx":4123
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.Frame fraProjectTeam 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   3495
         Begin MSComctlLib.ListView lsvTeam 
            Height          =   1815
            Left            =   0
            TabIndex        =   7
            Top             =   720
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlImages"
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
               Text            =   "Email"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Web"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Function"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblProjectTeamIntro 
            Caption         =   "The Attack Tool Kit (ATK) is developed and supported by the following people (alphabetical order):"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   3375
         End
      End
      Begin VB.Label lblSoftwareNameLong 
         Caption         =   "Attack Tool Kit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblSoftwareNameShort 
         Caption         =   "ATK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip tspAbout 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      TabWidthStyle   =   2
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            Object.ToolTipText     =   "Information about the project"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Project &Team"
            Object.ToolTipText     =   "Information about the ATK project team"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Donate"
            Object.ToolTipText     =   "Donate to support the ATK project"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Merchandising"
            Object.ToolTipText     =   "Attack Tool Kit merchandising (e.g. shirts)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************************************************************************************
' * Frame Description                                                                *
' *                                                                                  *
' * In this frame the about can be found: Information about the tool, project and    *
' * project team members.                                                            *
' ************************************************************************************

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.0 2004-10-13                                                           *
' * - Added the tab and frame for the PayPal donation.                               *
' * - Fixed the hidden painting copyright information.                               *
' * Version 2.0 2004-08-20                                                           *
' * - Added a TabStrip to divide the normal about and the data about the team.       *
' ************************************************************************************

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 27
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = application_name & " About"
    WriteLogEntry frmAbout.Caption & " opened.", 6

    'Adding the ATK project team members into the list view
    Call AddTeamMember("Abt, Björn", "flow@swissonline.ch", "http://www.inode.ch", "Concept Developement")
    Call AddTeamMember("Anner, Martin", "martin_anner@swissonline.ch", application_website_url, "ATK Alien Logo Painting")
    Call AddTeamMember("Brecht, Roland", "info@wireless-warrior.org", "http://www.wireless-warrior.org", "Beta Testing")
    Call AddTeamMember("Bytewolf", "bytewolf@gmail.com", application_website_url, "Beta Testing")
    Call AddTeamMember("Carvey, Harlan", "keydet89@yahoo.com", "http://www.windows-ir.com", "Beta Testing")
    Call AddTeamMember("Covello, Andrea", "andrea@covello.ch", "http://www.covello.ch", "Beta Testing")
    Call AddTeamMember("D., Eric", "info@xinulsystems.us", "http://www.xinulsystems.us", "Beta Testing")
    Call AddTeamMember("Gagliardi, Rocco", "rocco_gagliardi@gmx.net", "http://lupig.mine.nu", "Beta Testing")
    Call AddTeamMember("Häfner, Marcel", "haefner.marcel@heavy.ch", "http://www.lorky.heavy.ch", "Beta Testing")
    Call AddTeamMember("Keller, Fabian", "sitch@sitch.ch", "http://www.sitch.ch", "Beta Testing")
    Call AddTeamMember("Kyong Joo, Jung", "jyj9782@kornet.net", "http://www.chollian.net/~jyj9782", "Beta Testing")
    Call AddTeamMember("Längle, Erol", "erol_laengle@web.de", "http://www.getronics.ch", "Beta Testing")
    Call AddTeamMember("Lijian", "lijian1976@hotmail.com", application_website_url, "Bug fixing of modDNS in ATK 3.1")
    Call AddTeamMember("Moser, Max", "mmo@remote-exploit.org", "http://www.remote-exploit.org", "Beta Testing")
    Call AddTeamMember("Nester, David", "david@icrew.org", "http://www.icrew.org", "Plugin Developement")
    Call AddTeamMember("Pelkmann, Armin", "apelkmann@freenet.de", "http://sicherheit.freenet.de", "Beta Testing")
    Call AddTeamMember("Peschel, Gaby", "momolly@wireless-warrior.org", "http://www.wireless-warrior.org", "Beta Testing")
    Call AddTeamMember("Rogge, Marko", "mr@german-secure.de", "http://www.german-secure.de", "Beta Testing")
    Call AddTeamMember("Ruef, Marc", "marc.ruef@computec.ch", "http://www.computec.ch", "Chief Developer")
    Call AddTeamMember("Sengün, Gürkan", "gurkan@linuks.mine.nu", "http://www.linuks.mine.nu", "GPL Advisor")
    Call AddTeamMember("Spicher, Nico", "triplex@it-helpnet.de", "http://www.it-helpnet.de", "Beta Testing and Plugin Developement")
    Call AddTeamMember("Widmer, Pascal", "info@abteilung.ch", "http://www.abteilung.ch", "Icon Developement")
    Call AddTeamMember("Zumstein, Simon", "sizu@scip.ch", "http://www.scip.ch", "Beta Testing")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteLogEntry frmAbout.Caption & " closed.", 6
    Set frmAbout = Nothing
End Sub

Private Sub imgPayPalDonation_Click()
    Call lblProjectWebsite_Click
End Sub

Private Sub imgShirt_Click()
    Call lblShirtAdvertisement_Click
End Sub

Private Sub lblCopyrights_Click()
    ReadText (lblCopyrights.Caption & " " & lblName.Caption)
End Sub

Private Sub lblDescription_Click()
    ReadText (lblDescription.Caption)
End Sub

Private Sub lblEmail_Click()
    Dim DestinationEMail As String
    
    DestinationEMail = lblEmail.Caption
    
    WriteLogEntry "Writing an email to " & DestinationEMail, 6
    Call ShellExecute(hwnd, "Open", "mailto:" & DestinationEMail & "?subject=" & application_name, "", "", 1)
End Sub

Private Sub lblMoreInformation_Click()
    ReadText (lblMoreInformation.Caption)
End Sub

Private Sub lblProjectWebsite_Click()
    Dim WebSiteURL As String
    
    WebSiteURL = lblProjectWebsite.Caption
    
    'Load the project web site
    WriteLogEntry "Loading the project website " & WebSiteURL, 6
    Call ShellExecute(Me.hwnd, "Open", WebSiteURL, "", App.Path, 1)
End Sub

Private Sub lblShirtAdvertisement_Click()
    Dim strShirtShopURL As String

    strShirtShopURL = application_website_url & "shop/frames.html"

    'Load the shirt shop
    WriteLogEntry "Loading the shirt shop " & strShirtShopURL, 6
    Call ShellExecute(Me.hwnd, "Open", strShirtShopURL, "", App.Path, 1)
End Sub

Private Sub lblWeb_Click()
    Dim WebSiteURL As String
    
    WebSiteURL = lblWeb.Caption
    
    'Load my web site
    WriteLogEntry "Loading the website " & WebSiteURL, 6
    Call ShellExecute(Me.hwnd, "Open", WebSiteURL, "", App.Path, 1)
End Sub

Private Sub imgLogo_Click()
    Call lblProjectWebsite_Click
End Sub

Private Sub lsvTeam_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewColumnReorder(frmAbout.lsvTeam, ColumnHeader)
End Sub

Private Sub lsvTeam_DblClick()
    Dim WebSiteURL As String
    
    If lsvTeam.ListItems.Count <> 0 Then
        WebSiteURL = lsvTeam.SelectedItem.SubItems(2)
        
        'Load the project web site
        WriteLogEntry "Loading the ATK project team member website " & WebSiteURL, 6
        Call ShellExecute(Me.hwnd, "Open", WebSiteURL, "", App.Path, 1)
    End If
End Sub

Private Sub tspAbout_Click()
    Dim intSelectedItem As Integer
    
    intSelectedItem = tspAbout.SelectedItem.Index

    If intSelectedItem = 1 Then
        fraAboutInfo.Visible = True
        fraProjectTeam.Visible = False
        fraDonate.Visible = False
        fraMerchandising.Visible = False
    ElseIf intSelectedItem = 2 Then
        fraProjectTeam.Visible = True
        fraAboutInfo.Visible = False
        fraDonate.Visible = False
        fraMerchandising.Visible = False
    ElseIf intSelectedItem = 3 Then
        fraDonate.Visible = True
        fraProjectTeam.Visible = False
        fraAboutInfo.Visible = False
        fraMerchandising.Visible = False
    ElseIf intSelectedItem = 4 Then
        fraMerchandising.Visible = True
        fraAboutInfo.Visible = False
        fraDonate.Visible = False
        fraProjectTeam.Visible = False
    End If
End Sub

Private Sub tspAbout_KeyPress(KeyAscii As Integer)
    If KeyAscii = "27" Then
        Unload Me
    End If
End Sub

Private Sub AddTeamMember(ByRef strName As String, ByRef strEmail As String, _
    ByRef strWebsite As String, ByRef strFunction As String)

    Dim List As ListItem        'Needed for the listview handling

    'Write the log data into the log frame
    Set List = lsvTeam.ListItems.Add(, , strName)
        List.SubItems(1) = strEmail
        List.SubItems(2) = strWebsite
        List.SubItems(3) = strFunction
    
    'Set the right column width
    LVColumnWidth lsvTeam
End Sub
    


