VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPassPort 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Passport Settings"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   9615
   Icon            =   "frmPassPort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassPort.frx":0ECA
   ScaleHeight     =   5595
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar PB 
      Height          =   345
      Left            =   -30
      TabIndex        =   17
      Top             =   5280
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Max             =   80
      Scrolling       =   1
   End
   Begin VB.Timer tmrWEB 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8250
      Top             =   150
   End
   Begin VB.TextBox txtWEB 
      Height          =   645
      Left            =   2940
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   5040
   End
   Begin SHDocVwCtl.WebBrowser WEB 
      CausesValidation=   0   'False
      Height          =   1470
      Left            =   5100
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3150
      Width           =   4410
      ExtentX         =   7779
      ExtentY         =   2593
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdGetPassport 
      Caption         =   "Get Passport"
      Height          =   315
      Left            =   5100
      TabIndex        =   2
      ToolTipText     =   "STR|Click to start the retreival/refresh of the Passport"
      Top             =   4830
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1575
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   3060
      Width           =   4965
      Begin VB.TextBox txtPassport 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   0
         Tag             =   "Passport|Email"
         ToolTipText     =   "STR|Enter here your E-Mail address you use for Passport access"
         Top             =   540
         Width           =   4590
      End
      Begin VB.TextBox txtPassport 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Tag             =   "Passport|Password"
         ToolTipText     =   "STR|Enter here your E-Mails password"
         Top             =   1140
         Width           =   4590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   4590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Email Address to Use - Must be activated for passport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   4590
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   2505
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   570
      Width           =   9465
      Begin VB.TextBox txtPassport 
         Height          =   600
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Tag             =   "Passport|PassportProfile"
         ToolTipText     =   "STR|This is retreived for you\nThis is the profile link"
         Top             =   1740
         Width           =   9255
      End
      Begin VB.TextBox txtPassport 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Tag             =   "Passport|PassportTicket"
         ToolTipText     =   "STR|The PassportTicket is retreived for you\nThis expires every 12 hours or so\nSo a refresh is required to get a new one"
         Top             =   1230
         Width           =   9255
      End
      Begin VB.TextBox txtPassport 
         Height          =   600
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Tag             =   "Passport|Cookie"
         ToolTipText     =   "STR|The MSNREGCookie is retreived for you"
         Top             =   420
         Width           =   9255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "MSNREGCookie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "PassportTicket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "PassportProfile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1530
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   4830
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   4830
      Width           =   1365
   End
End
Attribute VB_Name = "frmPassPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bClose As Boolean
Private m_objTooltip    As cTooltip
' Const sRoom As String = "PassPortRetrievalRoom"
Const sRoom As String = "Teens"


Private Sub cmdCancel_Click()
      Screen.MousePointer = vbNormal
      tmrWEB.Enabled = False
      bClose = True
      Me.Hide
End Sub

Private Sub cmdGetPassport_Click()
Dim SUrl As String

      Me.txtPassport(0) = ""
      Me.txtPassport(1) = ""
      Me.txtPassport(2) = ""
      SUrl = "https://loginnet.passport.com/ppsecure/post.srf?&id=2260&ru=http://chat.msn.com/chatroom.msnw?rm=" & sRoom & "&login=" & Me.txtPassport(3) & "&passwd=" & Me.txtPassport(4) & ""
      PB.Value = 0

      tmrWEB.Enabled = True
      WEB.Navigate SUrl
      PB.Value = PB.Value + 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
End Sub
Private Function WaitBusy() As Boolean
      Do While WEB.Busy
         DoEvents
         If bClose Then
            WaitBusy = True
            GoTo Hell
         End If
      Loop
Hell:
End Function
Private Sub cmdOk_Click()
      Screen.MousePointer = vbNormal
      tmrWEB.Enabled = False
      bClose = True
      Save_Settings Me
      WEB.Navigate ""
      Me.Hide
End Sub
Private Sub Form_Load()
      WEB.Navigate "about: "
      PB.Value = 0
      LoadToolTips Me, m_objTooltip
      Call Load_Settings(Me)
End Sub

Private Sub tmrWEB_Timer()
      On Error Resume Next
      txtWEB = WEB.Document.documentelement.innerhtml
      DoEvents
End Sub

Private Sub txtWEB_Change()
Dim sTemp As String
Dim asData() As String
Dim i As Integer

      If PB.Value + 10 <= PB.Max Then
         PB.Value = PB.Value + 10
      End If
      tmrWEB.Enabled = False
      sTemp = txtWEB
      If Locate(sTemp, "MSNREGCookie") Then
         WEB.Navigate "about: "
         DoEvents
         sTemp = Replace(Replace(Mid$(sTemp, Locate(sTemp, "PARAM NAME=""MSNREGCookie"), Len(sTemp)), "temp += '<", ""), "  ", "")
         sTemp = Left$(sTemp, Locate(sTemp, "PARAM NAME=""ChatMode") - 3)
         asData = Split(sTemp, ";")

         For i = 0 To UBound(asData)
            sTemp = Split(asData(i), """")(3)
            txtPassport(i) = sTemp
         Next
         Screen.MousePointer = vbNormal
         WEB.Navigate "http://login.passport.com/logout.srf?&id=2260&ru=http://chat.msn.com/chatroom.msnw%3frm%3d" & sRoom & "&login=" & Me.txtPassport(3) & "&passwd=" & Me.txtPassport(4) & ""
         Exit Sub
      End If
      If Locate(WEB.LocationURL, "default.msnw") Then
         WEB.Navigate "about: "
      End If
      tmrWEB.Enabled = True

End Sub

Private Sub WEB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
      On Error Resume Next
      If WEB.LocationURL Like "*file:///C:/html.html*" Then
         WEB.SetFocus
         SendKeys "{enter}"
         PB.Value = PB.Value + 10
      End If
      If WEB.LocationURL Like "*Cookies*" Then
         WEB.Navigate "http://login.passport.com/login.srf"
      End If
      If WEB.LocationURL Like "http://login.passport.com/login.sr*" Then
         WEB.Navigate "https://login.passport.com/ppsecure/post.srf?id=2260&ru=http://chat.msn.com/chatroom.msnw%3frm%3d" & sRoom & "&login=" & Me.txtPassport(3) & "&passwd=" & Me.txtPassport(4) & ""
      End If
      If WEB.LocationURL Like "http://chat.msn.com/chatroom.msnw*" Then
         WEB.Navigate "http://chat.msn.com.my/chatroom_ui.msnw?rm=" & sRoom
         PB.Value = PB.Value + 10
      End If
'      txtWEB.Text = WEB.LocationURL

End Sub

