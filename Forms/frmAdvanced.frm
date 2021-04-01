VERSION 5.00
Begin VB.Form frmControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop at Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   60
      TabIndex        =   25
      ToolTipText     =   "STR|Stop auth once on second server"
      Top             =   8370
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add to Kick List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   2
      Left            =   1380
      Picture         =   "frmAdvanced.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "STR|Add selected guests to kick list"
      Top             =   3330
      Width           =   825
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add to Host List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   1
      Left            =   1380
      Picture         =   "frmAdvanced.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "STR|Add Selected guests to host list"
      Top             =   2400
      Width           =   825
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add to Owner List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   0
      Left            =   1380
      Picture         =   "frmAdvanced.frx":13D4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "STR|Add Selected guests to Owners List"
      Top             =   1470
      Width           =   825
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Silent Ban"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":1DBE
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "STR|Add guest to access list (BAN) \n (NO KICK)"
      Top             =   7500
      Width           =   1275
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Custom Kick/Ban"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":27A8
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "STR|Custom kick - Select own message\(CAN BAN)"
      Top             =   6780
      Width           =   1275
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kick (Advertising)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":3072
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "STR|Kick for advertising\n (NO BAN)"
      Top             =   6060
      Width           =   1275
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kick (Scrolling)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":393C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "STR|Kick for scrolling\n (NO BAN)"
      Top             =   5340
      Width           =   1275
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kick (Profanity)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":4206
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "STR|Kick for profanity\n (NO BAN)"
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton cmdKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kick (Disruptive)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":4AD0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "STR|Kick for disruptive behaviour\n (NO BAN)"
      Top             =   3900
      Width           =   1275
   End
   Begin VB.CommandButton cmdProfile 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   660
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":539A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "STR|Check profile"
      Top             =   3210
      Width           =   645
   End
   Begin VB.CommandButton cmdIdent 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ident"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":5D84
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "STR|Request user identity\nif user is using IRCDominator"
      Top             =   3210
      Width           =   645
   End
   Begin VB.CommandButton cmdTime 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   660
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":676E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "STR|Request guests local time"
      Top             =   2640
      Width           =   645
   End
   Begin VB.CommandButton cmdParticipant 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Participant"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   660
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":6AF8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "STR|Make the guest(s) a Paticipant"
      Top             =   2070
      Width           =   645
   End
   Begin VB.CommandButton cmdWhisper 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Whisper"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":6E82
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "STR|Initiate a whisper session"
      Top             =   2640
      Width           =   645
   End
   Begin VB.CommandButton cmdSpec 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Spectate"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":740C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "STR|Take Voice / Spectate user"
      Top             =   2070
      Width           =   645
   End
   Begin VB.CommandButton cmdBrown 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Host"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   660
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":7996
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "STR|Make the guest(s) a  Host"
      Top             =   1500
      Width           =   645
   End
   Begin VB.CommandButton cmdOwner 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Owner"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   30
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmAdvanced.frx":7F20
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "STR|Make the guest(s) a Owner"
      Top             =   1500
      Width           =   645
   End
   Begin VB.TextBox txtJoinKick 
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Text            =   "GoodBye"
      ToolTipText     =   "STR|Custom Message for auto kick"
      Top             =   1200
      Width           =   2385
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2190
      Top             =   30
   End
   Begin VB.CheckBox chkFlashNick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Flash Nick Name"
      Height          =   945
      Left            =   1740
      TabIndex        =   22
      Tag             =   "Only Works for Guest NickNames"
      Top             =   7680
      Width           =   825
   End
   Begin VB.CheckBox chkAutoKick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Auto Kick Joins"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Tag             =   "Advanced|AutoKick"
      ToolTipText     =   "STR|Kick any guest that joins"
      Top             =   990
      Width           =   1635
   End
   Begin VB.CheckBox chkAutoPart 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Auto Participant"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Tag             =   "Advanced|AutoP"
      ToolTipText     =   "STR|Make guest joined a paticipant\n(used when room is moderated)"
      Top             =   270
      Width           =   1635
   End
   Begin VB.CheckBox chkAutoOwner 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Auto Owner Joins"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Tag             =   "Advanced|AutoOwner"
      ToolTipText     =   "STR|Make guest joined an Owner"
      Top             =   510
      Width           =   1635
   End
   Begin VB.CheckBox chkAutoHost 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Auto Host Joins"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Tag             =   "Advanced|AutoHost"
      ToolTipText     =   "STR|Make guest joined a Host"
      Top             =   750
      Width           =   1635
   End
   Begin VB.CheckBox chkAutoV 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Auto Voice Joins"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Tag             =   "Advanced|AutoV"
      ToolTipText     =   "STR|Make guest joined have voice"
      Top             =   30
      Width           =   1635
   End
   Begin VB.CheckBox chkClone 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   270
      TabIndex        =   24
      ToolTipText     =   "STR|Use Clone use option\nWhen used with passport option\nWill allow the NickName to be any REAL NickName"
      Top             =   4530
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objTooltip    As cTooltip

Private Sub chkFlashNick_Click()
      If bJoined = True Then
         If chkFlashNick.Value = 0 Then
            Timer1.Enabled = False
            SendServer2 "NICK " & sNickJoined, True
            DoEvents
         Else
            Timer1.Enabled = True
         End If
      Else
         Timer1.Enabled = False
      End If
End Sub

Public Sub cmdAdd_Click(Index As Integer)
Dim sNick As String
Dim sList As String
Dim asNicks() As String
Dim i As Integer
Dim sTemp As String
Dim iLoop As Integer

      sList = BuildNameList
      asNicks = Split(sList, ",")
      For i = 0 To UBound(asNicks)
         sNick = TestNick(asNicks(i), True)
         Select Case Index
            Case 0
               ' Add Owner
               If HostLists.IsInList(sNick, lOwnerList) = False Then
                  HostLists.AddToList sNick, lOwnerList
                  HostLists.SavePrefs
               End If
            Case 1
               ' Add Host
               If HostLists.IsInList(sNick, lHostList) = False Then
                  HostLists.AddToList sNick, lHostList
                  HostLists.SavePrefs
               End If
            Case 2
               ' Add KickList
               If KickSettings.IsInList(sNick, lKickList) = False Then
                  ' Add to list
                  KickSettings.AddToList sNick, lKickList
                  KickSettings.SavePrefs
                  Call DoKick(sNick, KickSettings.KickList_Message, False)
                  iLoop = iLoop + 1
                  If iLoop > 5 Then
                     For iLoop = 1 To 300000
                        DoEvents
                     Next
                     iLoop = 0
                  End If
               End If
         End Select
      Next
End Sub

Public Sub cmdBrown_Click()
      DoUserOp "+o"
End Sub
Private Sub cmdDisable_Click()
      TellDisable
End Sub

Private Sub cmdGetGold_Click()
      AskSecretGold
End Sub

Private Sub cmdGetSecret_Click()
      AskSecret
End Sub
Public Sub cmdKick_Click(Index As Integer)
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim iLoop As Long
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
        
      ' Build List up first as user list can change
      sList = BuildNameList
      asList = Split(sList, ",")
      If Index = 4 Then
         ' Ask for custoM kick / ban
         Load frmKickBan
         frmKickBan.bSilent = False
         frmKickBan.Show
         Exit Sub
      End If
      If Index = 5 Then
         ' Ask for SILENT  ban
         Load frmKickBan
         frmKickBan.bSilent = True
         frmKickBan.Show
         Exit Sub
      End If
      For i = 0 To UBound(asList)
         sNick = asList(i)
         If Index < 4 Then
            Call DoKick(sNick, LoadResString(148 + Index), False)
         End If
         iLoop = iLoop + 1
         If iLoop > 5 Then
            For iLoop = 1 To 300000
               DoEvents
            Next
            iLoop = 0
         End If
      Next

Hell:
End Sub

Public Sub cmdOwner_Click()
      DoUserOp "+q"
End Sub

Public Sub cmdParticipant_Click()
      DoUserOp "-o"
End Sub

Public Sub cmdIdent_Click()
      AskIdent
End Sub

Public Sub cmdProfile_Click()
Dim sTemp As String
Dim sProf As String
Dim sNick As String

      On Error GoTo Hell

      sProf = "http://members.msn.com/Default.msnw?mpp=2208~%p&mid=2200"
      sNick = frmMain.tUsers.ListItems.Item(frmMain.tUsers.SelectedItem.Index).SubItems(2)

      sTemp = SendGetResponse("PROP " & sNick & " PUID", " :End of properties")
      sTemp = Split(sTemp, "PUID :")(1)
      sTemp = Split(sTemp, vbCrLf)(0)
      
      sProf = Replace(sProf, "%p", sTemp)
            
      MDIMain.wb.Navigate "about: blank"
      DoEvents
      MDIMain.wb.Navigate ("javascript:window.open('" & sProf & "')")
Hell:
End Sub

Private Sub cmdRelock_Click()
      TellRelock
End Sub

Public Sub cmdSpec_Click()
      DoUserOp "-v"
End Sub

Public Sub cmdTime_Click()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim sList As String
Dim asNicks() As String

      sList = BuildNameList
      asNicks = Split(sList, ",")

      On Error GoTo Hell
      If frmMain.tMe.ListItems(1).Selected = True Then
         Exit Sub
      End If
      For i = 0 To UBound(asNicks)
         sNick = TestNick(asNicks(i), True)
         Call PRIVMSG("TIME", sNick, False, False, pChat)
      Next
Hell:
        
End Sub

Public Sub cmdWhisper_Click()
      StartWhisper
End Sub

Private Sub Command1_Click()
      SendServer2 "Names"
End Sub

Private Sub Form_Activate()
      CheckActivated
End Sub
Public Sub CheckActivated()
      ' If bActivated And Me.cmdNuke.Value = False Then
      If bActivated Then
         ' MDIMain.mnuNuke.Visible = True
         ' cmdNuke.Visible = True
         ' cmdNukeRoom.Visible = True
         frmMain.chkPassport.Visible = True
         ' If bOwner = True Then
      Else
      End If
End Sub

Private Sub Form_Load()
      CheckActivated
      Load_Settings Me
      LoadToolTips Me, m_objTooltip

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         Cancel = True
         Me.WindowState = vbMinimized
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
      Save_Settings Me
End Sub

Private Sub Timer1_Timer()
      If Timer1.Tag = "" Then
         SendServer2 "NICK >", True
         DoEvents
         Timer1.Tag = sNickJoined
      Else
         SendServer2 "NICK " & sNickJoined, True
         DoEvents
         Timer1.Tag = ""
      End If
End Sub
