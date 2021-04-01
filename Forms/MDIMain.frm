VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "IRC Dominator - MSN Chat Room Manager"
   ClientHeight    =   7800
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   12420
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   7575
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   12360
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   12420
      Begin VB.Timer tmrTimer 
         Interval        =   1000
         Left            =   5280
         Top             =   510
      End
      Begin VB.Timer tmrActivity 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   5310
         Top             =   1440
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   1530
         ScaleHeight     =   315
         ScaleWidth      =   1905
         TabIndex        =   2
         Top             =   900
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Timer tmrJoin 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5310
         Top             =   990
      End
      Begin MSWinsockLib.Winsock Svr2 
         Left            =   1620
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Svr1 
         Left            =   1140
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ilstServer 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":0E42
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":115C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1476
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIMain.frx":1790
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Colours 
         Left            =   0
         Top             =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
      End
      Begin SHDocVwCtl.WebBrowser wb 
         Height          =   405
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   645
         ExtentX         =   1138
         ExtentY         =   714
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
         Location        =   "http:///"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAutoJoin 
         Caption         =   "Auto Join at Startup"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSep13 
      Caption         =   "|"
   End
   Begin VB.Menu mnuPrefopts 
      Caption         =   "Preferences"
      Begin VB.Menu mnuPrefs 
         Caption         =   "General Preferences"
      End
      Begin VB.Menu mnuPassPort 
         Caption         =   "PassPort Preferences"
      End
      Begin VB.Menu mnuWelcomePrefs 
         Caption         =   "Welcome Preferences"
      End
      Begin VB.Menu mnuAutoPrefs 
         Caption         =   "Auto Kick Prefences"
      End
      Begin VB.Menu mnuAutoHost 
         Caption         =   "Auto Host / Owner Preferences"
      End
      Begin VB.Menu mnuTrace 
         Caption         =   "Trace Preferences"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdvertising 
         Caption         =   "Advertising Preferences"
      End
      Begin VB.Menu mnuDoAdverts 
         Caption         =   "Enable Avertising"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuSep11 
      Caption         =   "|"
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "Connect"
   End
   Begin VB.Menu mnuAbort 
      Caption         =   "Abort Connection"
   End
   Begin VB.Menu mnuSep12 
      Caption         =   "|"
   End
   Begin VB.Menu mnuRejoin 
      Caption         =   "Part/Join Room"
   End
   Begin VB.Menu mnuPart 
      Caption         =   "Part Room"
   End
   Begin VB.Menu mnuJoin 
      Caption         =   "Join Room"
   End
   Begin VB.Menu mnuPass 
      Caption         =   "Enter Room Pass"
   End
   Begin VB.Menu mnuRoomOptions 
      Caption         =   "Room Properties"
      Begin VB.Menu mnuRoomModes 
         Caption         =   "Room Modes"
         Begin VB.Menu mnuRoomLimit 
            Caption         =   "Set Room Limit"
         End
         Begin VB.Menu mnuRoomModerated 
            Caption         =   "Moderated Room"
         End
         Begin VB.Menu mnuRoomInvite 
            Caption         =   "Room is Invite Only"
         End
         Begin VB.Menu mnuRoomSecret 
            Caption         =   "Room is Secret"
         End
         Begin VB.Menu mnuRoomPrivate 
            Caption         =   "Room is Private"
         End
         Begin VB.Menu mnuRoomGuestWhispers 
            Caption         =   "No Guest Whispers"
         End
         Begin VB.Menu mnuRoomNoWhispers 
            Caption         =   "No Whispers at all (hosts only)"
         End
         Begin VB.Menu mnuOver100 
            Caption         =   "Allow over 100 in room"
         End
      End
      Begin VB.Menu mnuTopic 
         Caption         =   "Change Topic"
      End
      Begin VB.Menu mnuOnJoin 
         Caption         =   "OnJoin Message"
      End
      Begin VB.Menu mnuOwnerKey 
         Caption         =   "Change Owner Key"
      End
      Begin VB.Menu mnuHostKey 
         Caption         =   "Change Host Key"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewAccess 
         Caption         =   "View Access List"
      End
      Begin VB.Menu mnuClearAccess 
         Caption         =   "Clear Access List (bans)"
      End
      Begin VB.Menu mnuGuestBan 
         Caption         =   "Guests Bans"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuCheckUpdate 
         Caption         =   "Check For Update"
      End
      Begin VB.Menu mnuMOTD 
         Caption         =   "Message of the Day"
      End
   End
   Begin VB.Menu mnuPopups 
      Caption         =   "Popups"
      Begin VB.Menu mnuPopup 
         Caption         =   "PopupMenu"
         Begin VB.Menu mnuProfile 
            Caption         =   "&View Profile"
         End
         Begin VB.Menu mnuWhisper 
            Caption         =   "&Whisper"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuIgnore 
            Caption         =   "Ignore User"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuTag 
            Caption         =   "&Tag User"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTime 
            Caption         =   "Local &Time"
         End
         Begin VB.Menu mnuIdent 
            Caption         =   "&Ident User"
         End
         Begin VB.Menu mnuSep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuKick 
            Caption         =   "&Kick"
            Begin VB.Menu mnuKickDis 
               Caption         =   "&Disruptive Behaviour"
            End
            Begin VB.Menu mnuKickProfanity 
               Caption         =   "&Profanity"
            End
            Begin VB.Menu mnuKickScrolling 
               Caption         =   "&Scrolling"
            End
            Begin VB.Menu mnuCustomKick 
               Caption         =   "&Custom Kick/Ban..."
            End
         End
         Begin VB.Menu mnuSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOwner 
            Caption         =   "Make &Owner"
         End
         Begin VB.Menu mnuHost 
            Caption         =   "Make &Host"
         End
         Begin VB.Menu mnuParticipant 
            Caption         =   "Make Participant"
         End
         Begin VB.Menu mnuSpectator 
            Caption         =   "Make &Spectator"
         End
         Begin VB.Menu mnuSep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuListAdd 
            Caption         =   "Add to Owner List"
            Index           =   0
         End
         Begin VB.Menu mnuListAdd 
            Caption         =   "Add to Host List"
            Index           =   1
         End
         Begin VB.Menu mnuListAdd 
            Caption         =   "Add to Kick List"
            Index           =   2
         End
      End
      Begin VB.Menu mnuChatPopup 
         Caption         =   "ChatPopup"
         Begin VB.Menu mnuChatCopy 
            Caption         =   "&Copy"
         End
         Begin VB.Menu mnuChatSelectAll 
            Caption         =   "&Select All"
         End
         Begin VB.Menu mnuChatClear 
            Caption         =   "C&lear All"
         End
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Public Sub mnuAdvertising_Click()
      frmAdvertising.Show 0, Me
End Sub

Public Sub mnuAutoHost_Click()
      frmHostLists.Show 1, Me
      Unload frmHostLists
      Set frmHostLists = Nothing
End Sub

Private Sub mnuAutoJoin_Click()

      If Me.mnuAutoJoin.Checked = True Then
         Me.mnuAutoJoin.Checked = False
         Call PutIni("General", "Auto Join", False)
      Else
         Me.mnuAutoJoin.Checked = True
         Call PutIni("General", "Auto Join", True)
      End If
End Sub

Private Sub mnuAutoPrefs_Click()
      frmAutoPrefs.Show 1, Me
      Unload frmAutoPrefs
      Set frmAutoPrefs = Nothing
End Sub

Private Sub mnuCheckUpdate_Click()
      If MsgBox("Launch AutoUpdate", vbQuestion + vbYesNo, "IRCDominator AutoUpdate") = vbYes Then
         Shell App.Path & "\AutoUpdater.exe", vbNormalFocus
         AppActivate "Live Update - Step 1"
      End If
End Sub

Private Sub mnuDoAdverts_Click()
Dim sTemp As String

      If Me.mnuDoAdverts.Checked = True Then
         Me.mnuDoAdverts.Checked = False
      Else
         Me.mnuDoAdverts.Checked = True
         frmAdvertising.lblInterval(0).Tag = 0
         frmAdvertising.lblInterval(1).Tag = 0
         frmAdvertising.lblInterval(2).Tag = 0
      End If
End Sub

Private Sub mnuListAdd_Click(Index As Integer)
      frmControl.cmdAdd_Click (Index)
End Sub

Private Sub mnuMOTD_Click()
      frmMOTD.Show 1
End Sub

Private Sub mnuOver100_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuOver100.Checked = True Then
         ' Turn off
         sTemp = sTemp & " +l 100"
         Me.mnuOver100.Checked = False
         frmControl.chkAutoPart.Value = 0
         SendServer2 sTemp
         DoEvents
         sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " DELETE HOST *!*@*"
         SendServer2 sTemp
         DoEvents
      Else
         ' Turn on
         sTemp = sTemp & " +l 100"
         Me.mnuOver100.Checked = True
         frmControl.chkAutoPart.Value = 1
         SendServer2 sTemp
         DoEvents
         sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " ADD HOST *!*@*"
         SendServer2 sTemp
         DoEvents
      End If
End Sub

' Private Sub mnuNukePrefs_Click()
' frmNukeKick.Show 0, Me
' End Sub

Private Sub mnuRoomGuestWhispers_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuRoomGuestWhispers.Checked = True Then
         ' Turn off
         sTemp = sTemp & " -W"
         Me.mnuRoomGuestWhispers.Checked = False
      Else
         ' Turn on
         sTemp = sTemp & " +W"
         Me.mnuRoomGuestWhispers.Checked = True
      End If
      SendServer2 sTemp
End Sub

Private Sub mnuRoomNoWhispers_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuRoomNoWhispers.Checked = True Then
         ' Turn off
         sTemp = sTemp & " -w"
         Me.mnuRoomNoWhispers.Checked = False
      Else
         ' Turn on
         sTemp = sTemp & " +w"
         Me.mnuRoomNoWhispers.Checked = True
      End If
      SendServer2 sTemp
End Sub

Private Sub mnuRoomOptions_click()
Dim iMyState As Integer

      On Error Resume Next
      iMyState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))
      Me.mnuRoomInvite.Enabled = False
      Me.mnuRoomLimit.Enabled = False
      Me.mnuRoomModerated.Enabled = False
      Me.mnuRoomPrivate.Enabled = False
      Me.mnuRoomSecret.Enabled = False
      Me.mnuOver100.Enabled = False
      Me.mnuGuestBan.Enabled = False
      Me.mnuClearAccess.Enabled = False
      Me.mnuOwnerKey.Enabled = False
      Me.mnuHostKey.Enabled = False
      Me.mnuTopic.Enabled = False
      Me.mnuViewAccess.Enabled = False
      Me.mnuOnJoin.Enabled = False
      Me.mnuRoomGuestWhispers.Enabled = False
      Me.mnuRoomNoWhispers.Enabled = False

      If iMyState = 1 Or iMyState = 2 Then
         Me.mnuRoomInvite.Enabled = True
         Me.mnuRoomLimit.Enabled = True
         Me.mnuRoomModerated.Enabled = True
         Me.mnuRoomPrivate.Enabled = True
         Me.mnuRoomSecret.Enabled = True
         Me.mnuOver100.Enabled = True
         Me.mnuGuestBan.Enabled = True
         Me.mnuClearAccess.Enabled = True
         Me.mnuOwnerKey.Enabled = True
         Me.mnuHostKey.Enabled = True
         Me.mnuTopic.Enabled = True
         Me.mnuViewAccess.Enabled = True
         Me.mnuOnJoin.Enabled = True
         Me.mnuRoomGuestWhispers.Enabled = True
         Me.mnuRoomNoWhispers.Enabled = True
      End If
End Sub
Private Sub MDIForm_Load()
Dim i As Integer

      If bLoadingApp Then
         frmSplash.pgProgress.Max = (16 * 5) + (Screen.FontCount * 5)
      End If
      frmMain.chkPassport.Visible = False
      MDIMain.mnuTrace.Visible = False
      bShowSysTray = True
      SetSystray
      ' If IsUnlocked("NP132262A") Then
      UnlockMe
      m_frmSysTray.mnuSystray1(4).Enabled = True
      ' Else
      ' MDIMain.mnuPassPort.Enabled = False
      ' End If
      If IsUnlocked("GeniusIsGOD") Then
         MDIMain.mnuPassPort.Enabled = True
         MDIMain.mnuTrace.Visible = True
         m_frmSysTray.mnuSystray1(4).Enabled = True
         m_frmSysTray.mnuSystray1(8).Visible = True
         frmMain.chkPassport.Visible = True
         frmControl.chkClone.Visible = True
         frmControl.chkClone.Left = 1380
         bActivated = True
      End If
      ' If IsUnlocked("SpecialAccess") Then
      ' frmMain.chkPassport.Visible = True
      ' m_frmSysTray.mnuSystray1(4).Enabled = True
      ' m_frmSysTray.mnuSystray1(8).Visible = True
      ' bSpecialActivated = True
      ' mnuUnlock.Visible = False
      ' UnlockSpecial
      ' End If
      ' If IsUnlocked("SpecialSpecial") Then
      ' frmControl.chkClone.Visible = True
      ' frmControl.cmdGetGold.Visible = True
      ' frmMain.chkPassport.Visible = True
      ' m_frmSysTray.mnuSystray1(4).Enabled = True
      ' m_frmSysTray.mnuSystray1(8).Visible = True
      ' bSpecialSpecialActivated = True
      ' mnuUnlock.Visible = False
      ' UnlockSpecial
      ' End If

      For i = 1 To 16
         picCol.BackColor = QBColor(FindColour(i))
         Colours.ListImages.Add i, , picCol.Image
      Next
      Me.mnuPopups.Visible = False
      Call SetupFormArrays
      IniFile.Section = "General"
      IniFile.Key = "AppVersion"
      IniFile.Default = App.Major & "." & App.Minor & "." & App.Revision
      IniFile.Value = App.Major & "." & App.Minor & "." & App.Revision
      
      bAdvanced = True
      Load frmMain
      DoEvents
      If GeneralSettings.ShowTrace Then
         Load frmTrace
      End If
      Load frmControl
      Load frmWelcomePrefs
      Load frmAdvertising
      Me.Width = (Screen.Width / 1.3) + 500
      Me.Height = Screen.Height / 1.3
      ' 124   Me.WindowState = vbMaximized
      StatusColour = eCols.CGray
      Me.Top = (Screen.Height / 2) - Me.Height / 2
      Me.Left = (Screen.Width / 2) - Me.Width / 2
      frmMain.Show
      If GeneralSettings.ShowTrace Then
         frmTrace.Show
      End If
      frmControl.Show
      ResetSelected
      Me.Caption = Me.Caption & " - Version " & App.Major & "." & App.Minor & "." & App.Revision
      Me.mnuAbort.Enabled = False
      Me.mnuRoomOptions.Enabled = False
      Me.mnuJoin.Enabled = False
      Me.mnuPart.Enabled = False
      Me.mnuRejoin.Enabled = False
      Me.mnuPass.Enabled = False
      ResizeMe
      Me.mnuAutoJoin.Checked = fGetIni("General", "Auto Join", False)
      MDIMain.wb.Navigate "about: blank"
      If Dir(App.Path & "\Update.exe") <> "" Then
         Kill (App.Path & "\Update.exe")
      End If
      If Me.mnuAutoJoin.Checked = True Then
         DoConnect
      End If

End Sub
Private Sub MDIForm_Resize()
      If bShowSysTray And Me.WindowState = vbMinimized Then
         Me.Visible = False
      End If
      ResizeMe
End Sub
Public Sub ResizeMe()
      On Error GoTo Hell

      ' If frmPrefs.chkShowTrace.Value Then
      If GeneralSettings.ShowTrace Then
         frmMain.Move 0, 0, Me.Width - frmControl.Width - 200, Me.ScaleHeight - 2500
         frmTrace.Move 0, frmMain.Height, frmMain.Width, (Me.ScaleHeight - frmMain.ScaleHeight) - 500
      Else
         frmMain.Move 0, 0, Me.Width - frmControl.Width - 200, Me.ScaleHeight
      End If
      ' frmControl.Width = 2250
      frmControl.Top = 0
      frmControl.Left = frmMain.Width
      frmControl.Height = Me.ScaleHeight
Hell:
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
      bShutDown = True
      bStarted = False
      DoEvents
      Svr1.Close
      Svr2.Close
      Me.tmrActivity.Enabled = False
      Me.tmrJoin.Enabled = False
      frmChat.wb.Navigate ("About: Blank")
      Unload frmPopup
      Set frmPopup = Nothing
      Unload frmChat
      Set frmChat = Nothing
      Unload frmPrefs
      Set frmPrefs = Nothing
      CloseAllWindows (frmWhisperForm)
      Unload frmPassPort
      Set frmPassPort = Nothing
      Unload frmAutoPrefs
      Set frmAutoPrefs = Nothing
      Unload frmWelcomePrefs
      Set frmWelcomePrefs = Nothing
      Unload frmAdvertising
      Set frmAdvertising = Nothing
      ' Unload frmNukeKick
      ' Set frmNukeKick = Nothing
      Unload frmTraceOptions
      Set frmTraceOptions = Nothing
      Unload frmCheckLatest
      Set frmCheckLatest = Nothing
      On Error Resume Next
      Unload m_frmSysTray
      Set m_frmSysTray = Nothing
End Sub
Private Sub mnuAbort_Click()
      Svr1.Close
      Svr2.Close
      frmChat.wb.Navigate ("")
      Unload frmChat
      Set frmChat = Nothing
      bStarted = False
      DoEvents
      RefreshConnection True
      trace "Auth - Aborted"
      Me.mnuAbort.Enabled = False
      Me.mnuRoomOptions.Enabled = False
      Me.mnuJoin.Enabled = False
      Me.mnuPart.Enabled = False
      Me.mnuRejoin.Enabled = False
      Me.mnuPass.Enabled = False
End Sub
Private Sub mnuAbout_Click()
      frmAbout.Show 1
End Sub
Private Sub mnuChatSelectAll_Click()
      frmMain.txtChat.SelStart = 0
      frmMain.txtChat.SelStart = 0
      frmMain.txtChat.SelLength = Len(frmMain.txtChat.Text)
End Sub
Private Sub mnuChatClear_Click()
      frmMain.txtChat.Text = ""
End Sub
Private Sub mnuChatCopy_Click()
      Clipboard.SetText (frmMain.txtChat.SelText)
End Sub
Private Sub mnuClearAccess_Click()
Dim sTemp As String
      If MsgBox("Are You Sure", vbYesNo, "Clear Access List") = vbYes Then
         sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " CLEAR"
         SendServer2 sTemp
      End If
End Sub
Private Sub mnuConnect_Click()
      If GeneralSettings.ShowTrace Then
         ' If frmPrefs.chkShowTrace.Value Then
         frmTrace.txtTrace.Text = ""
      End If
      RefreshConnection True
      bStarted = False
      DoEvents
      DoEvents
      Me.mnuJoin.Enabled = False
      Me.mnuPart.Enabled = False
      Me.mnuRejoin.Enabled = False
      Me.mnuPass.Enabled = False
      Me.mnuDoAdverts.Checked = False
      On Error Resume Next
      ' Set MSNChatOCX = Nothing
      ' Set MSNChatOCX = New MSNChatProtocolCtl
    
      frmWelcomePrefs.chkWelcome.Value = 0
      frmWelcomePrefs.chkWelcomeAway.Value = 0
      If frmMain.chkPassport Then
         Load frmPassPort
      End If
      If GeneralSettings.Chat_ChatX = True Then
         trace "-= ATTEMPTING TO REGISTER MSNChatX.OCX =-"
         If GeneralSettings.ShowOCXs = True Then
            ' We are going to use the New MSNChatX ocx Control
            If Dir(App.Path & "\msnchatTest.ocx") <> "" Then
               Shell "regsvr32.exe " & Chr(34) & App.Path & "\msnchatTest.ocx" & Chr(34)
            Else
               trace "-= The MSNChatX.OCX File is not in the IRCDominator Application Directory =-"
            End If
         Else
            If Dir(App.Path & "\msnchatTest.ocx") <> "" Then
               Shell "regsvr32.exe /s " & Chr(34) & App.Path & "\msnchatTest.ocx" & Chr(34)
            Else
               trace "-= The MSNChatX.OCX File is not in the IRCDominator Application Directory =-"
            End If
         End If
      Else
         trace "-= ATTEMPTING TO REGISTER MSNChatTest.OCX =-"
         If GeneralSettings.ShowOCXs = True Then
            If Dir(App.Path & "\msnchatTest.ocx") <> "" Then
               Shell "regsvr32.exe  " & Chr(34) & App.Path & "\msnchatTest.ocx" & Chr(34)
            Else
               trace "-= The MSNChatTest.OCX File is not in the IRCDominator Application Directory =-"
            End If
         Else
            If Dir(App.Path & "\msnchatTest.ocx") <> "" Then
               Shell "regsvr32.exe  /s " & Chr(34) & App.Path & "\msnchatTest.ocx" & Chr(34)
            Else
               trace "-= The MSNChatTest.OCX File is not in the IRCDominator Application Directory =-"
            End If
         End If
      End If
      
      DoConnect True
End Sub
Private Sub mnuCustomKick_Click()
      frmControl.cmdKick_Click (4)
End Sub
Private Sub mnuExit_Click()
      Unload MDIMain
End Sub

Private Sub mnuGuestBan_Click()
      frmGuestBan.Show vbModal, Me
      Unload frmGuestBan
      Set frmGuestBan = Nothing
End Sub

Private Sub mnuHostKey_Click()
Dim sTemp As String

      On Error GoTo Hell
      sTemp = InputBox("Please enter the room HOST(Brown) Pass Code", "Set Room HOST/Brown Pass", "")
      If sTemp <> "" Then
         sTemp = "PROP %#" & FixSpaces(sRoomJoined, True) & " HOSTKEY " & sTemp
         SendServer2 sTemp
      End If
Hell:
End Sub
Public Sub mnuJoin_Click()
      DoRoomMode
      DoEvents
      bJoined = False
      bWaiting = False
      Me.mnuJoin.Enabled = False
      Me.mnuPart.Enabled = True
      SendServer2 "JOIN %#" & sRoomJoined & " " & frmMain.lblVersion4Pass
      DoEvents
      DoEvents
End Sub
Private Sub mnuKickDis_Click()
      frmControl.cmdKick_Click (0)
End Sub
Private Sub mnuKickProfanity_Click()
      frmControl.cmdKick_Click (1)
End Sub
Private Sub mnuKickScrolling_Click()
      frmControl.cmdKick_Click (2)
End Sub
Private Sub mnuOnJoin_Click()
Dim sTemp As String

      On Error GoTo Hell
      sTemp = InputBox("Please enter the rooms OnJoin(Welcome) message you require", "Set Room OnJoin/WelcomeMessage", "")
      If sTemp <> "" Then
         sTemp = "PROP %#" & FixSpaces(sRoomJoined, True) & " OnJoin :" & sTemp
         SendServer2 sTemp
      End If
Hell:

End Sub
Private Sub mnuOwnerKey_Click()
Dim sTemp As String

      On Error GoTo Hell
      sTemp = InputBox("Please enter the room OWNER(Gold) Pass Code", "Set Room OWNER/Gold Pass", "")
      If sTemp <> "" Then
         sTemp = "PROP %#" & FixSpaces(sRoomJoined, True) & " OWNERKEY " & sTemp
         SendServer2 sTemp
      End If
Hell:

End Sub
Public Sub mnuPart_Click()
'      RegDeleteValue HKEY_LOCAL_MACHINE, sRoot, sSubKeyV4, sKey1
      RefreshConnection True
      SendServer2 "PART %#" & sRoomJoined
      Me.mnuPart.Enabled = False
      Me.mnuJoin.Enabled = True
      CloseAllWindows (frmWhisperForm)
      ResetSelected
End Sub
Private Sub mnuPass_Click()
      DoPass
End Sub
Public Sub mnuPassPort_Click()
      frmPassPort.Show vbModal, Me
      Unload frmPassPort
      Set frmPassPort = Nothing
End Sub
Public Sub mnuPrefs_Click()
      frmPrefs.Show 1, Me
      Unload frmPrefs
      Set frmPrefs = Nothing
      If GeneralSettings.ShowTrace Then
         frmTrace.Show
      Else
         Unload frmTrace
      End If
      MDIMain.ResizeMe
End Sub
Private Sub mnuRejoin_Click()
      Me.mnuPart_Click
      DoEvents
      Me.mnuJoin_Click
End Sub
Private Sub mnuReset_Click()
      ResizeMe
End Sub
Private Sub mnuRoomInvite_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuRoomInvite.Checked = True Then
         ' Turn off
         sTemp = sTemp & " -i"
         Me.mnuRoomInvite.Checked = False
      Else
         ' Turn on
         sTemp = sTemp & " +i"
         Me.mnuRoomInvite.Checked = True
      End If
      SendServer2 sTemp
End Sub
Private Sub mnuRoomLimit_Click()
Dim iLimit As Integer
Dim sTemp As String

      On Error GoTo Hell
      iLimit = InputBox("Please enter the room limit you require 1-100", "Set Room Limit", 100)
      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True) & " +l " & iLimit
      SendServer2 sTemp
Hell:
End Sub
Private Sub mnuRoomModerated_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuRoomModerated.Checked = True Then
         ' Turn off
         sTemp = sTemp & " -m"
         Me.mnuRoomModerated.Checked = False
      Else
         ' Turn on
         sTemp = sTemp & " +m"
         Me.mnuRoomModerated.Checked = True
      End If
      SendServer2 sTemp
End Sub

Private Sub mnuRoomPrivate_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuRoomPrivate.Checked = True Then
         ' Turn off
         sTemp = sTemp & " -p"
         Me.mnuRoomPrivate.Checked = False
      Else
         ' Turn on
         sTemp = sTemp & " +p"
         Me.mnuRoomPrivate.Checked = True
         Me.mnuRoomSecret.Checked = False
      End If
      SendServer2 sTemp
End Sub

Private Sub mnuRoomSecret_Click()
Dim sTemp As String

      sTemp = "MODE %#" & FixSpaces(sRoomJoined, True)
      If Me.mnuRoomSecret.Checked = True Then
         ' Turn off
         sTemp = sTemp & " -s"
         Me.mnuRoomSecret.Checked = False
      Else
         ' Turn on
         sTemp = sTemp & " +s"
         Me.mnuRoomSecret.Checked = True
         Me.mnuRoomPrivate.Checked = False
      End If
      SendServer2 sTemp
End Sub
Private Sub mnuTopic_Click()
Dim sTemp As String

      On Error GoTo Hell
      sTemp = InputBox("Please enter the room topic you require", "Set Room Topic", "")
      If sTemp <> "" Then
         sTemp = "TOPIC %#" & FixSpaces(sRoomJoined, True) & " :" & sTemp
         SendServer2 sTemp
      End If
Hell:
End Sub

Public Sub mnuTrace_Click()
      frmTraceOptions.Show 0, Me
End Sub

Private Sub mnuUnlock_Click()
      frmUnlock.sPassword = "NP132262A"
      Load frmUnlock
      frmUnlock.Show 1
      If frmUnlock.Tag = "UNLOCKED" Then
         UnlockMe
      End If
      Unload frmUnlock
End Sub

Private Sub mnuViewAccess_Click()
Dim sTemp As String
Dim i As Integer
Dim asTemp() As String

      sTemp = SendGetResponse("ACCESS %#" & FixSpaces(sRoomJoined, True) & " LIST ", " :End of access entries", False)
      Load frmLists
      asTemp = Split(sTemp, vbCrLf)
      frmLists.lstAcess.Clear
      For i = 0 To UBound(asTemp)
         If Locate(asTemp(i), "End of access entries") = 0 And Trim(asTemp(i)) <> "" Then
            If Locate(asTemp(i), " 804 ") Then
               sTemp = GetAfter(asTemp(i), "%#" & FixSpaces(sRoomJoined, True))
               frmLists.lstAcess.AddItem Trim(sTemp)
            Else
               ProcessData asTemp(i)
               SendServer1 asTemp(i)
            End If
         End If
      Next
      frmLists.Show vbModal
      Unload frmLists
      Set frmLists = Nothing
End Sub

Public Sub mnuWelcomePrefs_Click()
      frmWelcomePrefs.Show , Me
End Sub

Private Sub Picture1_Resize()
      ResizeMe
End Sub

Private Sub Svr1_Close()
      On Error Resume Next
      ' CheckOCX
End Sub

Private Sub Svr1_ConnectionRequest(ByVal requestID As Long)
      MDIMain.Svr1.Close
      MDIMain.Svr1.Accept requestID
      trace "*Started HandShake*"
End Sub
Private Sub Svr1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Dim sTemp As String

      Svr1.GetData sData, vbString
      If Not (bTryJoin) Then
         tmrJoin.Enabled = False
         tmrJoin.Enabled = True
      Else
         tmrJoin.Enabled = False
      End If
      If Locate(sData, "MODE %#") Then
         Exit Sub
      End If
      ' If Locate(sData, "IRCVERS") Then
      ' sTemp = GetLine(sData, "IRCVERS ")
      ' sData = Replace(sData, sTemp, "IRCVERS IRC7 MSN-OCX!8.00.0211.1802")
      ' End If
      If Locate(sData, "JOIN ") Then
         ' If frmPrefs.chkLeaveChatOpen.Value = 0 Then
         CloseServer1
         Exit Sub
         ' End If
      End If
      If Locate(sData, "FINDS %#") > 0 And Locate(sData, "\\") > 0 Then
         sData = Replace(sData, "\\", "\")
      End If
      If Locate(sData, "\\b") Or Locate(sData, "\\c") Or Locate(sData, "\\r") Then
         sData = Replace(sData, "\\", "\")
      End If
      If Locate(sData, "FINDS %#") Then
         sRoomJoined = GetLine(sData, "FINDS %#")
         sRoomJoined = GetAfter(sRoomJoined, "%#")
      End If

      SendServer2 sData
      DoEvents
End Sub
Private Sub Svr1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
      ' CheckOCX
End Sub

Private Sub Svr2_Close()
      Chat "Server Closed Connection", True, eCols.CRed
      Svr1.Close
      bConnected = False
      ' CheckOCX
End Sub

Private Sub Svr2_Connect()
Dim sRoom As String
Dim sNick As String

      On Error Resume Next
            
      sNick = ">" & frmMain.txtNick
      
      bConnecting = True
      trace "Contacted MSN Server Succesfully! (" & Svr2.RemoteHost & ":" & Svr2.RemotePort & ")"
      Svr1.Close
      DoEvents
      
      Svr1.LocalPort = 6667
      ' trace "Local Connection is on 127.0.0.1:" & 6667
      Svr1.Listen
      If Err Then GoTo Hell
      DoEvents
      Exit Sub
         
Hell:
      
112   trace "Opps - Summat wrong (" & Err.Description & ")"
113   trace "Re-trying"

114   iPort = iPort + 1
115   DoConnect False

End Sub
Private Sub Svr2_DataArrival(ByVal bytesTotal As Long)
Dim sGetData As String
Dim sNewAddress As String
Dim iSpace As Integer
Dim sTemp As String
      ' *** CodeSmart ErrorHead TagStart | Please Do Not  Modify
      ' Code Added By CodeSmart
      ' =============================================================================
      On Error GoTo Err_Svr2_DataArrival:
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      ' *** CodeSmart ErrorHead TagEnd | Please Do Not Modify
Static sTotalData As String

101   bError = False
102   Svr2.GetData sGetData, vbString
104   sTotalData = sTotalData & sGetData
      
105   If Right$(sGetData, 2) <> vbCrLf Then
106      Exit Sub
      End If
107   If Locate(sGetData, "Unknown Command") Or Locate(sGetData, "QUIT") Then Exit Sub
108   If Locate(sTotalData, "613 nick :207.") Then
109      trace "<" & sTotalData
         sTemp = Left$(sTotalData, Locate(sTotalData, " :207") + 1)
         sTemp = Mid$(sTotalData, Locate(sTotalData, ":207") + 1, Len(sTotalData))
         sNewAddress = Left$(sTemp, Locate(sTotalData, " "))
110      SendServer1 sTotalData
111      DoEvents
112      Svr2.Close
115      Svr2.Connect sNewAddress, "6667"
         trace "Attempting Connection to new Server"
116      DoEvents
117      sTotalData = ""
118      Exit Sub
      End If
119   If Locate(sGetData, "AUTH GateKeeper *") Or Locate(sGetData, "AUTH GateKeeperPassport *") Then
120      If frmControl.Check1.Value Then
121         If iAuthCount > 1 Then
122            CloseServer1
123            trace sGetData
124            bConnected = True
125            iAuthCount = 0
126            Exit Sub
            Else
127            iAuthCount = iAuthCount + 1
            End If
         End If
      End If
128   If Locate(sGetData, "AUTH GateKeeper *") Or Locate(sGetData, "AUTH GateKeeperPassport *") Then
129      If iAuthCount > 1 And frmControl.chkClone.Value Then
130         SendServer2 "NICK " & frmMain.txtNick & vbCrLf
131         DoEvents
132         Exit Sub
         Else
133         iAuthCount = iAuthCount + 1
         End If
      End If

134   sWaitedFor = sWaitedFor & sGetData
135   If bWaiting Then Exit Sub
136   ProcessData sTotalData
137   SendServer1 sTotalData
138   sTotalData = ""
139   sWaitedFor = ""
      ' *** CodeSmart ErrorFoot TagStart | Please Do Not Modify
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      Exit Sub
Err_Svr2_DataArrival:
      MsgBox ("Error Encounterd in Svr2_DataArrival @ " & Erl & " " & Err.Description)
      ' =============================================================================
Exit_Svr2_DataArrival:
      ' *** CodeSmart ErrorFoot TagEnd | Please Do Not Modify
End Sub
Private Sub Svr2_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
      If bError = False Then
         bError = True
         If Number = 10053 Or Number = 10065 Then
            Chat LoadResString(104), True, eCols.CRed
            ' RefreshConnection False
            ' DoConnect
         Else
            trace "Server 2 Error - " & Number & " - " & Description
         End If
      End If
      trace Description
End Sub
Private Sub mnuOwner_Click()
      frmControl.cmdOwner_Click
End Sub
Private Sub mnuHost_Click()
      frmControl.cmdBrown_Click
End Sub
Private Sub mnuParticipant_Click()
      frmControl.cmdParticipant_Click
End Sub
Private Sub mnuSpectator_Click()
      frmControl.cmdSpec_Click
End Sub
Private Sub mnuTime_Click()
      frmControl.cmdTime_Click
End Sub
Private Sub mnuIdent_Click()
      frmControl.cmdIdent_Click
End Sub
Private Sub mnuProfile_Click()
      frmControl.cmdProfile_Click
End Sub

Private Sub tmrActivity_Timer()
Dim sTemp As String
      
      frmMain.Caption = sChatCaption & " - No Activity for " & iAliveTimer & " Mins"
      ' frmMain.Refresh
      ' If frmPrefs.chkTestAlive.Value Then
      If GeneralSettings.TestAlive Then
         ' If iAliveTimer >= frmPrefs.sldAlive.Value Then
         If iAliveTimer >= GeneralSettings.AliveTime Then
            ' Do Alive Test
            iAliveTimer = 0
            sTemp = SendGetResponse("PING .", "PONG")
            If sTemp = "" Then
               ' Dead
               DoConnect
            End If
         End If
      End If
End Sub


Public Sub Advertise()
Dim sTemp As String
Dim i As Integer
Dim pPrefSettings As ePrefs

      iAliveTimer = iAliveTimer + 1
      If Me.mnuDoAdverts.Checked = True Then
         iAdvertTimer = iAdvertTimer + 1
         If iAdvertTimer > 100 Then iAdvertTimer = 0
      
         With frmAdvertising
            For i = 0 To .Check1.UBound
               If .Check1(i).Value Then
                  .lblInterval(i).Tag = .lblInterval(i).Tag + 1
                  ' Debug.Print I, .sldAdvertise(I).Value, .lblInterval(I).Tag
                  If .lblInterval(i).Tag >= .sldAdvertise(i).Value Then
                     ' If .lblInterval(i).Tag >= .sldAdvertise(i) Then
                     ' Do Advert
                     .lblInterval(i).Tag = 0
                  
                     sTemp = .txtAdvertMessage(i)
                     Chat "Advertised " & sTemp, True, eCols.CBlack
                     Select Case i
                        Case 0
                           pPrefSettings = pAdvert1
                        Case 1
                           pPrefSettings = pAdvert2
                        Case 2
                           pPrefSettings = pAdvert3
                     End Select
                     If .chkAction(i).Value Then
                        sTemp = "ACTION " & sTemp
                        Call PRIVMSG(sTemp, "", False, True, pPrefSettings, True, False)
                     Else
                        Call PRIVMSG(sTemp, "", True, True, pPrefSettings, True, False)
                     End If
                     ' End If
                  End If
               End If
            Next
         End With
      End If
End Sub
Public Sub TestScroll()
Dim iVis As Integer
Dim iTopLine As Long
Dim iTotalLines As Long

      iVis = GetVisibleLines(frmMain.txtChat) + 1
      iTopLine = GetVisibleLine(frmMain.txtChat)
      iTotalLines = GetLineCount(frmMain.txtChat)
      If iTopLine < (iTotalLines - iVis) Then
         bScollingChat = True
      Else
         bScollingChat = False
      End If
      iCounter = iCounter + 1
      If iCounter > 2 Then
         frmMain.imgServer.Picture = MDIMain.ilstServer.ListImages(1).Picture
         SetIcon
         iCounter = 0
      End If
End Sub

Private Sub tmrJoin_Timer()
      If bConnecting = True Then
         '
         DoConnect
      Else
         If bJoined = False Then
            MDIMain.mnuJoin_Click
         End If
      End If
End Sub

Private Sub tmrTimer_Timer()
Static iAdvertCount As Integer
Static iChecker As Integer

      iChecker = iChecker + 1
      ' Debug.Print iChecker
      If iChecker > (60 * 30) Then
         Load frmCheckLatest
         frmCheckLatest.Check
         iChecker = 0
      End If
      If bAdvertise = True Then
         iAdvertCount = iAdvertCount + 1
         If iAdvertCount > 60 Then
            ' 60 Seconds
            Call Advertise
            iAdvertCount = 0
         End If
      End If
      If bJoined = True Then
         Call TestScroll
      End If
End Sub
Public Sub SetIcon()
      If bShowSysTray Then
         m_frmSysTray.IconHandle = frmMain.imgServer.Picture.Handle
      End If
End Sub
Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
      If (eButton = vbRightButton) Then
         m_frmSysTray.ShowMenu
      End If
End Sub
Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
      Me.WindowState = vbNormal
      Me.Show
      Me.ZOrder
End Sub
Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
      ' Debug.Print sKey, lIndex
      Select Case UCase(sKey)
         Case "SHUTDOWN"
            Unload Me
         Case "ABOUT"
            frmAbout.Show 1
         Case "WEB"
            pShell "http://homepage.ntlworld.com/mrenigma", Me
         Case "OPEN"
            Me.WindowState = vbNormal
            Me.Show
            Me.ZOrder
         Case "PREFSGEN"
            mnuPrefs_Click
         Case "PREFSPASSPORT"
            mnuPassPort_Click
         Case "PREFSWELCOME"
            mnuWelcomePrefs_Click
         Case "PREFSKICK"
            mnuAutoPrefs_Click
         Case "PREFSHOST"
            mnuAutoHost_Click
         Case "PREFSTRACE"
            mnuTrace_Click
         Case "PREFSADVERT"
            mnuAdvertising_Click
      End Select
End Sub
Public Sub SetSystray()
      If bShowSysTray Then
         If TypeName(m_frmSysTray) = "Nothing" Then
            Set m_frmSysTray = New frmSysTray
            Load m_frmSysTray
            ' With m_frmSysTray
            ' .AddMenuItem "&Open IRCDominator", "OPEN", True
            ' .AddMenuItem "-"
            ' .AddMenuItem "Preferences", "Prefs", False
            ' .AddMenuItem "&Enigma Wares Home Page", "Web"
            ' .AddMenuItem "&About...", "About"
            ' .AddMenuItem "-"
            ' .AddMenuItem "&Shutdown IRCDominator", "EXIT"
            ' .ToolTip = "IRCDominator"
            ' End With
            m_frmSysTray.ToolTip = "IRCDominator"
            SetIcon
         End If
      Else
         On Error Resume Next
         Unload m_frmSysTray
         Set m_frmSysTray = Nothing
      End If
End Sub
