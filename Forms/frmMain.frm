VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F7BA9F11-0A5D-11D0-97C9-0000C09400C4}#2.0#0"; "Splitter.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C000&
   Caption         =   "Main Chat Window"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11370
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   11370
   Begin SSSplitter.SSSplitter SPLITTER 
      Height          =   6510
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   11483
      _Version        =   131074
      AutoSize        =   1
      SplitterResizeStyle=   1
      SplitterBarAppearance=   0
      BorderStyle     =   1
      BackColor       =   16761024
      PaneTree        =   "frmMain.frx":22A2
      Begin VB.PictureBox pctTop 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   15
         ScaleHeight     =   660
         ScaleWidth      =   11340
         TabIndex        =   13
         Top             =   15
         Width           =   11340
         Begin VB.CheckBox chkWelcomeAway 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Welcome Back Aways"
            Height          =   285
            Left            =   6810
            TabIndex        =   24
            Tag             =   "Welcome|Away"
            ToolTipText     =   "STR|Select this to turn on welcome back messages"
            Top             =   270
            Width           =   2025
         End
         Begin VB.CheckBox chkWelcome 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Welcome Guests"
            Height          =   285
            Left            =   6810
            TabIndex        =   23
            Tag             =   "Welcome|Active"
            ToolTipText     =   "STR|Select this to turn on the welcome message"
            Top             =   30
            Width           =   1635
         End
         Begin VB.CheckBox chkPassport 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Use Passport"
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
            Left            =   4470
            TabIndex        =   22
            ToolTipText     =   "STR|Select this to join the room using your passport"
            Top             =   360
            Width           =   2775
         End
         Begin VB.CheckBox chkHex 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Use Hex"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2010
            TabIndex        =   16
            ToolTipText     =   "STR|Select this option if you are joining a room\nusing hex codes"
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtNick 
            Height          =   285
            Left            =   4440
            TabIndex        =   15
            Text            =   "DreamChild"
            ToolTipText     =   "RES|173"
            Top             =   0
            Width           =   2265
         End
         Begin VB.TextBox txtRoomName 
            Height          =   285
            Left            =   1110
            TabIndex        =   14
            Text            =   "DreamRoom"
            ToolTipText     =   "RES|172"
            Top             =   0
            Width           =   2265
         End
         Begin MSComctlLib.ImageList Icons 
            Left            =   11190
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":237E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2918
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":2EB2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":344C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":39E6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":3E38
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Image imgServer 
            Height          =   480
            Left            =   8820
            Picture         =   "frmMain.frx":428A
            Top             =   60
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pass Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   21
            Top             =   210
            Width           =   915
         End
         Begin VB.Label lblVersion4Pass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PassCode Here"
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
            Height          =   195
            Left            =   30
            TabIndex        =   20
            Top             =   420
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nick Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   3450
            TabIndex        =   18
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   17
            Top             =   30
            Width           =   1035
         End
      End
      Begin MSComctlLib.ListView tUsers 
         CausesValidation=   0   'False
         Height          =   4710
         Left            =   8655
         TabIndex        =   10
         ToolTipText     =   "STR|The chatters list"
         Top             =   1620
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   8308
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         Icons           =   "Icons"
         SmallIcons      =   "Icons"
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Host"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RealNick"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.PictureBox pctButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   8655
         ScaleHeight     =   90
         ScaleWidth      =   2700
         TabIndex        =   9
         Top             =   6405
         Width           =   2700
      End
      Begin VB.PictureBox pctTalk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   15
         ScaleHeight     =   525
         ScaleWidth      =   8565
         TabIndex        =   8
         Top             =   5970
         Width           =   8565
         Begin VB.CommandButton cmdAction 
            Enabled         =   0   'False
            Height          =   375
            Left            =   6780
            Picture         =   "frmMain.frx":4594
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "RES|157"
            Top             =   60
            Width           =   375
         End
         Begin VB.CommandButton cmdSend 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5580
            Picture         =   "frmMain.frx":4BBE
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "RES|155"
            Top             =   60
            Width           =   1155
         End
         Begin VB.TextBox txtCommand 
            Height          =   375
            Left            =   90
            TabIndex        =   0
            ToolTipText     =   "RES|171"
            Top             =   45
            Width           =   5355
         End
      End
      Begin VB.PictureBox pctChat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5145
         Left            =   15
         ScaleHeight     =   5145
         ScaleWidth      =   8565
         TabIndex        =   7
         Top             =   750
         Width           =   8565
         Begin RichTextLib.RichTextBox txtChat 
            Height          =   4695
            Left            =   120
            TabIndex        =   12
            Top             =   70
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   8281
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmMain.frx":5EB4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4725
            Left            =   570
            Top             =   60
            Width           =   7935
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   4725
            Left            =   60
            Shape           =   4  'Rounded Rectangle
            Top             =   60
            Width           =   585
         End
      End
      Begin VB.PictureBox pctMe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   8655
         ScaleHeight     =   795
         ScaleWidth      =   2700
         TabIndex        =   4
         Top             =   750
         Width           =   2700
         Begin VB.CommandButton cmdAway 
            BackColor       =   &H00FFFFC0&
            Height          =   435
            Left            =   1830
            MaskColor       =   &H00FFFFC0&
            Picture         =   "frmMain.frx":5F2F
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "UNAWAY"
            ToolTipText     =   "RES|162"
            Top             =   0
            Width           =   495
         End
         Begin MSComctlLib.ListView tMe 
            CausesValidation=   0   'False
            Height          =   435
            Left            =   30
            TabIndex        =   5
            ToolTipText     =   "STR|This is you"
            Top             =   480
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   767
            SortKey         =   1
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            Icons           =   "Icons"
            SmallIcons      =   "Icons"
            ForeColor       =   0
            BackColor       =   16761024
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Host"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblUsers 
            AutoSize        =   -1  'True
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "People Chatting"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   90
            TabIndex        =   6
            Top             =   90
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label lblBlank 
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   4230
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Implements ISubClass

' Dim iCount              As Integer
Private m_objTooltip    As cTooltip
Dim sSentText(500) As String
Dim iCounter As Integer
Dim iScrollCounter As Integer

Private Sub RecordSent(sMessage As String)
Dim i As Integer
      sSentText(iCounter) = txtCommand.Text
      iCounter = iCounter + 1
      iScrollCounter = iCounter
      If iCounter > 100 Then
         ' Ripple Back
         For i = 1 To UBound(sSentText)
            sSentText(i - 1) = sSentText(i)
         Next
      End If
End Sub

Private Sub chkWelcome_Click()
      frmWelcomePrefs.chkWelcome.Value = chkWelcome.Value
End Sub

Private Sub chkWelcomeAway_Click()
      frmWelcomePrefs.chkWelcomeAway.Value = chkWelcomeAway.Value
End Sub
Private Sub cmdAction_Click()
Dim sTalk As String
        
      cmdAction.Enabled = False
        
      On Error Resume Next
      If Left$(txtCommand, 1) <> "/" Then
         Call PRIVMSG("ACTION " & txtCommand, "", False, True, pChat)
         Chat "   " & TestNick(sNickJoined, False) & " " & txtCommand, True, eCols.CPurple, 3
         Call RecordSent(txtCommand.Text)
         txtCommand = ""
      End If
End Sub
Private Sub cmdAway_Click()
      If bConnected Then
         If cmdAway.Tag = "UNAWAY" Then
            cmdAway.Tag = "AWAY"
            SendServer2 "AWAY :AWAY"
            MEAway True
         Else
            cmdAway.Tag = "UNAWAY"
            SendServer2 "AWAY"
            MEAway False
         End If
      End If
End Sub
Private Sub cmdSend_Click()
Dim sTalk As String
Dim sCommand As String

      If cmdSend.Enabled = True Then
        
         On Error Resume Next
         If Left$(txtCommand, 1) <> "/" Then
            Call PRIVMSG(txtCommand, "", True, True, pChat, True)
            Chat "   " & TestNick(sNickJoined, False) & " : ", False, eCols.CBlack, 2
            Chat txtCommand, True, FindColour(GeneralSettings.Chat_Colour), GetStyle(False), GeneralSettings.Chat_Font
            Call RecordSent(txtCommand.Text)
            txtCommand = ""
            cmdSend.Enabled = False

         Else
            ' Send Raw
            sTalk = Mid$(txtCommand, 2, Len(txtCommand))
            sCommand = UCase(Mid$(sTalk & "      ", 1, Locate(sTalk, " ")))
            Select Case sCommand
               Case "ACCESS "
                  If Locate(sTalk, "%#") = 0 Then
                     ' Add Room Name
                     sTalk = sCommand & " %#" & sRoomJoined & " " & Mid$(sTalk, 8, Len(sTalk))
                  End If
               Case "PASS "
                  ' Enter Room Pass
                  DoPass
                  txtCommand.Text = ""
                  Exit Sub
            End Select
            If bActivated = False Then
               ' Take out kill char
               sTalk = Replace(sTalk, Chr(1), "")
            End If
            SendServer2 sTalk
            If UCase$(Left$(sTalk, 5)) = "JOIN " Then
               sRoomJoined = Mid$(sTalk, 8, Len(sTalk))
            End If
         End If
      End If
End Sub

Private Sub Form_Activate()
      Me.txtCommand.SetFocus
End Sub
Private Sub Form_Load()
Dim lngEventMask   As Long
Dim T As Long
Dim sNew As String
Dim sDiv As String

      ' *** CodeSmart ErrorHead TagStart | Please Do Not  Modify
      ' Code Added By CodeSmart
      ' =============================================================================
      On Error GoTo Err_Form_Load:
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      ' *** CodeSmart ErrorHead TagEnd | Please Do Not Modify

      With frmMain.txtChat
         ' lngEventMask = SendMessage(.hwnd, EM_GETEVENTMASK, 0, ByVal CLng(0))
         ' If lngEventMask Xor ENM_LINK Then
         ' lngEventMask = lngEventMask Or ENM_LINK
         ' End If
         ' T = SendMessage(.hwnd, EM_SETEVENTMASK, 0, ByVal CLng(lngEventMask))
         T = SendMessage(.hwnd, EM_AUTOURLDETECT, CLng(1), ByVal CLng(0))
      
      End With
      '
      ' glngOriginalhWnd = frmMain.hwnd
      '
      ' T = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, AddressOf RichTextBoxSubProc)

100   Me.Caption = sChatCaption
101   LoadToolTips Me, m_objTooltip
102   iPort = 6667
103   tUsers.FlatScrollBar = False
104   RefreshPass
105   Me.txtRoomName = fGetIni("LastUsed", "Last Room", "DreamRoom")
106   Me.txtNick = fGetIni("LastUsed", "Last Nick", "DreamChild")
107   Me.chkHex = fGetIni("LastUsed", "Hex Room", 0)


      ' *** CodeSmart ErrorFoot TagStart | Please Do Not Modify
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      Exit Sub
Err_Form_Load:
      MsgBox ("Error Encounterd in Form_Load @ " & Erl & " " & Err.Description)
      ' =============================================================================
Exit_Form_Load:
      ' *** CodeSmart ErrorFoot TagEnd | Please Do Not Modify
End Sub
Public Sub RefreshPass()
      lblVersion4Pass = CStr(ReadRegistry(HKEY_CURRENT_USER, sPassRoot & "\4.0", "userdata1"))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         Cancel = True
         Me.WindowState = vbMinimized
      End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy

Dim T As Long

      T = SetWindowLong(glngOriginalhWnd, GWL_WNDPROC, glnglpOriginalWndProc)
      
      Call PutIni("LastUsed", "Last Room", Me.txtRoomName)
      Call PutIni("LastUsed", "Last Nick", Me.txtNick)
      Call PutIni("LastUsed", "Hex Room", Me.chkHex)
End Sub
Private Sub SPLITTER_Resize(ByVal BorderPanes As SSSplitter.Panes)
      On Error Resume Next
      tUsers.ColumnHeaders.Item(1).Width = pctMe.Width - 400
      Shape1.Height = pctChat.Height - 100
      Shape2.Height = Shape1.Height
      Shape2.Width = pctChat.Width
      txtChat.Width = pctChat.Width - 120
      txtChat.Height = Shape1.Height - 40
    
      cmdAway.Left = pctMe.Width - cmdAway.Width - 100
      tMe.Width = cmdAway.Left
      tMe.ColumnHeaders.Item(1).Width = tMe.Width - 50
      lblBlank.Width = pctMe.Width
        
      Me.cmdSend.Left = pctTalk.Width - Me.cmdAction.Width - Me.cmdSend.Width - 150
      Me.cmdAction.Left = pctTalk.Width - Me.cmdAction.Width - 100
      Me.txtCommand.Width = cmdSend.Left - 100
End Sub
Private Sub tMe_GotFocus()
Dim i As Integer
      On Error GoTo Hell
      For i = 1 To tUsers.ListItems.Count
         tUsers.ListItems(i).Selected = False
      Next
Hell:

End Sub
Private Sub tMe_ItemClick(ByVal Item As MSComctlLib.ListItem)
      DoMyOpTest True
End Sub
Private Sub tMe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error GoTo Hell
      If Button = 2 Then
         If tMe.ListItems.Count > 0 Then
            DoMyOpTest True
            PopupMenu MDIMain.mnuPopup
         End If
      End If
Hell:
End Sub
Private Sub tUsers_GotFocus()
      On Error Resume Next
      tMe.ListItems(1).Selected = False
End Sub
Private Sub tUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
      DoMyOpTest False
      On Error GoTo Hell
      Exit Sub
Hell:
      frmControl.cmdTime.Enabled = False
      frmControl.cmdWhisper.Enabled = False
      frmControl.cmdIdent.Enabled = False
      frmControl.cmdProfile.Enabled = False
      MDIMain.mnuTime.Enabled = False
      MDIMain.mnuWhisper.Enabled = False
      MDIMain.mnuIdent.Enabled = False
      MDIMain.mnuProfile.Enabled = False
      MDIMain.mnuKick.Enabled = False
End Sub
Private Sub tUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error GoTo Hell
      If Button = 2 Then
         Debug.Print X, Y
         If tUsers.ListItems.Count > 0 And tUsers.SelectedItem.Index > 0 Then
            DoMyOpTest False
            PopupMenu MDIMain.mnuPopup
         End If
      End If
Hell:
End Sub
Private Sub txtChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error GoTo Hell
      If Button = 2 Then
         If txtChat.Text = "" Then
            MDIMain.mnuChatClear.Enabled = False
            MDIMain.mnuChatSelectAll.Enabled = False
         Else
            MDIMain.mnuChatClear.Enabled = True
            MDIMain.mnuChatSelectAll.Enabled = True
         End If
         If txtChat.SelLength > 0 Then
            MDIMain.mnuChatCopy.Enabled = True
         Else
            MDIMain.mnuChatCopy.Enabled = False
         End If
         PopupMenu MDIMain.mnuChatPopup
      End If
Hell:
      

End Sub

Private Sub txtChat_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
      Stop
End Sub

Private Sub txtCommand_Change()
'      If bConnected Then
         If txtCommand.Text <> "" Then
            cmdSend.Enabled = True
            cmdAction.Enabled = True
         Else
            cmdAction.Enabled = False
            cmdSend.Enabled = False
         End If
'      End If
End Sub
Private Sub txtCommand_KeyPress(KeyAscii As Integer)
      On Error Resume Next
      If KeyAscii = 13 Then
         cmdSend_Click
         KeyAscii = 0
      End If
End Sub
Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 65 Or KeyCode = 83 Or KeyCode = 87 Then
         KeyCode = 0
      End If
End Sub
Private Sub txtCommand_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = 38 Then
         KeyCode = 0
         ' Up pressed
         iScrollCounter = iScrollCounter - 1
         If iScrollCounter < 0 Then
            iScrollCounter = 0
            Exit Sub
         End If
         txtCommand.Text = sSentText(iScrollCounter)
      End If
      If KeyCode = 40 Then
         KeyCode = 0
         ' Up pressed
         iScrollCounter = iScrollCounter + 1
         If iScrollCounter > iCounter Then
            iScrollCounter = iCounter
            Exit Sub
         End If
         txtCommand.Text = sSentText(iScrollCounter)
      End If
      If txtCommand.Text <> "" Then
         If KeyCode = 65 And Shift = 2 Then
            ' Ctrl A - Action
            cmdAction_Click
            KeyCode = 0
         End If
         If KeyCode = 87 And Shift = 2 Then
            ' Ctrl W - Whisper
            KeyCode = 0
         End If
         If KeyCode = 83 And Shift = 2 Then
            ' Ctrl S - Send
            cmdSend_Click
            KeyCode = 0
         End If
      End If
End Sub
