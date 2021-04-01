VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWhisper 
   BackColor       =   &H00FFC0C0&
   Caption         =   "%n : 1 to 1 Whisper"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   Icon            =   "frmWhisper.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdProfile 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5880
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmWhisper.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin RichTextLib.RichTextBox txtNickName 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   150
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   556
      _Version        =   393217
      BackColor       =   16761024
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      MultiLine       =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmWhisper.frx":0884
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
   Begin VB.CommandButton cmdSend 
      Enabled         =   0   'False
      Height          =   405
      Left            =   5220
      Picture         =   "frmWhisper.frx":0902
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2940
      Width           =   1185
   End
   Begin RichTextLib.RichTextBox txtCommand 
      Height          =   405
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "RES|171"
      Top             =   2940
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   714
      _Version        =   393217
      TextRTF         =   $"frmWhisper.frx":1BF8
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
   Begin RichTextLib.RichTextBox txtMessages 
      Height          =   2175
      Left            =   90
      TabIndex        =   2
      Top             =   615
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmWhisper.frx":1C73
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Height          =   2235
      Left            =   510
      Top             =   600
      Width           =   6225
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2235
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   630
      TabIndex        =   3
      Top             =   150
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   0
      Picture         =   "frmWhisper.frx":1CEE
      Stretch         =   -1  'True
      Top             =   30
      Width           =   570
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear All"
      End
   End
End
Attribute VB_Name = "frmWhisper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objTooltip    As cTooltip
Private m_NickName As String
Private m_WindowIndex As Integer
Private sTitleBar As String
Dim sSentText(100) As String
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

Public Property Let WindowIndex(vData As Integer)
      m_WindowIndex = vData
End Property
Public Property Get WindowIndex() As Integer
      WindowIndex = m_WindowIndex
End Property

Public Property Let Nickname(vData As String)
      m_NickName = vData
End Property
Public Property Get Nickname() As String
      Nickname = m_NickName
End Property


Private Sub cmdProfile_Click()
Dim sTemp As String
Dim sProf As String
Dim sNick As String

      On Error GoTo Hell

      sProf = "http://members.msn.com/Default.msnw?mpp=2208~%p&mid=2200"
      sNick = TestNick(m_NickName, True)

      sTemp = SendGetResponse("PROP " & sNick & " PUID", " :End of properties")
      sTemp = Split(sTemp, "PUID :")(1)
      sTemp = Split(sTemp, vbCrLf)(0)
      
      sProf = Replace(sProf, "%p", sTemp)
      
      MDIMain.WB.Navigate "about: blank"
      DoEvents
      MDIMain.WB.Navigate ("javascript:window.open('" & sProf & "')")
Hell:

End Sub

Private Sub cmdSend_Click()
Dim sTalk As String

      If cmdSend.Enabled = True Then
         cmdSend.Enabled = False
        
         On Error Resume Next
         Call WHISPER(txtCommand.Text, m_NickName, True, pChat)
         Chat "   " & TestNick(sNickJoined, False) & " : ", False, eCols.CNavy, 2, , txtMessages
         Chat txtCommand.Text, True, FindColour(GeneralSettings.Chat_Colour), GetStyle(pChat), GeneralSettings.Chat_Font, txtMessages
         Call RecordSent(txtCommand.Text)
         txtCommand = ""
      End If
End Sub
Private Sub Form_Activate()
      If sTitleBar = "" Then
         sTitleBar = Me.Caption
      End If
      Me.txtCommand.SetFocus
      If Left$(TestNick(m_NickName, True), 1) = ">" Then
         Me.cmdProfile.Visible = False
      End If
End Sub
Private Sub Form_Load()
      LoadToolTips Me, m_objTooltip
      
      ' If bActivated Then
      ' '         Me.cmdNuke.Visible = True
      ' Else
      ' Me.cmdNuke.Value = False
      ' End If
End Sub
Private Sub Form_Resize()
      On Error Resume Next
      Shape1.Height = Me.ScaleHeight - Me.txtCommand.Height - 150 - Shape1.Top
      Me.txtCommand.Top = Me.ScaleHeight - Me.txtCommand.Height - 70
      Me.txtMessages.Height = Shape1.Height - 35
      Shape2.Height = Shape1.Height
      Shape2.Width = Me.ScaleWidth - Shape2.Left
      Me.txtCommand.Width = Me.ScaleWidth - 200
      Me.cmdSend.Top = Me.txtCommand.Top
      Me.cmdSend.Left = Me.ScaleWidth - Me.cmdSend.Width - 200
      Me.txtCommand.Width = Me.ScaleWidth - Me.cmdSend.Width - 400
      Me.txtNickName.Width = Me.ScaleWidth - Me.txtNickName.Left
      Me.txtMessages.Width = Me.ScaleWidth - 80
      Me.txtNickName.Width = Me.Width - Image1.Width - 600
      Me.cmdProfile.Left = Me.Width - Me.cmdProfile.Width - 150
      ' Me.cmdNuke.Left = Me.Width - Me.cmdNuke.Width - 150
End Sub
Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
      If m_WindowIndex <> -1 Then
         Set frmWhisperForm(m_WindowIndex) = Nothing
      Else
         Set frmWhisper = Nothing
      End If
End Sub

Private Sub txtCommand_Change()
      If txtCommand.Text <> "" Then
         cmdSend.Enabled = True
      Else
         cmdSend.Enabled = False
      End If
End Sub
Private Sub txtCommand_KeyPress(KeyAscii As Integer)
      On Error Resume Next
      If KeyAscii = 13 Then
         cmdSend_Click
         KeyAscii = 0
      End If
End Sub
Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 38 Or KeyCode = 40 Then
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
End Sub

Private Sub mnuSelectAll_Click()
      txtMessages.SelStart = 0
      txtMessages.SelStart = 0
      txtMessages.SelLength = Len(txtMessages.Text)
End Sub
Private Sub mnuClear_Click()
      txtMessages.Text = ""
End Sub
Private Sub mnuCopy_Click()
      Clipboard.SetText (txtMessages.SelText)
End Sub
Private Sub txtMessages_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      On Error GoTo Hell
      If Button = 2 Then
         If txtMessages.Text = "" Then
            mnuClear.Enabled = False
            mnuSelectAll.Enabled = False
         Else
            mnuClear.Enabled = True
            mnuSelectAll.Enabled = True
         End If
         If txtMessages.SelLength > 0 Then
            mnuCopy.Enabled = True
         Else
            mnuCopy.Enabled = False
         End If
         PopupMenu Me.mnuPopup
      End If
Hell:

End Sub
