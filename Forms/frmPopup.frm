VERSION 5.00
Begin VB.Form frmPopup 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   0
      Width           =   3045
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "dfasffdsfasdfasdf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1305
         Left            =   330
         TabIndex        =   1
         Top             =   210
         Width           =   2235
      End
      Begin VB.Image Image1 
         Height          =   1860
         Left            =   0
         Picture         =   "frmPopup.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2970
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3270
      Top             =   30
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Type RECT
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
End Type

Dim i As Integer
Dim taskbar As Long
Dim MyHeight As Long
Public iStayUp As Long
Public sDo As String

' Timer1 Interval = 15 (may need To be adjusted)
' Timer2 Interval = time To keep form showing (I used 4000)
' Timer3 Interval = the same as Timer1

Private Sub Form_Load()
Dim NormalWindowStyle As Long
Dim WindowRect As RECT

      On Error Resume Next

      MyHeight = Me.Height
      ' Get the normal window style to or layered property to
      ' NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
      ' ' Make windows 2000/XP recognize window as layered window
      ' SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
      ' ' Make windows 2000/XP change transparency level
      ' SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
      ' set i to 100% (for transparency)
      i = 100
      SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
      taskbar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.Bottom) * Screen.TwipsPerPixelX

      Me.Top = Screen.Height - taskbar
      Me.Height = 0
      Me.Left = Screen.Width - Me.Width
      Timer1.Interval = 1
      Timer1.Tag = "UP"
      StayOnTop Me.hwnd, True

      ' Timer1.Enabled = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lblMessage.ForeColor = &HFF&
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lblMessage.ForeColor = &HFF&

End Sub

Private Sub lblMessage_Click()
      If sDo = "CHECK" Then
         ' MDIMain.Hide
         ' DoEvents
         Shell App.Path & "\AutoUpdater.exe", vbNormalFocus

         ' frmUpdate.Show
      End If
      ' StayOnTop Me.hwnd, False
      Timer1.Enabled = False
      Unload Me
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      lblMessage.ForeColor = &HFF8080
End Sub

' Timer that moves form up
Private Sub Timer1_Timer()
   
      On Error Resume Next
      DoEvents
      Select Case Timer1.Tag

         Case "UP"
            ' get the size of the taskbar
            ' move form up until it sits on top of taskbarPrivate
            If (Me.Top + MyHeight + taskbar) > Screen.Height Then
               Me.Top = Me.Top - 35
               Me.Height = Me.Height + 35
               DoEvents
            Else
               Timer1.Enabled = False
               Timer1.Interval = iStayUp
               Timer1.Tag = "WAIT"
               Timer1.Enabled = True
            End If
         Case "WAIT"
            Timer1.Enabled = False
            Timer1.Interval = 1
            Timer1.Tag = "DOWN"
            Timer1.Enabled = True
         Case "DOWN"
            If Me.Top < Screen.Height - taskbar And i > 0 Then
               Me.Top = Me.Top + 40
               Me.Height = Me.Height - 40
               i = i - 1
               ' MakeTransparent (I)
            Else
               Timer1.Enabled = False
               Timer1.Tag = "DONE"
               StayOnTop Me.hwnd, False
               If sDo = "CHECK" Then
                  Unload frmCheckLatest
                  Set frmCheckLatest = Nothing
               End If
               Unload Me
            End If
      End Select
End Sub

' routine that makes a form transparent
' TransAmount ranges from 0 to 100, resem bles a percentage
' Private Sub MakeTransparent(TransAmount As Integer)
' Dim Transparency As Byte
'
' ' Get new transparency level
' Transparency = (255 * TransAmount) / 100
' ' Set the new transparency level
' SetLayeredWindowAttributes Me.hwnd, 0, Transparency, LWA_ALPHA
' End Sub

