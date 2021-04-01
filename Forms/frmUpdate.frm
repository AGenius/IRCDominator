VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update - Step 1"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5835
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   3
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton Next 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4290
      TabIndex        =   2
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Image imgLogo 
      Height          =   3300
      Left            =   120
      Picture         =   "frmUpdate.frx":0ECA
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   3660
      TabIndex        =   12
      Top             =   60
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "The MSN Chat Room Manager."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   1020
      Width           =   2265
   End
   Begin VB.Image imgIcon 
      Height          =   630
      Left            =   330
      Picture         =   "frmUpdate.frx":77AC
      Stretch         =   -1  'True
      Top             =   360
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enigma Ware's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   480
      Index           =   0
      Left            =   1290
      TabIndex        =   8
      Top             =   120
      Width           =   2865
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   4470
      TabIndex        =   7
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "IRCDominator"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   1350
      TabIndex        =   6
      Top             =   480
      Width           =   4275
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   720
      Width           =   2115
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H008080FF&
      BorderWidth     =   16
      Height          =   945
      Index           =   1
      Left            =   180
      Top             =   210
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Live Update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1710
      TabIndex        =   4
      Top             =   1710
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1620
      X2              =   5610
      Y1              =   4470
      Y2              =   4470
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IRCDominator Live Update. Click on Next to start the live update process."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1740
      TabIndex        =   1
      Top             =   2850
      Width           =   3615
   End
   Begin VB.Label Welcometext 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1740
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1620
      X2              =   5610
      Y1              =   4485
      Y2              =   4485
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   780
      Index           =   3
      Left            =   1050
      TabIndex        =   11
      Top             =   840
      Width           =   4065
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FF8080&
      Height          =   225
      Index           =   4
      Left            =   1050
      TabIndex        =   10
      Top             =   540
      Width           =   4065
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   2
      Left            =   4830
      Top             =   180
      Width           =   555
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   0
      Left            =   4830
      Top             =   1200
      Width           =   555
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
      frmUpdate.Visible = False ' if u press the cancel button on the first form closes it and returns to to the main form
      Unload frmUpdate2
      Unload frmUpdate3
      Unload Me
End Sub

Private Sub Form_Load()
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
StayOnTop Me.hwnd, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         frmUpdate.Visible = False ' if u press the cancel button on the first form closes it and returns to to the main form
      End If
End Sub

Private Sub Next_Click()
      frmUpdate2.Show
      frmUpdate.Hide
      frmUpdate2.NextButton.Enabled = False
      frmUpdate2.Back.Enabled = True
      frmUpdate2.Connectionstatus.Caption = "Connecting..."
      frmUpdate2.Check
End Sub
