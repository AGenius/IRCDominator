VERSION 5.00
Begin VB.Form frmKickBan 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Kick/Ban"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmKickBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optBanPassport 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ban PassPort For"
      Height          =   285
      Left            =   1410
      TabIndex        =   7
      Top             =   900
      Width           =   1695
   End
   Begin VB.ComboBox cboMessage 
      Height          =   315
      ItemData        =   "frmKickBan.frx":08CA
      Left            =   90
      List            =   "frmKickBan.frx":08DA
      TabIndex        =   0
      Top             =   450
      Width           =   5055
   End
   Begin VB.ComboBox cboBans 
      Height          =   315
      ItemData        =   "frmKickBan.frx":0913
      Left            =   3120
      List            =   "frmKickBan.frx":0928
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   900
      Width           =   2025
   End
   Begin VB.OptionButton optBanNick 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ban Nick For"
      Height          =   285
      Left            =   1410
      TabIndex        =   2
      Top             =   1200
      Width           =   1875
   End
   Begin VB.OptionButton optNoBan 
      BackColor       =   &H00FFC0C0&
      Caption         =   "No Ban"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2610
      TabIndex        =   5
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1020
      TabIndex        =   4
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kick Message"
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
      Left            =   105
      TabIndex        =   6
      Top             =   150
      Width           =   1200
   End
End
Attribute VB_Name = "frmKickBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bSilent As Boolean

Private Sub cmdCancel_Click()
      Me.Hide
      Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim iLoop As Long
Dim sList As String
Dim asList() As String
Dim iTime As Integer

      On Error GoTo Hell
      sList = BuildNameList
      asList = Split(sList, ",")
      For i = 0 To UBound(asList)
         sNick = asList(i)
         If Me.optNoBan.Value Then
            Call DoKick(sNick, Me.cboMessage, False, , , , bSilent)
         Else
            iTime = Me.cboBans.ItemData(Me.cboBans.ListIndex)
            If Me.optBanPassport.Value Then
               ' Ban Passport
               Call DoKick(sNick, Me.cboMessage, True, True, iTime, cboBans.List(cboBans.ListIndex), bSilent)
            Else
               ' Ban Nick Name
               Call DoKick(sNick, Me.cboMessage, True, False, iTime, cboBans.List(cboBans.ListIndex), bSilent)
            End If
         End If
         iLoop = iLoop + 1
         If iLoop > 5 Then
            For iLoop = 1 To 1000000
               DoEvents
            Next
            iLoop = 0
         End If
      Next
Hell:
      Me.Hide
      Unload Me
End Sub
Private Sub Form_Activate()
      If bSilent Then
         Me.cboBans.Enabled = True
         Me.cboMessage.Text = "Silent Ban"
         Me.cboMessage.Locked = True
         Me.cboBans.ListIndex = Me.cboBans.ListCount - 1
         Me.optNoBan.Value = False
         Me.optNoBan.Enabled = False
         Me.optBanPassport.Value = True
      End If
End Sub
Private Sub Form_Load()
      Call BuildBanList(cboBans)
      Me.cboBans.ListIndex = 0
      Me.cboBans.Enabled = False
End Sub
Private Sub optBanNick_Click()
      If Me.optBanNick.Value Then
         Me.cboBans.Enabled = True
      End If
End Sub
Private Sub optBanPassport_Click()
      If Me.optBanPassport.Value Then
         Me.cboBans.Enabled = True
      End If
End Sub
Private Sub optNoBan_Click()
      If Me.optNoBan.Value Then
         Me.cboBans.Enabled = False
      End If
End Sub
