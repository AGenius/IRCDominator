VERSION 5.00
Begin VB.Form frmGuestBan 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guest Deny Options"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmGuestBan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Remove"
      Height          =   390
      Left            =   3420
      TabIndex        =   5
      Top             =   690
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   300
      TabIndex        =   4
      Top             =   690
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1890
      TabIndex        =   3
      Top             =   690
      Width           =   1140
   End
   Begin VB.OptionButton optNoBan 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Permenant"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   150
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.ComboBox cboBans 
      Height          =   315
      ItemData        =   "frmGuestBan.frx":0ECA
      Left            =   3000
      List            =   "frmGuestBan.frx":0EEF
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   2025
   End
   Begin VB.OptionButton optBan 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ban Guests For"
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   150
      Width           =   1695
   End
End
Attribute VB_Name = "frmGuestBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
        Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim sTemp As String
Dim iTime As Integer

        iTime = Me.cboBans.ItemData(Me.cboBans.ListIndex)

        If MsgBox("Are You Sure", vbYesNo, "Enter Guest Ban Access") = vbYes Then
            sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " ADD DENY >* "
            If Me.optBan.Value Then
                sTemp = sTemp & iTime
            End If
            SendServer2 sTemp
            Me.Hide
        End If
End Sub

Private Sub Command1_Click()
Dim sTemp As String

        If MsgBox("Are You Sure", vbYesNo, "Clear Guest Ban Access") = vbYes Then
            sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " DELETE DENY >*"
            SendServer2 sTemp
            Me.Hide
        End If
End Sub

Private Sub Form_Load()
        Me.cboBans.ListIndex = 0
        Me.cboBans.Enabled = False
End Sub

Private Sub optBan_Click()
        If Me.optBan.Value Then
            Me.cboBans.Enabled = True
        End If
End Sub

Private Sub optNoBan_Click()
        If Me.optNoBan.Value Then
            Me.cboBans.Enabled = False
        End If
End Sub
