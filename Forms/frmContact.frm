VERSION 5.00
Begin VB.Form frmContact 
   BackColor       =   &H00CFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact Details"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5115
   Icon            =   "frmContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3810
      TabIndex        =   0
      Top             =   990
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   2
      Left            =   105
      Picture         =   "frmContact.frx":08CA
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   105
      Picture         =   "frmContact.frx":1794
      Top             =   570
      Width           =   720
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email: theabsolutegenius@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   930
      MouseIcon       =   "frmContact.frx":265E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2460
      Width           =   3150
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Index           =   3
      Left            =   930
      TabIndex        =   2
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   2580
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enigma Ware Contact Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   105
      Picture         =   "frmContact.frx":2AA0
      Top             =   1380
      Width           =   720
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Image1_Click(Index As Integer)
      frmUnlock.sPassword = "TheOwnerGenius"
      ' MDIMain.AL1.Password = Encrypt1("E†VôwæV'ÒtVæ–W7")
      Load frmUnlock
      If Not (bActivated) Then
         frmUnlock.Show 1
      End If
      If frmUnlock.Tag = "UNLOCKED" Then
         bActivated = True
         ' Unlock Owner stuff here
         UnlockMe
      End If
      Unload frmUnlock
End Sub

Private Sub lblEmail_Click()
Dim sEmailTo As String

      sEmailTo = "mailto:theabsolutegenius@hotmail.com?subject=IRC Dominator "
      Call ShellExecute(hwnd, "Open", sEmailTo, "", "", 1)
End Sub

Private Sub OKButton_Click()
      Me.Hide
End Sub
