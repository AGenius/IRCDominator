VERSION 5.00
Begin VB.Form frmUnlock 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IRC Dominator Unlock"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "Unlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6630
      Top             =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   6045
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   7245
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   2220
         TabIndex        =   16
         Top             =   150
         Visible         =   0   'False
         Width           =   4755
      End
      Begin VB.Frame frameRegister 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Unlock Code"
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
         Height          =   2145
         Left            =   270
         TabIndex        =   5
         Top             =   2280
         Width           =   6735
         Begin VB.CommandButton cmdRequest 
            Caption         =   "Request Unlock"
            Height          =   345
            Left            =   150
            TabIndex        =   14
            Top             =   1590
            Width           =   1305
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   315
            Left            =   5070
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtSoftwareCode 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1590
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   3285
         End
         Begin VB.TextBox txtUnlockKey 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1590
            TabIndex        =   7
            Top             =   600
            Width           =   3285
         End
         Begin VB.CommandButton cmdRegister 
            Caption         =   "Unlock!"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5070
            TabIndex        =   6
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label lblWarning 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Request to send me an Email with the unlock information and please do not alter the subject."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   435
            Index           =   0
            Left            =   1530
            TabIndex        =   15
            Top             =   1590
            Width           =   5055
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblWarning 
            BackStyle       =   0  'Transparent
            Caption         =   $"Unlock.frx":08CA
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   1
            Left            =   90
            TabIndex        =   12
            Top             =   930
            Width           =   6585
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Software Serial No:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Unlock key:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1485
         End
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email: ircdominator@btopenworld.com"
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
         Left            =   2250
         MouseIcon       =   "Unlock.frx":097D
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2040
         Width           =   3075
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enigma Ware"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2250
         TabIndex        =   4
         Top             =   525
         Width           =   2025
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2490
         TabIndex        =   3
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5070
         TabIndex        =   2
         Top             =   1380
         Width           =   810
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed By Åß§ølµ†€•G€ñïµ§ 2001"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   1740
         Width           =   2895
      End
      Begin VB.Image imgLogo 
         Height          =   1395
         Left            =   330
         Picture         =   "Unlock.frx":0DBF
         Stretch         =   -1  'True
         Top             =   450
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sPassword As String


Private Sub cmdCancel_Click()
      bActivated = False
      Me.Hide
End Sub

Private Sub cmdRegister_Click()
      If txtUnlockKey <> "" Then
         If CheckUnlocked(sPassword, txtUnlockKey) Then
            Call WriteUnlockKey(txtUnlockKey)
         
            ' MDIMain.AL1.LiberationKey = txtUnlockKey
            ' If MDIMain.AL1.RegisteredUser Then
            Me.Tag = "UNLOCKED"
            Me.Hide
         Else
            MsgBox "Wrong key, try again!"
            txtUnlockKey.SelStart = 0
            txtUnlockKey.SelLength = Len(txtUnlockKey)
            txtUnlockKey.SetFocus
         End If
      End If
End Sub
'
' Private Sub cmdUnRegister_Click()
' Dim R As VbMsgBoxResult
' R = MsgBox("Are you sure that you want to unregister this software?", vbYesNo)
' If R = vbYes Then
' ActiveLock1.LiberationKey = "0"
' Unload Me
' End If
' End Sub

Private Sub cmdRequest_Click()
      ' SendUnlock
End Sub



Private Sub Form_Load()
      ' If the user hasn't registered yet,
      ' shows the registration frame
        
      Me.lblProductName = App.ProductName
      Me.lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        
      If IsUnlocked(sPassword) = True Then
         ' If MDIMain.AL1.RegisteredUser Then
         Me.Tag = "UNLOCKED"
         Timer1.Enabled = True
      Else
         Me.txtSoftwareCode = CreateSerialNo(sPassword)
         ' MDIMain.AL1.SoftwareCode
         Me.txtUnlockKey = ""
      End If
      Me.txtSubject = "IRCDominator!" & App.Major & "." & App.Minor & "." & App.Revision & "!" & Me.txtSoftwareCode
End Sub

'Private Sub lblEmail_Click()
'      SendUnlock
'End Sub
' Private Sub SendUnlock()
'
' Load frmEmailAdd
' frmEmailAdd.txtSubject = Me.txtSubject
' frmEmailAdd.sMailtoString = "mailto:ircdominator@btopenworld.com?subject=IRCDominator!" & App.Major & "." & App.Minor & "." & App.Revision & "!" & Me.txtSoftwareCode
' frmEmailAdd.txtTO = "ircdominator@btopenworld.com"
' Me.Hide
' frmEmailAdd.Show 1, Me
' Unload frmEmailAdd
' Set frmEmailAdd = Nothing
' Me.Show 1
' End Sub

Private Sub Timer1_Timer()
      Me.Hide
End Sub

Private Sub txtUnlockKey_Change()
      If txtUnlockKey <> "" Then
         Me.cmdRegister.Enabled = True
      Else
         Me.cmdRegister.Enabled = False
      End If
End Sub
