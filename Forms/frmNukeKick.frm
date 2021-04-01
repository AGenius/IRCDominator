VERSION 5.00
Begin VB.Form frmNukeKick 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuke Kicking Options"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmNukeKick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   16
      Top             =   6330
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   15
      Top             =   6330
      Width           =   1365
   End
   Begin VB.Frame fraFrame 
      BackColor       =   &H00C0C000&
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox chkDisable 
         BackColor       =   &H00C0C000&
         Caption         =   "Disable Users IRCDominator"
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Tag             =   "ActiveLists|DomDisable"
         Top             =   2970
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.OptionButton optBan 
         BackColor       =   &H00C0C000&
         Caption         =   "Ban For"
         Height          =   225
         Left            =   1410
         TabIndex        =   17
         Tag             =   "ActiveLists|NukeBan"
         Top             =   525
         Width           =   885
      End
      Begin VB.Frame fraList 
         BackColor       =   &H00C0C000&
         Height          =   5745
         Left            =   4710
         TabIndex        =   7
         Top             =   210
         Width           =   3075
         Begin VB.ListBox lstNickNames 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2610
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   13
            Tag             =   "NukeKicks"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   705
            Left            =   120
            Picture         =   "frmNukeKick.frx":08D2
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3360
            Width           =   885
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   705
            Left            =   1110
            Picture         =   "frmNukeKick.frx":119C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3360
            Width           =   885
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   705
            Left            =   2100
            Picture         =   "frmNukeKick.frx":1A66
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3360
            Width           =   885
         End
         Begin VB.TextBox txtNickName 
            Height          =   345
            Left            =   90
            TabIndex        =   9
            Top             =   2910
            Width           =   2895
         End
         Begin VB.TextBox txtKickingMessage 
            Height          =   1200
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   8
            Tag             =   "ActiveLists|KicksMessage"
            Text            =   "frmNukeKick.frx":2330
            Top             =   4410
            Width           =   2865
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Kick Message"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   4140
            Width           =   1080
         End
      End
      Begin VB.TextBox txtNukeMessage 
         Height          =   1500
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "ActiveLists|NukeMessage"
         Text            =   "frmNukeKick.frx":233C
         Top             =   1050
         Width           =   4275
      End
      Begin VB.OptionButton optNoBan 
         BackColor       =   &H00C0C000&
         Caption         =   "No Ban"
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Tag             =   "ActiveLists|NukeNoBan"
         Top             =   525
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CheckBox chkKickNuke 
         BackColor       =   &H00C0C000&
         Caption         =   "Auto Kick for Nuking"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Tag             =   "ActiveLists|KickNuking"
         Top             =   0
         Width           =   2070
      End
      Begin VB.ComboBox cboBans 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNukeKick.frx":235C
         Left            =   2400
         List            =   "frmNukeKick.frx":2371
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "ActiveLists|NukeBanTime"
         Top             =   480
         Width           =   2145
      End
      Begin VB.CheckBox chkAddKickList 
         BackColor       =   &H00C0C000&
         Caption         =   "Add To Auto Kick List"
         Height          =   225
         Left            =   240
         TabIndex        =   1
         Tag             =   "ActiveLists|NukeAddAutoKick"
         Top             =   2640
         Width           =   2625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Kick Message"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmNukeKick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub chkKickNuke_Click()
      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If chkKickNuke.Value = 1 Then
'         cmdAdd.Enabled = True
         chkAddKickList.Enabled = True
         cboBans.Enabled = True
         optBan.Enabled = True
         cmdClear.Enabled = True
         txtKickingMessage.Enabled = True
         txtNickName.Enabled = True
         txtNukeMessage.Enabled = True
         lstNickNames.Enabled = True
         optNoBan.Enabled = True
'         cmdRemove.Enabled = True
         cboBans.Enabled = True
         If Me.optNoBan.Value Then
            Me.optNoBan_Click
         Else
            Me.optBan_Click
         End If
      Else
         cmdAdd.Enabled = False
         chkAddKickList.Enabled = False
         cboBans.Enabled = False
         optBan.Enabled = False
         cmdClear.Enabled = False
         txtKickingMessage.Enabled = False
         txtNickName.Enabled = False
         txtNukeMessage.Enabled = False
         lstNickNames.Enabled = False
         optNoBan.Enabled = False
         cmdRemove.Enabled = False
         cboBans.Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub
Private Sub cmdAdd_Click()
      If FindInList(lstNickNames, txtNickName) > -1 Then
         Call MsgBox("Already in List", vbOKOnly, "Add User")
      Else
         lstNickNames.AddItem txtNickName
         txtNickName = ""
      End If
End Sub
Private Sub cmdCancel_Click()
      Me.Hide
End Sub
Private Sub cmdClear_Click()
      If lstNickNames.ListCount > 0 Then
         If MsgBox("Are You Sure", vbYesNo, "Clear NickNames List") = vbYes Then
            lstNickNames.Clear
            txtNickName.Text = ""
         End If
      End If
End Sub
Private Sub cmdOk_Click()
      Call SaveListSettings

      Call Save_Settings(Me)
      Me.Hide
End Sub
Public Sub SaveListSettings()

      Call SaveList(Me.lstNickNames, lstNickNames.Tag)

End Sub
Private Sub cmdRemove_Click()
      If lstNickNames.ListIndex > -1 Then
         lstNickNames.RemoveItem lstNickNames.ListIndex
         txtNickName.Text = ""
         cmdAdd.Enabled = False
         cmdRemove.Enabled = False
      End If
End Sub
Private Sub Form_Load()
      Call Load_Settings(Me)
      chkKickNuke_Click
      If bOwner Then chkDisable.Visible = True
      
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         Me.Hide
         Cancel = True
      End If
End Sub
Private Sub lstNickNames_Click()
      If lstNickNames.ListIndex <> -1 Then
         txtNickName = lstNickNames.List(lstNickNames.ListIndex)
         cmdRemove.Enabled = True
         cmdAdd.Enabled = False
      End If
End Sub
Public Sub optBan_Click()
      If Me.optBan.Value Then
         Me.cboBans.Enabled = True
      End If
End Sub
Public Sub optNoBan_Click()
      If Me.optNoBan.Value Then
         Me.cboBans.Enabled = False
      End If
End Sub
Private Sub txtNickName_Change()
      If txtNickName <> "" Then
         cmdAdd().Enabled = True
         cmdRemove().Enabled = False
      Else
         cmdAdd().Enabled = False
         cmdRemove().Enabled = False
      End If
End Sub
