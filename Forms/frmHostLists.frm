VERSION 5.00
Begin VB.Form frmHostLists 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Host / Owner Preferences"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmHostLists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8280
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
      Left            =   1560
      TabIndex        =   16
      Top             =   5580
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
      Left            =   30
      TabIndex        =   15
      Top             =   5580
      Width           =   1365
   End
   Begin VB.CheckBox chkListActive 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto Host List"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   13
      Tag             =   "ActiveLists|Hosts"
      Top             =   -30
      Width           =   1485
   End
   Begin VB.CheckBox chkListActive 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto Owner List"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Tag             =   "ActiveLists|Owners"
      Top             =   -30
      Width           =   1665
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00C0C000&
      Height          =   5445
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3765
      Begin VB.TextBox txtNickName 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   3930
         Width           =   3568
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   705
         Index           =   0
         Left            =   2760
         Picture         =   "frmHostLists.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "STR|Clear the users from the list"
         Top             =   4380
         Width           =   885
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   705
         Index           =   0
         Left            =   1470
         Picture         =   "frmHostLists.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "STR|Remove a user from the list"
         Top             =   4380
         Width           =   885
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   705
         Index           =   0
         Left            =   120
         Picture         =   "frmHostLists.frx":205E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "STR|Add a user to the list"
         Top             =   4380
         Width           =   885
      End
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
         Height          =   3630
         Index           =   0
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "Owners.dat"
         Top             =   240
         Width           =   3568
      End
      Begin VB.Label lblListCount 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   5130
         Width           =   3555
      End
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00C0C000&
      Height          =   5445
      Index           =   1
      Left            =   4380
      TabIndex        =   0
      Top             =   0
      Width           =   3765
      Begin VB.TextBox txtNickName 
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   3930
         Width           =   3568
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   705
         Index           =   1
         Left            =   2760
         Picture         =   "frmHostLists.frx":2928
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "STR|Clear the users from the list"
         Top             =   4380
         Width           =   885
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   705
         Index           =   1
         Left            =   1470
         Picture         =   "frmHostLists.frx":31F2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "STR|Remove a user from the list"
         Top             =   4380
         Width           =   885
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   705
         Index           =   1
         Left            =   120
         Picture         =   "frmHostLists.frx":3ABC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "STR|Add a user to the list"
         Top             =   4380
         Width           =   885
      End
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
         Height          =   3630
         Index           =   1
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   1
         Tag             =   "Hosts.dat"
         Top             =   240
         Width           =   3568
      End
      Begin VB.Label lblListCount 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   5130
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmHostLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objTooltip    As cTooltip

Private Sub cmdAdd_Click(Index As Integer)
      If FindInList(lstNickNames(Index), Trim(txtNickName(Index))) > -1 Then
         Call MsgBox("Already in List", vbOKOnly, "Add User")
      Else
         lstNickNames(Index).AddItem Trim(txtNickName(Index))
         txtNickName(Index) = ""
      End If
      lblListCount(Index).Caption = lstNickNames(Index).ListCount & " Names in list"

End Sub
Public Sub chkListActive_Click(Index As Integer)
      On Error Resume Next
      If chkListActive(Index).Value Then
         fraList(Index).Enabled = True
         cmdClear(Index).Enabled = True
         txtNickName(Index).Enabled = True
         lstNickNames(Index).Enabled = True
      Else
         fraList(Index).Enabled = False
         cmdAdd(Index).Enabled = False
         cmdRemove(Index).Enabled = False
         cmdClear(Index).Enabled = False
         txtNickName(Index).Enabled = False
         lstNickNames(Index).Enabled = False
      End If
End Sub
Private Sub cmdCancel_Click()
      Me.Hide
End Sub
Private Sub cmdClear_Click(Index As Integer)
      If lstNickNames(Index).ListCount > 0 Then
         If MsgBox("Are You Sure", vbYesNo, "Clear NickNames List") = vbYes Then
            lstNickNames(Index).Clear
            txtNickName(Index).Text = ""
         End If
      End If
      lblListCount(Index).Caption = lstNickNames(Index).ListCount & " Names in list"

End Sub
Private Sub cmdOk_Click()
      HostLists.List_Owners_Active = CBool(Me.chkListActive(0))
      HostLists.List_Hosts_Active = CBool(Me.chkListActive(1))
      HostLists.SaveControl lstNickNames(0), lOwnerList
      HostLists.SaveControl lstNickNames(1), lHostList
      HostLists.SavePrefs
      Me.Hide
End Sub
Private Sub cmdRemove_Click(Index As Integer)
      If lstNickNames(Index).ListIndex > -1 Then
         lstNickNames(Index).RemoveItem lstNickNames(Index).ListIndex
         txtNickName(Index).Text = ""
         cmdAdd(Index).Enabled = False
         cmdRemove(Index).Enabled = False
      End If
      lblListCount(Index).Caption = lstNickNames(Index).ListCount & " Names in list"

End Sub
Private Sub Form_Load()
Dim i As Integer

      LoadToolTips Me, m_objTooltip
      For i = 0 To cmdAdd.UBound
         cmdAdd(i).Enabled = False
         cmdRemove(i).Enabled = False
      Next
      Call HostLists.FillControl(lstNickNames(0), lOwnerList)
      Call HostLists.FillControl(lstNickNames(1), lHostList)

      lblListCount(0).Caption = lstNickNames(0).ListCount & " Names in list"
      lblListCount(1).Caption = lstNickNames(1).ListCount & " Names in list"
      Me.chkListActive(0).Value = Abs(HostLists.List_Owners_Active)
      Me.chkListActive(1).Value = Abs(HostLists.List_Hosts_Active)
      Me.chkListActive_Click (0)
      Me.chkListActive_Click (1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
End Sub

Private Sub lstNickNames_Click(Index As Integer)
      If lstNickNames(Index).ListIndex <> -1 Then
         txtNickName(Index) = lstNickNames(Index).List(lstNickNames(Index).ListIndex)
         cmdRemove(Index).Enabled = True
         cmdAdd(Index).Enabled = False
      End If
End Sub
Private Sub txtNickName_Change(Index As Integer)
      If txtNickName(Index) <> "" Then
         cmdAdd(Index).Enabled = True
         cmdRemove(Index).Enabled = False
      Else
         cmdAdd(Index).Enabled = False
         cmdRemove(Index).Enabled = False
      End If
End Sub
