VERSION 5.00
Begin VB.Form frmTraceOptions 
   BackColor       =   &H00C0C000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server Trace Preferences"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkMODE 
      BackColor       =   &H00C0C000&
      Caption         =   "Show MODE Changes"
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
      Left            =   60
      TabIndex        =   9
      Tag             =   "TRACE|MODES"
      Top             =   2820
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CheckBox chkAuth 
      BackColor       =   &H00C0C000&
      Caption         =   "Show AUTH info"
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
      Left            =   60
      TabIndex        =   8
      Tag             =   "TRACE|AUTH"
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   1350
      TabIndex        =   7
      Top             =   3330
      Width           =   1245
   End
   Begin VB.CheckBox chkALL 
      BackColor       =   &H00C0C000&
      Caption         =   "Select ALL"
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
      Left            =   60
      TabIndex        =   6
      Top             =   450
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CheckBox chkPRIVMSG 
      BackColor       =   &H00C0C000&
      Caption         =   "Show PRIVMSG's"
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
      Left            =   60
      TabIndex        =   5
      Tag             =   "TRACE|PRIVMSG"
      Top             =   2190
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CheckBox chkJoins 
      BackColor       =   &H00C0C000&
      Caption         =   "Show Joins"
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
      Left            =   60
      TabIndex        =   4
      Tag             =   "TRACE|JOINS"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CheckBox chkParts 
      BackColor       =   &H00C0C000&
      Caption         =   "Show Parts"
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
      Left            =   60
      TabIndex        =   3
      Tag             =   "TRACE|PARTS"
      Top             =   1860
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CheckBox chkAccess 
      BackColor       =   &H00C0C000&
      Caption         =   "Show Access List (When Viewing)"
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
      Left            =   60
      TabIndex        =   2
      Tag             =   "TRACE|ACCESS"
      Top             =   1260
      Value           =   1  'Checked
      Width           =   3105
   End
   Begin VB.CheckBox chkWhispers 
      BackColor       =   &H00C0C000&
      Caption         =   "Show Whispers"
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
      Left            =   60
      TabIndex        =   0
      Tag             =   "TRACE|WHISPERS"
      Top             =   960
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Select Which Items to show in trace window"
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
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3630
   End
End
Attribute VB_Name = "frmTraceOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkALL_Click()
      Me.chkAccess.Value = chkALL.Value
      Me.chkJoins.Value = chkALL.Value
      Me.chkParts.Value = chkALL.Value
      Me.chkPRIVMSG.Value = chkALL.Value
      Me.chkWhispers.Value = chkALL.Value
      Me.chkAuth.Value = chkALL.Value
      Me.chkMODE.Value = chkALL.Value
End Sub

Private Sub cmdOK_Click()
      Save_Settings Me
      Me.Hide
End Sub

Private Sub Form_Load()
      Load_Settings Me
End Sub
