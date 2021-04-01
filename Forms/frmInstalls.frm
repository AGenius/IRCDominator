VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInstalls 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2865
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmInstalls.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   2805
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin MSComctlLib.ProgressBar pgProgress 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   2430
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while the required sound files and registry settings are installed to your computer This will only happen once."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   720
         Index           =   0
         Left            =   1350
         TabIndex        =   4
         Top             =   1080
         Width           =   5340
      End
      Begin VB.Image imgLogo 
         Height          =   2490
         Left            =   60
         Picture         =   "frmInstalls.frx":000C
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Installing Registry Settings Please wait ......"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   2130
         Width           =   4890
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
         Left            =   1380
         TabIndex        =   2
         Top             =   210
         Width           =   2895
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
         Left            =   1590
         TabIndex        =   1
         Top             =   540
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmInstalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      Me.lblProductName = App.ProductName
      DoEvents
End Sub
Private Sub Form_Paint()
      DoEvents
      InstallSoundsNow
      Me.Hide
End Sub
