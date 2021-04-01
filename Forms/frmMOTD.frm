VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMOTD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Of The Day"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9480
   Icon            =   "frmMOTD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Top             =   4920
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      Picture         =   "frmMOTD.frx":08CA
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   90
      Width           =   465
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   4200
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   9270
      ExtentX         =   16351
      ExtentY         =   7408
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Message Of the day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   660
      TabIndex        =   3
      Top             =   180
      Width           =   2790
   End
End
Attribute VB_Name = "frmMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Check()
Dim WebHost As String
Dim sStr As String

      WebHost = fGetIni("AutoUpdate", "WebHost", "http://homepage.ntlworld.com/mrenigma")
      sStr = WebHost & "/MOTD/MOTD.HTML"
      WB.Navigate sStr
End Sub


Private Sub cmdClose_Click()
      Me.Hide
      Unload Me
End Sub

Private Sub Form_Load()
      Check
End Sub

