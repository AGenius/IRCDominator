VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmChat 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WB 
      CausesValidation=   0   'False
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7515
      ExtentX         =   13256
      ExtentY         =   9022
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
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'        frmMain.SERVER1.Close
'        frmMain.SERVER2.Close
'End Sub

Private Sub Form_Resize()
        wb.Top = 0
        wb.Left = 0
        wb.Width = Me.ScaleWidth
        wb.Height = Me.ScaleHeight
End Sub
