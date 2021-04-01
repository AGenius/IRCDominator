VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTrace 
   Caption         =   "Server Trace"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   Icon            =   "frmTrace.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   9375
   Begin VB.TextBox txtTrace 
      Height          =   1545
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   9285
   End
   Begin RichTextLib.RichTextBox txtTraceold 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1085
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTrace.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Me.WindowState = vbMinimized
    End If

End Sub

Private Sub Form_Resize()
        On Error Resume Next
        Me.txtTrace.Top = 0
        Me.txtTrace.Left = 0
        Me.txtTrace.Width = Me.ScaleWidth
        Me.txtTrace.Height = Me.ScaleHeight
End Sub
