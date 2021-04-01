VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLists 
   BackColor       =   &H00808000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Room Access List"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   6900
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import List"
      Height          =   375
      Left            =   4770
      TabIndex        =   4
      Top             =   4830
      Width           =   1335
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record List"
      Height          =   375
      Left            =   3330
      TabIndex        =   3
      Top             =   4830
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected"
      Height          =   375
      Left            =   1890
      TabIndex        =   2
      Top             =   4830
      Width           =   1335
   End
   Begin VB.ListBox lstAcess 
      Height          =   4335
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   10455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   150
      TabIndex        =   0
      Top             =   4830
      Width           =   1365
   End
End
Attribute VB_Name = "frmLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
Dim i As Integer
Dim sTemp As String
Dim sKey As String
Dim sEntry As String
Dim iLoop As Long

    If MsgBox("Are you sure you wish to remove the selected entries from the access list", vbYesNo, "Remove from Access") = vbYes Then
    Me.Enabled = False
        For i = 0 To Me.lstAcess.ListCount - 1
            If Me.lstAcess.Selected(i) = True Then
                sTemp = lstAcess.List(i)
                sKey = Split(sTemp, " ")(0)
                sEntry = Split(sTemp, " ")(1)
                sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " DELETE " & sKey & " " & sEntry
                SendServer2 sTemp
                DoEvents
                iLoop = iLoop + 1
                If iLoop > 5 Then
                    For iLoop = 1 To 1000000
                        DoEvents
                    Next
                    iLoop = 0
                End If
            End If
        Next
    End If
    Me.Hide
End Sub
Private Sub cmdImport_Click()
Dim iFile As Integer
Dim sTemp As String
Dim iLoop As Long

    cmdlg.Filter = "*.dat"
    ' cmdlg.FileName = "*.txt"
    cmdlg.FilterIndex = 0
    cmdlg.ShowOpen
        
    If Dir(cmdlg.FileName) <> "" And cmdlg.FileName <> "" Then
        ' File Exists
        iFile = FreeFile
        Open cmdlg.FileName For Input Access Read As #iFile
Me.Enabled = False
        Do
            If EOF(iFile) Then
                Exit Do
            End If
            Line Input #iFile, sTemp
            sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " ADD " & Split(sTemp, " ")(0) & " " & Split(sTemp, " ")(1) & " " & (Split(sTemp, " ")(2)) & " :" & Split(sTemp, ":")(1)
            SendServer2 sTemp
            DoEvents
            iLoop = iLoop + 1
            If iLoop > 5 Then
                For iLoop = 1 To 1000000
                    DoEvents
                Next
                iLoop = 0
            End If
        Loop
        Close #iFile
    End If
    Me.Hide
End Sub
Private Sub cmdOk_Click()
    Me.Hide
End Sub
Private Sub cmdRecord_Click()
    Call SaveList(Me.lstAcess, "Accesslst.dat")
    MsgBox "Access List Saved"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lstAcess.Width = Me.Width - 500
    Me.lstAcess.Height = Me.Height - 1500
    Me.cmdOK.Top = Me.Height - 1000
    Me.cmdDelete.Top = cmdOK.Top
    Me.cmdRecord.Top = cmdOK.Top
    Me.cmdImport.Top = cmdOK.Top

End Sub
