VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdate23 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update - Step 3"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmUpdate23.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5835
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5430
      Top             =   2100
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   435
      Left            =   1650
      TabIndex        =   16
      Top             =   3960
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   767
      _Version        =   393216
      BackColor       =   16761024
      FullWidth       =   263
      FullHeight      =   29
   End
   Begin VB.CommandButton cmdFin 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   4590
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   1650
      ScaleHeight     =   195
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   3390
      Width           =   3975
      Begin VB.Label Connectionstatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting..."
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3975
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6690
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin AutoUpdater.ProgYbar UpdateStatus 
      Height          =   285
      Left            =   1650
      TabIndex        =   17
      Top             =   3660
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   503
      ForeColor       =   16744576
      BackColor       =   12632256
      Max             =   100
      Mode            =   0
      Border          =   1
      Mark            =   0   'False
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H008080FF&
      BorderWidth     =   16
      Height          =   945
      Index           =   1
      Left            =   180
      Top             =   210
      Width           =   975
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   540
      TabIndex        =   13
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "IRCDominator"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   1350
      TabIndex        =   12
      Top             =   480
      Width           =   4275
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   4470
      TabIndex        =   11
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enigma Ware's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   480
      Index           =   9
      Left            =   1290
      TabIndex        =   10
      Top             =   120
      Width           =   2865
   End
   Begin VB.Image imgIcon 
      Height          =   630
      Left            =   330
      Picture         =   "frmUpdate23.frx":0ECA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   660
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "The MSN Chat Room Manager."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   1020
      Width           =   2265
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   3660
      TabIndex        =   8
      Top             =   60
      Width           =   2235
   End
   Begin VB.Image imgLogo 
      Height          =   3300
      Left            =   120
      Picture         =   "frmUpdate23.frx":1794
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Live Update"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1710
      TabIndex        =   7
      Top             =   1710
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1620
      X2              =   5610
      Y1              =   4470
      Y2              =   4470
   End
   Begin VB.Label Welcometext 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Continue..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1710
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1620
      X2              =   5610
      Y1              =   4485
      Y2              =   4485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please click the ""Update"" button to download the latest AutoUpdate."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1710
      TabIndex        =   6
      Top             =   2790
      Width           =   3945
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   780
      Index           =   3
      Left            =   1050
      TabIndex        =   14
      Top             =   840
      Width           =   4065
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FF8080&
      Height          =   225
      Index           =   4
      Left            =   1050
      TabIndex        =   15
      Top             =   540
      Width           =   4065
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   0
      Left            =   4830
      Top             =   1200
      Width           =   555
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   2
      Left            =   4830
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "frmUpdate23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataByte() As Byte
Dim i As Integer                'Code for downloading the file
Public Terr As Boolean
Public Dw_Url As String
Dim TFile As Long
Dim sSize As String
Dim bGettingFile As Boolean
Dim bDone As Boolean
Dim bCanceled As Boolean
Dim prgValue As Double

Private Sub Cancel_Click()
      CancelMe
End Sub

Sub CancelMe()
Dim GetFile

      On Error Resume Next
      If Inet1.StillExecuting Then
         GetFile = MsgBox("Are you sure you want to cancel the update...", vbYesNo)
         If GetFile = vbNo Then
            Exit Sub
         Else
            ' Cancels the download and cancels the update
            Inet1.Cancel
            bCanceled = True
            bDone = True
            DoEvents
            Kill App.Path & "\IRCDominator.ex_"
            Terr = True
'            MDIMain.Show
            Me.Hide
            Timer1.Enabled = True
         End If
      Else
'         MDIMain.Show
         Me.Hide
         Timer1.Enabled = True
      End If

End Sub
Private Sub cmdupdate_Click()
Dim sPath As String

      cmdupdate.Enabled = False
      Animation1.AutoPlay = True

      ClearAnim Animation1
      If Not (InIDE()) Then
         ' load and play avi
         LoadResAVI Animation1, "101"
      End If

      sPath = fGetIni("AutoUpdate", "WebHost", "http://homepage.ntlworld.com/mrenigma") & "/Update"

      Label2.Caption = "Please wait... while AutoUpdate is downloaded... "
      On Error Resume Next

      TFile = FreeFile
      Connectionstatus.Caption = "Please wait downloading update..."
      
      Kill App.Path & "\AutoUpdate.ex_"
      
      Open App.Path & "\AutoUpdate.ex_" For Binary As #TFile
      bGettingFile = False
      sSize = Inet1.OpenURL(sPath & "/Update/updatersize.txt", icString)
'      Stop
'      sSize = "389120"
      DoEvents
      sSize = Replace(Replace(Mid$(sSize, InStr(sSize, "=") + 1, Len(sSize)), vbCrLf, ""), ".", "")
      UpdateStatus.Max = sSize
      UpdateStatus.DrawBar 0
      bGettingFile = True
      Inet1.Execute sPath & "/AutoUpdate.exe", "GET " & sPath & "/AutoUpdate.exe " & App.Path & "\AutoUpdate.ex_"
End Sub

Private Sub Form_Load()
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
StayOnTop Me.hwnd, True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         CancelMe
      End If
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim vtData As Variant
Dim strData As String

      bDone = False
      On Error Resume Next
      Select Case State
         Case 1
            Connectionstatus.Caption = "Resolving Host..."
         Case 3
            Connectionstatus.Caption = "Connecting..."
         Case icConnected
            Connectionstatus.Caption = "Connected..."
         Case icReceivingResponse
            Connectionstatus.Caption = "Receiving..."
         Case icDisconnected
            Connectionstatus.Caption = "Disconnected..."
         Case icDisconnecting
            Connectionstatus.Caption = "Disconnecting..."
         Case icResponseCompleted  ' 12
            If bGettingFile = True Then
               ' Get first chunk.
               DataByte() = Inet1.GetChunk(1024, icByteArray)
               Put #TFile, , DataByte()
               If UBound(DataByte()) >= 0 And Not (bDone) Then
                  prgValue = prgValue + UBound(DataByte)
                  UpdateStatus.DrawBar prgValue
               End If
               DoEvents
               Do While Not bDone
                  If bCanceled Then
                     Inet1.Cancel
                     Close #TFile
                     Me.Hide
                     Timer1.Enabled = True
                     Exit Sub
                  End If
                  ' Get next chunk.
                  DataByte() = Inet1.GetChunk(1024, icByteArray)
                  If UBound(DataByte()) >= 0 And Not (bDone) Then
                     prgValue = prgValue + UBound(DataByte)
                     UpdateStatus.DrawBar prgValue
                  End If
                  Put #TFile, , DataByte()
                  DoEvents
                  If UBound(DataByte()) <= 0 Then
                     bDone = True
                     Me.Animation1.AutoPlay = False
                  End If
               Loop
            End If
      End Select
      If bDone And bGettingFile Then
         If Inet1.StillExecuting = False Then
            Connectionstatus.Caption = "Update Downloaded"
            cmdFin.Enabled = True
            Cancel.Enabled = False
            cmdupdate.Enabled = False
            Inet1.Cancel
            Close #TFile
         End If
         Welcometext.Caption = "Complete..."
         UpdateStatus.DrawBar 100
         Connectionstatus.Caption = "Download Complete"
         cmdFin.Enabled = True
         Cancel.Enabled = False
         cmdupdate.Enabled = False
      End If
End Sub
Private Sub cmdFin_Click()
Dim FileNumber As Integer
Dim filebuffer() As Byte

      On Error Resume Next
      If Dir(App.Path & "\Update.exe") <> "" Then
         Kill (App.Path & "\Update.exe")
      End If

      filebuffer = LoadResData(101, "CUSTOM")
      FileNumber = FreeFile
      Open App.Path & "\Update.exe" For Binary Access Write As #FileNumber
      Put #FileNumber, , filebuffer
      Close #FileNumber
      Inet1.Cancel
      MsgBox "This program will now restart...", vbInformation, "EnigmaWare Live Update"
      Shell App.Path & "\Update.exe"
      End
End Sub

Private Sub Timer1_Timer()
      Timer1.Enabled = False
      Unload Me
End Sub
