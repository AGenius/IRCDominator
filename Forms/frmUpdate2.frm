VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmUpdate2 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update - Step 2"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmUpdate2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5835
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   4560
      Left            =   6345
      TabIndex        =   19
      Top             =   315
      Width           =   6450
      ExtentX         =   11377
      ExtentY         =   8043
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
   Begin AutoUpdater.ProgYbar UpdateStatus 
      Height          =   285
      Left            =   1650
      TabIndex        =   18
      Top             =   3690
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5340
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1650
      ScaleHeight     =   195
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   3390
      Width           =   3975
      Begin VB.Label Connectionstatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connected"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton Back 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2940
      TabIndex        =   5
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton NextButton 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton CloseUpdate 
      Caption         =   "&End"
      Height          =   375
      Left            =   4260
      TabIndex        =   9
      Top             =   4590
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label UpdateYes 
      BackStyle       =   0  'Transparent
      Caption         =   "There is a update available, please click next to..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   4050
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Image imgIcon 
      Height          =   630
      Left            =   330
      Picture         =   "frmUpdate2.frx":0ECA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   660
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      Index           =   0
      Left            =   1290
      TabIndex        =   12
      Top             =   120
      Width           =   2865
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
      TabIndex        =   11
      Top             =   1020
      Width           =   2265
   End
   Begin VB.Image imgLogo 
      Height          =   3300
      Left            =   120
      Picture         =   "frmUpdate2.frx":1794
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Welcometext 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
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
      Left            =   1740
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while I connect to the internet to check for an AutoUpdater update..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1770
      TabIndex        =   3
      Top             =   2850
      Width           =   3615
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
      Index           =   1
      Left            =   1710
      TabIndex        =   2
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1620
      X2              =   5610
      Y1              =   4485
      Y2              =   4485
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
   Begin VB.Label lblInfo 
      BackColor       =   &H00FF8080&
      Height          =   225
      Index           =   4
      Left            =   1050
      TabIndex        =   17
      Top             =   540
      Width           =   4065
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   780
      Index           =   3
      Left            =   1050
      TabIndex        =   16
      Top             =   840
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
      TabIndex        =   10
      Top             =   60
      Width           =   2235
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
Attribute VB_Name = "frmUpdate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RASCONN
      dwSize As Long
      hRasConn As Long
      szEntryName(256) As Byte          ' CODE TO CHECK IF UN ARE CONNECTED TO THE INTERNET
      szDeviceType(16) As Byte
      szDeviceName(128) As Byte
End Type

Private Declare Function RasEnumConnectionsA& Lib "RasApi32.DLL" (lprasconn As Any, lpcb&, lpcConnections&)

Private Sub Back_Click()
      frmUpdate.Visible = True        'back button if u want to go back a step it will close this from
      frmUpdate2.Visible = False       'hide this form and make the first appear again
      UpdateStatus.DrawBar 0
      ' UpdateStatus.Value = "0"      'sets the progressbar value 2 "0" so u dont get errors
      Connectionstatus.Caption = "Disconnected" 'sets the caption text to Disconnected
End Sub

Private Sub Cancel_Click()
      frmUpdate.Visible = False 'if u click cancel resets everything and go's back to main form
      frmUpdate2.Visible = False
      UpdateStatus.DrawBar 0
      Unload frmUpdate
      Unload Me
End Sub

Private Sub Command1_Click()
      frmUpdate.Visible = True       'proceeds 2 next step
      frmUpdate2.Visible = False
End Sub

Private Sub CloseUpdate_Click()
      frmUpdate.Visible = False
      frmUpdate2.Visible = False
      frmUpdate3.Visible = False
      UpdateStatus.DrawBar 0
      ' UpdateStatus.Value = "0"
      Connectionstatus.Caption = "Disconnected"
      ' MDIMain.Visible = True
End Sub


Private Sub Form_Load()
      Picture2.BackColor = &HC0C0C0
      Me.Top = Screen.Height / 2 - Me.Height / 2
      Me.Left = Screen.Width / 2 - Me.Width / 2
      StayOnTop Me.hwnd, True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         frmUpdate.Visible = False 'if u click cancel resets everything and go's back to main form
         frmUpdate2.Visible = False
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
      Unload Me

End Sub

Private Sub NextButton_Click()
      UpdateStatus.DrawBar 0
      frmUpdate3.Visible = True      'Proceeds to next step
      frmUpdate2.Visible = False
      frmUpdate3.cmdupdate.Enabled = True
      Unload frmUpdate
      Unload frmUpdate2
      Unload Me
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim vtData As Variant
Dim strData As String

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
      End Select
End Sub

Public Sub Check()
Dim WebHost As String
Dim LatestVersion As String
Dim lCurVer As Long
Dim sVersion As String
Dim sStr As String
Dim bTest As Byte

      ' You can put in your web host here
      ' Geocities, Yahoo, Angelfire...all work :)
      WebHost = fGetIni("AutoUpdate", "WebHost", "http://homepage.ntlworld.com/mrenigma")
      sStr = WebHost & "/Update/newversion.txt"
      WB.Navigate sStr
      
      LatestVersion = Inet1.OpenURL(sStr, icString)
      If Locate(LatestVersion, "File Not Found") Then
         Connectionstatus.Caption = "Update Information not Found"
         Exit Sub
      End If
      If Left$(LatestVersion, 7) <> "VERSION" Then
         Exit Sub
      End If
      sVersion = Mid$(LatestVersion, Locate(LatestVersion, "="), 999)
      LatestVersion = Replace(Replace(Mid$(LatestVersion, InStr(1, LatestVersion, "=") + 1, Len(LatestVersion)), vbCrLf, ""), ".", "")

      IRCIniFile.Section = "General"
      IRCIniFile.Key = "AppVersion"
      lCurVer = Replace(IRCIniFile.Value, ".", "")
      

      ' Trim off the CrLf if it exists
      If Right$(LatestVersion, 2) = vbCrLf Then LatestVersion = Left$(LatestVersion, Len(LatestVersion) - 2)

      If CLng(LatestVersion) > lCurVer Then
         ' Notify the user there's a newer version available
         NextButton.Enabled = True
         UpdateYes.Caption = "There is a update available, please click next..."
         UpdateYes.Visible = True       'then when it gets to higher than 97 it stps itself so u dont get an error
         UpdateStatus.DrawBar 99
         Connectionstatus.Caption = "Connected: Found New Version :" & sVersion
      Else
         ' Notify the user that they're using the most current version
         NextButton.Enabled = False
         UpdateYes.Caption = "There are no updates available"
         Connectionstatus.Caption = "No New Update"
         UpdateYes.Visible = True       'then when it gets to higher than 97 it stps itself so u dont get an error
         UpdateStatus.DrawBar 99
         ' UpdateStatus.Value = 99
      End If
 

End Sub
