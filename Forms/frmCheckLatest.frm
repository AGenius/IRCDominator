VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCheckLatest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Latest Version Checker"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckLatest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5010
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   465
      Left            =   5910
      TabIndex        =   2
      Top             =   390
      Width           =   1035
      ExtentX         =   1826
      ExtentY         =   820
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
      Location        =   ""
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1350
      TabIndex        =   1
      Top             =   3930
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Height          =   2355
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   4305
   End
End
Attribute VB_Name = "frmCheckLatest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bHidePopup As Boolean

' ***IMPORTANT***
' Create a TXT file in Notepad with the following line:
'
' VERSION=2.0 [or current version]
'
' ^-- blank line
' Name this file "newversion.txt" and post it on your web
' server.  Then, edit the lines below (in the subs:
' Winsock_Connect and cmdCheck_Click) with that URL.
' ***IMPORTANT***


Sub Pause(duration)
      ' This will pause for the duration [duration is in seconds]
Dim Current As Long

      Current = Timer
      Do Until Timer - Current >= duration
         DoEvents
      Loop
End Sub

Public Sub Check()
Dim WebHost As String
Dim LatestVersion As String
Dim lCurVer As Long

      ' You can put in your web host here
      ' Geocities, Yahoo, Angelfire...all work :)
      On Error Resume Next
      WebHost = fGetIni("AutoUpdate", "WebHost", "http://homepage.ntlworld.com/mrenigma")
      LatestVersion = Inet1.OpenURL(WebHost & "/Update/newversion.txt", icString)
      If Left$(LatestVersion, 7) <> "VERSION" Then
         Exit Sub
      End If

      ' Define the latest version
      LatestVersion = Replace(Replace(Mid$(LatestVersion, InStr(LatestVersion, "=") + 1, Len(LatestVersion)), vbCrLf, ""), ".", "")
      lCurVer = App.Major & App.Minor & App.Revision
      ' Stop
      ' Trim off the CrLf if it exists
      If Right$(LatestVersion, 2) = vbCrLf Then LatestVersion = Left$(LatestVersion, Len(LatestVersion) - 2)

      If CLng(LatestVersion) > lCurVer Then
         ' Notify the user there's a newer version available
         If bHidePopup = False Then
            Load frmPopup
            frmPopup.lblMessage = "There is a newer version of " & App.ProductName & "  available!  Click here to Launch AutoUpdate!"
            frmPopup.iStayUp = 5000
            frmPopup.Show
            frmPopup.sDo = "CHECK"
            frmPopup.Timer1.Enabled = True
         End If
         Me.Tag = "NEW"
      Else
         ' Notify the user that they're using the most current version
         Me.Tag = "OLD"
      End If
End Sub


Private Sub cmdOk_Click()
      Me.Hide
End Sub

