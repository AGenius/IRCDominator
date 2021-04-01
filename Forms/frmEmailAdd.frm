VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEmailAdd 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mail unlock code"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmEmailAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   8205
      Left            =   7530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Text            =   "frmEmailAdd.frx":038A
      Top             =   90
      Width           =   4755
   End
   Begin VB.PictureBox picstatus 
      BackColor       =   &H00D6DFDE&
      Height          =   300
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   7320
      TabIndex        =   25
      Top             =   8460
      Width           =   7380
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status : Ide"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   30
         Width           =   810
      End
      Begin VB.Image imgstatus 
         Height          =   240
         Left            =   15
         Top             =   15
         Width           =   240
      End
   End
   Begin MSWinsockLib.Winsock smtp 
      Left            =   3630
      Top             =   7890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   1980
      TabIndex        =   6
      Top             =   7830
      Width           =   1425
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   405
      Left            =   270
      TabIndex        =   5
      Top             =   7830
      Width           =   1425
   End
   Begin VB.Frame frmEmail 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   -30
      TabIndex        =   11
      Top             =   2430
      Width           =   7485
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   2250
         TabIndex        =   0
         Top             =   0
         Width           =   4785
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   2250
         TabIndex        =   1
         Top             =   390
         Width           =   4785
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Height          =   4245
         Left            =   60
         TabIndex        =   12
         Top             =   1110
         Width           =   7305
         Begin VB.TextBox txtMessage 
            Height          =   2055
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   4
            Top             =   1950
            Width           =   6945
         End
         Begin IRCDominator.Bevel Bevel2 
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   1380
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   873
            BevelStyle      =   2
            BackColor       =   12640511
            Begin VB.PictureBox picbar 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Height          =   355
               Left            =   60
               ScaleHeight     =   360
               ScaleWidth      =   6780
               TabIndex        =   14
               Top             =   60
               Width           =   6780
               Begin VB.TextBox txtSubject 
                  Height          =   285
                  Left            =   840
                  Locked          =   -1  'True
                  TabIndex        =   15
                  Top             =   30
                  Width           =   5850
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "&Subject :"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Index           =   2
                  Left            =   105
                  TabIndex        =   16
                  Top             =   75
                  Width           =   630
               End
            End
         End
         Begin IRCDominator.Bevel Bevel1 
            Height          =   765
            Left            =   120
            TabIndex        =   17
            Top             =   420
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   1349
            BevelStyle      =   2
            BackColor       =   12640511
            Begin VB.TextBox txtTO 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   540
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   90
               Width           =   6225
            End
            Begin VB.TextBox txtFrom 
               Height          =   285
               Left            =   540
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   435
               Width           =   6225
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "From :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   90
               TabIndex        =   21
               Top             =   450
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "To :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   105
               TabIndex        =   20
               Top             =   120
               Width           =   270
            End
         End
         Begin VB.Label lblWarning 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "This is the contents of the E-mail - You may add a message if you wish"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   0
            TabIndex        =   22
            Top             =   150
            Width           =   7215
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Name or NickName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   60
         Width           =   2025
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your E-Mail Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   3
         Left            =   450
         TabIndex        =   23
         Top             =   450
         Width           =   1725
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   2
         X1              =   0
         X2              =   7380
         Y1              =   900
         Y2              =   900
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Please Select the E-mail Transport Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   945
      Left            =   270
      TabIndex        =   10
      Top             =   1350
      Width           =   6735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Send E-mails using your email client (eg Outlook express) if configured."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   330
         Width           =   6405
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Send E-mails direct (built in smtp sending routine)."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   630
         Value           =   -1  'True
         Width           =   4605
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   4425
      Picture         =   "frmEmailAdd.frx":0390
      Top             =   7950
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   4695
      Picture         =   "frmEmailAdd.frx":071A
      Top             =   7935
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEmailAdd.frx":0AA4
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   555
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   840
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "WARNING - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   8
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEmailAdd.frx":0B55
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Index           =   0
      Left            =   1830
      TabIndex        =   7
      Top             =   60
      Width           =   5325
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   60
      Picture         =   "frmEmailAdd.frx":0BF5
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmEmailAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sMailtoString As String
Dim Response As String

Private Type MailConfig
      MailServerPort As Integer
      MailServer As String
      MailFrom As String
      MailTo As String
      Subject As String
      MailBody As String
      MailMess As String
      StrDate As String
End Type

Private TMail As MailConfig

Private Sub cmdCancel_Click()
      Me.Hide
End Sub

Private Sub cmdSend_Click()

      cmdSend.Enabled = False
      If Option1.Value And sMailtoString <> "" Then
         Call ShellExecute(hwnd, "Open", sMailtoString, "", "", 1)
      Else
         ' Send via smtp
         On Error Resume Next
         TMail.MailTo = Trim(Me.txtTO)
         TMail.Subject = Trim(Me.txtSubject)
         TMail.StrDate = Format(Now, "ddd, dd mmm yyyy hh:mm:ss  +0200")
         TMail.MailMess = Me.txtMessage
         TMail.MailFrom = Me.txtFrom
    
         If isVaildEmail(TMail.MailFrom) = False Then
            MsgBox "You have not enter a invalid email address the mail will not be sent.", vbCritical, "Inviald E-Mail Address"
            txtAddress.SetFocus
            txtAddress.SelStart = 0
            txtAddress.SelLength = Len(txtAddress)
            cmdSend.Enabled = True

            Exit Sub
         End If
         If Len(TMail.Subject) <= 0 Then
            TMail.Subject = "No Subject...."
         End If
            
         ' Main mail message body
         TMail.MailBody = "Date: " & TMail.StrDate & vbCrLf _
         & "From: """ & Me.txtName & """ " & "<" & TMail.MailFrom & ">" & vbCrLf _
         & "X-Mailer: AbsoluteGenius Mailer V1" & vbCrLf _
         & "X-Accept-Language: en" & vbCrLf _
         & "MIME-Version: 1.0" & vbCrLf _
         & "To: ""Dominator Unlock"" <" & TMail.MailTo & "> " & vbCrLf _
         & "Subject: " & TMail.Subject & vbCrLf _
         & "Content-Type: text/html;" & vbCrLf _
         & vbTab & "charset=" & Chr(34) & "iso-8859-1" & Chr(34) & vbCrLf _
         & "Content-Transfer-Encoding: 7bit" & vbCrLf _
         & vbCrLf & TMail.MailMess & vbCrLf & "." & vbCrLf
         TMail.MailServer = "mail.btinternet.com"
         TMail.MailServer = "smtp.hotpop.com"
         TMail.MailServerPort = 25
         SendEmail
         cmdSend.Enabled = True
      End If
End Sub
Sub SendEmail()
Dim MailBody As String

      On Error Resume Next
      smtp.Close
      smtp.LocalPort = 0
      smtp.Protocol = sckTCPProtocol
      ' smtp.RemoteHost = TMail.MailServer
      ' smtp.RemotePort = TMail.MailServerPort
      ' smtp.Connect
      imgstatus.Picture = Image1(2).Picture
      imgstatus.Refresh
      lblStatus.Caption = "Status: Logging in"
      lblStatus.Refresh
      
      smtp.RemoteHost = "pop.hotpop.com"
      smtp.RemotePort = "110"
      smtp.Connect
      
      If WaitFor("+OK") = False Then
         lblStatus.Caption = "Status: Failed to Login"
         lblStatus.Refresh
         Exit Sub
      End If
      Reply "USER mrenigma"
      If WaitFor("+OK") = False Then
         lblStatus.Caption = "Status: Failed to Login"
         lblStatus.Refresh
         Exit Sub
      End If
      Reply "PASS svompa"
      If WaitFor("+OK") = False Then
         lblStatus.Caption = "Status: Failed to Login"
         lblStatus.Refresh
         Exit Sub
      End If
      Reply "STAT"
      If WaitFor("+OK") = False Then
         lblStatus.Caption = "Status: Failed to Login"
         lblStatus.Refresh
         Exit Sub
      End If
      ' Reply "RETR"
      ' If WaitFor("+OK") = False Then
      ' lblStatus.Caption = "Status: Failed to Login"
      ' lblStatus.Refresh
      ' Exit Sub
      ' End If
      
      
      Reply "QUIT"
      smtp.Close
      smtp.RemoteHost = TMail.MailServer
      smtp.RemotePort = TMail.MailServerPort
      smtp.Connect
      If WaitFor("220") = False Then
         lblStatus.Caption = "Status: Mail send failed"
         lblStatus.Refresh
         Exit Sub
      End If
      lblStatus.Caption = "Status: Connecting to " & smtp.RemoteHost
      lblStatus.Refresh
      Reply "HELO mrenigma@hotpop.com"
      If WaitFor("250") = False Then
         lblStatus.Caption = "Status: Mail send failed"
         lblStatus.Refresh
         Exit Sub
      End If
      lblStatus.Caption = "Status: Sending mail message"
      lblStatus.Refresh
      '
      Reply "MAIL FROM: mrenigma@hotpop.com"
      If WaitFor("250") = False Then
         lblStatus.Caption = "Status: Mail send failed"
         lblStatus.Refresh
         Exit Sub
      End If
      Reply "RCPT TO: " & TMail.MailTo
      If WaitFor("250") = False Then
         lblStatus.Caption = "Status: Mail send failed"
         lblStatus.Refresh
         Exit Sub
      End If
      Reply "DATA"
      If WaitFor("354") = False Then
         lblStatus.Caption = "Status: Mail send failed"
         lblStatus.Refresh
         Exit Sub
      End If
      Reply TMail.MailBody
      If WaitFor("250") = False Then
         lblStatus.Caption = "Status: Mail send failed"
         lblStatus.Refresh
         Exit Sub
      End If
      lblStatus.Caption = "Status: Mail message sent"
      lblStatus.Refresh
      Reply "QUIT"
      lblStatus.Caption = "Status: Closing connection."
      lblStatus.Refresh
      WaitFor ("221")
      smtp.Close
      imgstatus.Picture = Me.Image1(1).Picture
      imgstatus.Refresh
      lblStatus.Caption = "Status : Ide"
      lblStatus.Refresh
      Me.Hide
End Sub
Function WaitFor(ResponseCode As String) As Boolean
Dim start As Long
Dim tmr As Long

      ' This code in this function was not writen by me just found on the net
      ' But just like to say thank's to who ever did write it.
      WaitFor = True
      start = Timer ' Time event so won't get stuck in loop
      While Len(Response) = 0
         tmr = Timer - start
         DoEvents ' Let System keep checking for incoming response **IMPORTANT**
         If tmr > 10 Then
            ' Time in seconds to wait
            WaitFor = False
            MsgBox "SMTP service error, timed out while waiting for response", 64
            Exit Function
         End If
      Wend
      While Left(Response, 3) <> ResponseCode
         tmr = Timer - start
         DoEvents
         If tmr > 10 Then
            WaitFor = False
            Exit Function
         End If
      Wend
      Response = "" ' Sent response code to blank **IMPORTANT**
End Function
Sub Reply(StrBuff As String)
      If smtp.State = sckConnected Then
         smtp.SendData StrBuff & vbCrLf
         trace StrBuff & vbCrLf
      End If
End Sub
Sub trace(sText As String)
      Text1 = Text1 & sText
      Text1.SelStart = Len(Text1)
      Text1.Refresh
End Sub

Private Sub Form_Load()
      imgstatus.Picture = Image1(1).Picture
End Sub

Private Sub Option1_Click()
      Me.Height = 3480
      Me.cmdSend.Top = 2490
      Me.cmdCancel.Top = 2490
      Me.frmEmail.Visible = False
End Sub

Private Sub Option2_Click()
      Me.Height = 8835
      Me.cmdSend.Top = 7830
      Me.cmdCancel.Top = 7830
      Me.frmEmail.Visible = True
End Sub

Private Sub smtp_DataArrival(ByVal bytesTotal As Long)
      smtp.GetData Response
      trace Response
End Sub

Private Sub txtAddress_Change()
      Me.txtFrom = txtAddress
End Sub
