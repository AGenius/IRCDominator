VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrefs 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Preferences"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6315
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   5310
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   5310
      Width           =   1365
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5205
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9181
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   12632064
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPrefs.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notifications"
      TabPicture(1)   =   "frmPrefs.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraNotifications"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Messages"
      TabPicture(2)   =   "frmPrefs.frx":0F02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmMessageFormat"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Whispers"
      TabPicture(3)   =   "frmPrefs.frx":0F1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraWhispers"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Server Info"
      TabPicture(4)   =   "frmPrefs.frx":0F3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraNotifications 
         Height          =   4550
         Left            =   -74850
         TabIndex        =   48
         Top             =   450
         Width           =   6000
         Begin VB.CommandButton cmdSounds 
            Height          =   555
            Left            =   1410
            Picture         =   "frmPrefs.frx":0F56
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   1200
            Width           =   555
         End
         Begin VB.CheckBox chkNotifyJoins 
            Caption         =   "Show Joins"
            Height          =   255
            Left            =   150
            TabIndex        =   52
            Tag             =   "Chat|ShowJoins"
            ToolTipText     =   "Select this to show\nguests who join the room"
            Top             =   330
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin VB.CheckBox chkNotifyLeaves 
            Caption         =   "Show Departures"
            Height          =   255
            Left            =   150
            TabIndex        =   51
            Tag             =   "Chat|ShowLeaves"
            ToolTipText     =   "Select this to show guests\nthat leave the room"
            Top             =   600
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin VB.CheckBox chkNotifyAways 
            Caption         =   "Away Changes"
            Height          =   255
            Left            =   150
            TabIndex        =   50
            Tag             =   "Chat|ShowAways"
            ToolTipText     =   "Select this to show when guests\nhave set themselves away or back"
            Top             =   870
            Value           =   1  'Checked
            Width           =   1665
         End
         Begin VB.CheckBox chkNotifySounds 
            Caption         =   "Play Sounds"
            Height          =   255
            Left            =   150
            TabIndex        =   49
            Tag             =   "Chat|PlaySounds"
            ToolTipText     =   "Select this to play notification sounds"
            Top             =   1140
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Notifications"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   54
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Frame frmMessageFormat 
         Height          =   4550
         Left            =   -74850
         TabIndex        =   38
         Top             =   450
         Width           =   6000
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   1455
            Left            =   120
            TabIndex        =   55
            Top             =   2970
            Width           =   2445
            Begin VB.ComboBox cboSize 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmPrefs.frx":1820
               Left            =   150
               List            =   "frmPrefs.frx":182D
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Tag             =   "Chat|DisplaySize"
               Top             =   570
               Width           =   2025
            End
            Begin VB.CheckBox chkNoFormat 
               Caption         =   "No Text Formatting"
               Height          =   255
               Left            =   180
               TabIndex        =   56
               Tag             =   "Chat|NoFormatting"
               Top             =   1110
               Width           =   1965
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Font Size for messages:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   6
               Left            =   150
               TabIndex        =   59
               Top             =   360
               Width           =   2040
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Chat Display"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   5
               Left            =   90
               TabIndex        =   58
               Top             =   0
               Width           =   1080
            End
         End
         Begin VB.ComboBox cboFonts 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   780
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Tag             =   "Chat|Font"
            Top             =   660
            Width           =   3705
         End
         Begin VB.CheckBox chkChatFontBold 
            Caption         =   "Bold"
            Height          =   255
            Left            =   1560
            TabIndex        =   40
            Tag             =   "Chat|StyleBold"
            Top             =   1110
            Width           =   765
         End
         Begin VB.CheckBox chkChatFontItalic 
            Caption         =   "Italic"
            Height          =   255
            Left            =   2430
            TabIndex        =   39
            Tag             =   "Chat|StyleItalic"
            Top             =   1110
            Width           =   765
         End
         Begin MSComctlLib.ImageCombo cboChatColours 
            Height          =   330
            Left            =   780
            TabIndex        =   42
            Tag             =   "Chat|Colour"
            Top             =   1530
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Message Format"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Use These Settings for my messages I send"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   46
            Top             =   360
            Width           =   3990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Style:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   780
            TabIndex        =   45
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   44
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "Colour:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   43
            Top             =   1560
            Width           =   615
         End
      End
      Begin VB.Frame fraWhispers 
         Height          =   4550
         Left            =   -74850
         TabIndex        =   20
         Top             =   450
         Width           =   6000
         Begin VB.CheckBox chkWhisperWindows 
            Caption         =   "Show Whispers in seperate Windows"
            Height          =   255
            Left            =   150
            TabIndex        =   36
            Tag             =   "Chat|WhisperWindow"
            ToolTipText     =   "Select this to show whispered messages\n in seperate windows"
            Top             =   3000
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.Frame fraWhisperMessages 
            Height          =   2235
            Left            =   150
            TabIndex        =   23
            Top             =   750
            Width           =   5745
            Begin VB.CheckBox chkNotifyWhispers 
               Caption         =   "Notify Request On Screen"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Tag             =   "Chat|WhisperNotify"
               ToolTipText     =   "Select this to notify on screen when\nsomeone is trying to whisper you\nwhile whispers are turned off"
               Top             =   1860
               Value           =   1  'Checked
               Width           =   2565
            End
            Begin VB.ComboBox cboWhisperFont 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2790
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Tag             =   "Chat|WhisperFont"
               Top             =   1410
               Width           =   2805
            End
            Begin VB.CheckBox chkWhisperBold 
               Caption         =   "Bold"
               Height          =   255
               Left            =   3600
               TabIndex        =   28
               Tag             =   "Chat|WhisperStyleBold"
               Top             =   1860
               Width           =   765
            End
            Begin VB.CheckBox chkWhisperItalic 
               Caption         =   "Italic"
               Height          =   255
               Left            =   4470
               TabIndex        =   27
               Tag             =   "Chat|WhisperStyleItalic"
               Top             =   1860
               Width           =   765
            End
            Begin VB.OptionButton optPVTMessage 
               Caption         =   "On Screen Message"
               Height          =   285
               Left            =   1890
               TabIndex        =   26
               Tag             =   "Chat|WhisperPrivMessage"
               ToolTipText     =   "Select this to send a private message\nwith your response"
               Top             =   300
               Width           =   2535
            End
            Begin VB.OptionButton optWhisper 
               Caption         =   "Whisper Back"
               Height          =   285
               Left            =   210
               TabIndex        =   25
               Tag             =   "Chat|WhisperWhisper"
               ToolTipText     =   "Select this to whisper back your response"
               Top             =   300
               Value           =   -1  'True
               Width           =   1515
            End
            Begin VB.TextBox txtWhisperResponse 
               Height          =   675
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   24
               Tag             =   "Chat|WhisperResponse"
               Top             =   660
               Width           =   5475
            End
            Begin MSComctlLib.ImageCombo cboWhisperColour 
               Height          =   330
               Left            =   810
               TabIndex        =   31
               Tag             =   "Chat|WhisperColour"
               Top             =   1410
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   582
               _Version        =   393216
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Style:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   11
               Left            =   2820
               TabIndex        =   35
               Top             =   1890
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Font:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   10
               Left            =   2265
               TabIndex        =   34
               Top             =   1500
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Colour:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   33
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Send This Response to Whisper "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   32
               Top             =   30
               Width           =   2805
            End
         End
         Begin VB.CheckBox chkWhisperMessage 
            Caption         =   "Respond With Message"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Tag             =   "Chat|WhisperMessage"
            ToolTipText     =   "Select this to send a response\nmessage to the guest"
            Top             =   510
            Value           =   1  'Checked
            Width           =   2505
         End
         Begin VB.CheckBox chkNoWhispers 
            Caption         =   "Turn Whispers Off"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Tag             =   "Chat|NoWhispers"
            ToolTipText     =   "Select this to disable whispers"
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Whispers"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   37
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4550
         Left            =   -74850
         TabIndex        =   14
         Top             =   450
         Width           =   6000
         Begin VB.CommandButton cmdChatX 
            Caption         =   "Set CLSID"
            Height          =   315
            Left            =   2250
            TabIndex        =   72
            Top             =   1470
            Width           =   885
         End
         Begin VB.CheckBox chkChatX 
            Caption         =   "Use MSNChatX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   71
            Top             =   1500
            Width           =   1905
         End
         Begin VB.TextBox txtClassID 
            Height          =   285
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "Make sure this CLSID is correct or it wont connect"
            Top             =   1095
            Width           =   5100
         End
         Begin VB.TextBox txtIP 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "This is the IP Address of the MSN Chat Server"
            Top             =   540
            Width           =   2445
         End
         Begin VB.TextBox txtOCXVersion 
            Height          =   285
            Left            =   2790
            TabIndex        =   15
            ToolTipText     =   "This is the Chat OCX Version"
            Top             =   540
            Width           =   2445
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "MSN Chat CLSID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   67
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "MSN Chat Server Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   2490
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "MSN Chat Server IP Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   2460
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "MSN Chat OCX Version"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   2790
            TabIndex        =   17
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4550
         Left            =   150
         TabIndex        =   3
         Top             =   450
         Width           =   6000
         Begin VB.CheckBox chkMOTD 
            Caption         =   "Show Message Of the Day at Startup"
            Height          =   255
            Left            =   2745
            TabIndex        =   70
            Tag             =   "Chat|MOTD"
            ToolTipText     =   "Tick this to see the Message of the Day\nWhen IRCDominator Starts up"
            Top             =   870
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin VB.CheckBox chkShowOCX 
            Caption         =   "Show Register Of OCX's"
            Height          =   255
            Left            =   2745
            TabIndex        =   69
            Tag             =   "Chat|Show OCXs"
            ToolTipText     =   "This option is to test if the Registering of the Chat OCX's is working"
            Top             =   600
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin VB.CheckBox chkToolTips 
            Caption         =   "Show ToolTips"
            Height          =   255
            Left            =   2745
            TabIndex        =   68
            Tag             =   "Chat|AutoJoinKick"
            Top             =   315
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin VB.Frame fraAlive 
            Height          =   1395
            Left            =   90
            TabIndex        =   60
            Top             =   3060
            Width           =   4515
            Begin VB.CheckBox chkTestAlive 
               Caption         =   "Test Connection is Alive if no activity after"
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Tag             =   "Chat|TestAlive"
               ToolTipText     =   "Select this to perform a\ntest to check your connection\nis still alive on the server\nand reconnect if no response"
               Top             =   330
               Width           =   3285
            End
            Begin MSComctlLib.Slider sldAlive 
               Height          =   510
               Left            =   360
               TabIndex        =   62
               Tag             =   "Chat|AliveTimer"
               Top             =   690
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   900
               _Version        =   393216
               LargeChange     =   1
               Min             =   1
               SelStart        =   5
               TickStyle       =   1
               Value           =   5
               TextPosition    =   1
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Test Alive Status"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   65
               Top             =   0
               Width           =   1470
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C000&
               BackStyle       =   0  'Transparent
               Caption         =   "Minutes"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   13
               Left            =   3780
               TabIndex        =   64
               Top             =   360
               Width           =   555
            End
            Begin VB.Label lblAliveMins 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3480
               TabIndex        =   63
               Top             =   360
               Width           =   225
            End
         End
         Begin VB.CheckBox chkAutoAfterKick 
            Caption         =   "AutoRejoin When Kicked"
            Height          =   255
            Left            =   150
            TabIndex        =   9
            Tag             =   "Chat|AutoJoinKick"
            Top             =   330
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin VB.CheckBox chkLocalTime 
            Caption         =   "Mask Local Time requests"
            Height          =   255
            Left            =   150
            TabIndex        =   8
            Tag             =   "Chat|MaskLocalTime"
            ToolTipText     =   "Select this to not send your local time nfo back\nbut instead send the message below"
            Top             =   1140
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin VB.TextBox txtTimeReply 
            Height          =   675
            Left            =   390
            MultiLine       =   -1  'True
            TabIndex        =   7
            Tag             =   "Chat|LocalTime"
            Text            =   "frmPrefs.frx":1847
            ToolTipText     =   "Enter your text to appear when some one checks your time"
            Top             =   1440
            Width           =   5475
         End
         Begin VB.CheckBox chkTryJoin 
            Caption         =   "Try Join - Until Joined"
            Height          =   285
            Left            =   150
            TabIndex        =   6
            Tag             =   "Chat|TryJoin"
            ToolTipText     =   "Use full when Keep alive is set\nso it will attempt to rejoin a room\nuntil your previous nick has fallen out"
            Top             =   2220
            Width           =   1905
         End
         Begin VB.CheckBox chkShowTrace 
            Caption         =   "Show Server Trace"
            Height          =   255
            Left            =   150
            TabIndex        =   5
            Tag             =   "Chat|ShowTrace"
            Top             =   870
            Width           =   3075
         End
         Begin VB.CheckBox chkAutoJoin 
            Caption         =   "AutoJoin When connected"
            Height          =   255
            Left            =   150
            TabIndex        =   4
            Tag             =   "Chat|AutoJoin"
            Top             =   600
            Value           =   1  'Checked
            Width           =   3075
         End
         Begin MSComctlLib.Slider sldJoin 
            Height          =   450
            Left            =   450
            TabIndex        =   10
            Tag             =   "Chat|RejoinTimer"
            Top             =   2520
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   794
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   20
            SelStart        =   5
            TickStyle       =   1
            Value           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "General Settings"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seconds"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   2340
            TabIndex        =   12
            Top             =   2265
            Width           =   660
         End
         Begin VB.Label lblJoinSecs 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2070
            TabIndex        =   11
            Top             =   2265
            Width           =   225
         End
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objTooltip    As cTooltip
Dim bLoading As Boolean

Private Sub chkChatX_Click()
      If chkChatX.Value = 1 Then
         cmdChatX.Enabled = True
      Else
         cmdChatX.Enabled = False
      End If
End Sub

Private Sub chkTestAlive_Click()
      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If chkTestAlive.Value = 1 Then
         sldAlive.Enabled = True
      Else
         sldAlive.Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub

Private Sub chkTryJoin_Click()
      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If chkTryJoin.Value = 1 Then
         Me.sldJoin.Enabled = True
      Else
         Me.sldJoin.Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub
Public Sub chkWhisperMessage_Click()
      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If chkWhisperMessage.Value = 1 Then
         optPVTMessage.Enabled = True
         optWhisper.Enabled = True
         txtWhisperResponse.Enabled = True
         Me.cboWhisperColour.Enabled = True
         Me.cboWhisperFont.Enabled = True
         Me.chkWhisperBold.Enabled = True
         Me.chkWhisperItalic.Enabled = True
      Else
         optPVTMessage.Enabled = False
         optWhisper.Enabled = False
         txtWhisperResponse.Enabled = False
         Me.cboWhisperColour.Enabled = False
         Me.cboWhisperFont.Enabled = False
         Me.chkWhisperBold.Enabled = False
         Me.chkWhisperItalic.Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub
Private Sub chkNoWhispers_Click()

      If Not (bLoading) Then

         ' ***-CodeSmart Linker TagStart | Please Do Not Modify
         If chkNoWhispers.Value = 1 Then
            chkWhisperBold.Enabled = True
            chkWhisperItalic.Enabled = True
            chkWhisperMessage.Enabled = True
            cboWhisperColour.Enabled = True
            optWhisper.Enabled = True
            txtWhisperResponse.Enabled = True
            cboWhisperFont.Enabled = True
            Me.chkNotifyWhispers.Enabled = True
            chkWhisperMessage_Click
         Else
            chkWhisperBold.Enabled = False
            chkWhisperItalic.Enabled = False
            chkWhisperMessage.Enabled = False
            cboWhisperColour.Enabled = False
            optWhisper.Enabled = False
            txtWhisperResponse.Enabled = False
            cboWhisperFont.Enabled = False
            optPVTMessage.Enabled = False
            optWhisper.Enabled = False
            txtWhisperResponse.Enabled = False
            Me.chkNotifyWhispers.Enabled = False
         End If
         ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
      End If
End Sub
Private Sub cmdCancel_Click()
      Me.Hide
End Sub
Private Sub cmdDefaults_Click()
Dim i As Integer

      cboFonts.ListIndex = FindInCombo(cboFonts, "Tahoma")
      cboSize.ListIndex = FindInCombo(cboSize, "Medium")
      cboChatColours.ComboItems.Item(1).Selected = True
      Me.chkChatFontBold.Value = 0
      Me.chkChatFontItalic.Value = 0
      Me.chkNoFormat.Value = 0
      Me.chkNotifyAways.Value = 1
      Me.chkNotifyJoins.Value = 1
      Me.chkNotifyLeaves.Value = 1
      Me.chkNotifySounds.Value = 1
      Me.chkNoWhispers.Value = 0
      Me.chkWhisperWindows.Value = 1
End Sub

Private Sub cmdChatX_Click()
      Me.txtClassID = "ECCDBA05-B58F-4509-AE26-CF47B2FFC3FE"
End Sub

Private Sub cmdOk_Click()
      ' Save_Settings Me
      SaveSettings
      GeneralSettings.SavePrefs
      Me.Hide
End Sub

Private Sub cmdSounds_Click()
      InstallSounds
Dim T As Double
      On Error Resume Next
      T = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", 5)
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim cboItem As ComboItem
      ' *** CodeSmart ErrorHead TagStart | Please Do Not  Modify
      ' Code Added By CodeSmart
      ' =============================================================================
100   On Error GoTo Err_Form_Load:
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      ' *** CodeSmart ErrorHead TagEnd | Please Do Not Modify
        
      ' IniFile.Path = App.Path & App.FileDescription & "\.dll"
101   LoadToolTips Me, m_objTooltip
102   IniFile.Section = "General"
103   bLoading = True
104   Call BuildFontList(cboFonts)
105   Call BuildFontList(Me.cboWhisperFont)
106   Call BuildColourList(Me.cboChatColours)
107   Call BuildColourList(Me.cboWhisperColour)
      ' 113   Load_Settings Me
108   Call PopulateSettings
109   bLoading = False
110   chkNoWhispers_Click
111   chkTestAlive_Click
112   lblAliveMins.Caption = sldAlive.Value
113   lblJoinSecs.Caption = Me.sldJoin.Value
      ' If bActivated Then Me.Check1.Visible = True

      ' *** CodeSmart ErrorFoot TagStart | Please Do Not Modify
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
114   Exit Sub
115 Err_Form_Load:
116   MsgBox ("Error Encounterd in frmPrefs @ " & Erl & " " & Err.Description)
      ' =============================================================================
117 Exit_Form_Load:
      ' *** CodeSmart ErrorFoot TagEnd | Please Do Not Modify
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         Me.Hide
         Cancel = True
      End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
End Sub
Private Sub sldAlive_Scroll()
      lblAliveMins.Caption = sldAlive.Value
End Sub
Private Sub sldJoin_Scroll()
      lblJoinSecs.Caption = Me.sldJoin.Value
End Sub

Private Sub PopulateSettings()
      ' *** CodeSmart ErrorHead TagStart | Please Do Not  Modify
      ' Code Added By CodeSmart
      ' =============================================================================
      On Error GoTo Err_PopulateSettings:
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      ' *** CodeSmart ErrorHead TagEnd | Please Do Not Modify
100   Me.cboChatColours.ComboItems.Item(GeneralSettings.Chat_Colour).Selected = True
101   Me.cboWhisperColour.ComboItems.Item(GeneralSettings.Whisper_Colour).Selected = True
102   Me.cboFonts.Text = GeneralSettings.Chat_Font
103   Me.cboSize.ListIndex = GeneralSettings.Chat_DisplaySize
104   Me.cboWhisperFont.Text = GeneralSettings.Whisper_Font
105   Me.chkAutoAfterKick.Value = Abs(GeneralSettings.AutoJoinKick)
106   Me.chkAutoJoin.Value = Abs(GeneralSettings.AutoJoin)
107   Me.chkChatFontBold.Value = Abs(GeneralSettings.Chat_StyleBold)
108   Me.chkChatFontItalic.Value = Abs(GeneralSettings.Chat_StyleItalic)
109   Me.chkLocalTime.Value = Abs(GeneralSettings.MaskLocalTime)
110   Me.chkNoFormat.Value = Abs(GeneralSettings.NoFormatting)
111   Me.chkNotifyAways.Value = Abs(GeneralSettings.Notify_Aways)
112   Me.chkNotifyJoins.Value = Abs(GeneralSettings.Notify_Joins)
113   Me.chkNotifyLeaves.Value = Abs(GeneralSettings.Notify_Leaves)
114   Me.chkNotifySounds.Value = Abs(GeneralSettings.PlaySounds)
115   Me.chkNoWhispers.Value = Abs(GeneralSettings.Whisper_NoWhispers)
116   Me.chkWhisperMessage = Abs(GeneralSettings.Whisper_Message)
117   Me.chkShowTrace.Value = Abs(GeneralSettings.ShowTrace)
118   Me.chkTestAlive.Value = Abs(GeneralSettings.TestAlive)
119   Me.chkTryJoin.Value = Abs(GeneralSettings.TryJoin)
120   Me.chkWhisperBold = Abs(GeneralSettings.Whisper_StyleBold)
121   Me.chkWhisperItalic = Abs(GeneralSettings.Whisper_StyleItalic)
122   Me.chkWhisperMessage = Abs(GeneralSettings.Whisper_Message)
123   Me.chkWhisperWindows = Abs(GeneralSettings.Whisper_Window)
124   Me.optPVTMessage.Value = Abs(GeneralSettings.Whisper_PrivMessage)
125   Me.optWhisper.Value = Abs(GeneralSettings.Whisper_Whisper)
126   Me.sldAlive.Value = GeneralSettings.AliveTime
127   Me.sldJoin.Value = GeneralSettings.RejoinTimer
128   Me.txtTimeReply = GeneralSettings.LocalTime
129   Me.txtWhisperResponse = GeneralSettings.Whisper_Response
130   Me.txtIP = GeneralSettings.ServerIP
131   Me.txtOCXVersion = GeneralSettings.ChatOCXVersion
132   Me.txtClassID = GeneralSettings.ChatCLASSID
133   Me.chkToolTips = Abs(GeneralSettings.ShowToolTips)
134   Me.chkShowOCX = Abs(GeneralSettings.ShowOCXs)
135   Me.chkMOTD = Abs(GeneralSettings.ShowMOTD)
136   Me.chkChatX = Abs(GeneralSettings.Chat_ChatX)
      If chkChatX.Value = 1 Then
         cmdChatX.Enabled = True
      Else
         cmdChatX.Enabled = False
      End If

      ' *** CodeSmart ErrorFoot TagStart | Please Do Not Modify
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      Exit Sub
Err_PopulateSettings:
      MsgBox ("Error Encounterd in PopulateSettings @ " & Erl & " " & Err.Description)
      ' =============================================================================
Exit_PopulateSettings:
      ' *** CodeSmart ErrorFoot TagEnd | Please Do Not Modify
End Sub

Private Sub SaveSettings()
      GeneralSettings.Chat_Colour = Me.cboChatColours.SelectedItem.Index
      GeneralSettings.Whisper_Colour = Me.cboWhisperColour.SelectedItem.Index
      GeneralSettings.Chat_Font = Me.cboFonts.Text
      GeneralSettings.Chat_DisplaySize = Me.cboSize.ListIndex
      GeneralSettings.Whisper_Font = Me.cboWhisperFont.Text
      GeneralSettings.AutoJoinKick = Me.chkAutoAfterKick.Value
      GeneralSettings.AutoJoin = Me.chkAutoJoin.Value
      GeneralSettings.Chat_StyleBold = Me.chkChatFontBold.Value
      GeneralSettings.Chat_StyleItalic = Me.chkChatFontItalic.Value
      GeneralSettings.MaskLocalTime = Me.chkLocalTime.Value
      GeneralSettings.NoFormatting = Me.chkNoFormat.Value
      GeneralSettings.Notify_Aways = Me.chkNotifyAways.Value
      GeneralSettings.Notify_Joins = Me.chkNotifyJoins.Value
      GeneralSettings.Notify_Leaves = Me.chkNotifyLeaves.Value
      GeneralSettings.PlaySounds = Me.chkNotifySounds.Value
      GeneralSettings.Welcome_Active = Me.chkNoWhispers.Value
      GeneralSettings.ShowTrace = Me.chkShowTrace.Value
      GeneralSettings.TestAlive = Me.chkTestAlive.Value
      GeneralSettings.TryJoin = Me.chkTryJoin.Value
      GeneralSettings.Whisper_Message = Me.chkWhisperMessage.Value
      GeneralSettings.Whisper_NoWhispers = Me.chkNoWhispers.Value
      GeneralSettings.Whisper_StyleBold = Me.chkWhisperBold
      GeneralSettings.Whisper_StyleItalic = Me.chkWhisperItalic
      GeneralSettings.Whisper_Message = Me.chkWhisperMessage
      GeneralSettings.Whisper_Window = Me.chkWhisperWindows
      GeneralSettings.Whisper_PrivMessage = Me.optPVTMessage.Value
      GeneralSettings.Whisper_Whisper = Me.optWhisper.Value
      GeneralSettings.AliveTime = Me.sldAlive.Value
      GeneralSettings.RejoinTimer = Me.sldJoin.Value
      GeneralSettings.LocalTime = Me.txtTimeReply
      GeneralSettings.Whisper_Response = Me.txtWhisperResponse
      GeneralSettings.ServerIP = Me.txtIP
      GeneralSettings.ChatOCXVersion = Me.txtOCXVersion
      GeneralSettings.ChatCLASSID = Me.txtClassID
      GeneralSettings.ShowToolTips = Me.chkToolTips
      GeneralSettings.ShowOCXs = Me.chkShowOCX
      GeneralSettings.ShowMOTD = Me.chkMOTD
      GeneralSettings.Chat_ChatX = Me.chkChatX

End Sub

