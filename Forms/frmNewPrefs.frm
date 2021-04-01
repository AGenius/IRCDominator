VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "Threed20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrefs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MosFax Properties Page"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6270
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6270
   Begin Threed.SSPanel pnlSettings 
      Height          =   3465
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   5220
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   6112
      _Version        =   131073
      BackStyle       =   1
      BevelOuter      =   1
      Begin VB.CommandButton cmdBrowseViewer 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   63
         Top             =   1440
         Width           =   285
      End
      Begin VB.CommandButton cmdBrowseLogos 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   66
         Top             =   780
         Width           =   285
      End
      Begin VB.TextBox txtLogoPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   65
         Tag             =   "FaxImage|LogoPath"
         Top             =   780
         Width           =   3315
      End
      Begin VB.TextBox txtViewerPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   64
         Tag             =   "General|ViewerPath"
         Top             =   1440
         Width           =   3315
      End
      Begin VB.CommandButton cmdBrowseDataBase 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3480
         TabIndex        =   62
         Top             =   2040
         Width           =   285
      End
      Begin VB.TextBox txtDataBase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   61
         Tag             =   "General|DataBase"
         Top             =   2040
         Width           =   3315
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   345
         Index           =   18
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   609
         _Version        =   131073
         ForeColor       =   8388608
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Miscellaneous Settings"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Path to Logo's:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   18
         Top             =   1140
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Path to Viewer:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   19
         Top             =   1740
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "DataBase Path:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSCheck chkStartPaused 
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Tag             =   "General|Paused"
         ToolTipText     =   "Enable this to record the history for each fax"
         Top             =   2760
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Paused at Startup"
         Value           =   1
      End
      Begin Threed.SSCheck chkFaxHistory 
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Tag             =   "General|Auto Follow"
         ToolTipText     =   "Enable this to record the history for each fax"
         Top             =   2490
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Auto Follow Faxes in progress"
      End
   End
   Begin VB.ListBox lstPanel 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      ItemData        =   "frmNewPrefs.frx":27A2
      Left            =   240
      List            =   "frmNewPrefs.frx":27AF
      TabIndex        =   20
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A5BFC2&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   510
      Width           =   6825
      Begin VB.Line linered 
         BorderColor     =   &H000000FF&
         Index           =   2
         X1              =   30
         X2              =   8000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.ComboBox cboComPort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Tag             =   "Modem|ComPort"
      Text            =   "1"
      Top             =   600
      Width           =   705
   End
   Begin VB.TextBox txtLocalID 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Tag             =   "General|LocalID"
      Text            =   "MOS Computers"
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox txtDialPrefix 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Tag             =   "Modem|DialPrefix"
      Top             =   1050
      Width           =   615
   End
   Begin Threed.SSCommand cmdAutoDetect 
      Height          =   315
      Left            =   3750
      TabIndex        =   0
      Top             =   600
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   131073
      PictureFrames   =   1
      Windowless      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewPrefs.frx":27E2
      Caption         =   "Auto Detect"
      PictureAlignment=   9
   End
   Begin Threed.SSCommand cmdSaveSettings 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Save Changes to Settings"
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1058
      _Version        =   131073
      CaptionStyle    =   1
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewPrefs.frx":2C34
      Caption         =   "Save Settings"
      Alignment       =   4
      PictureAlignment=   9
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   600
      Left            =   240
      TabIndex        =   6
      Top             =   2790
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1058
      _Version        =   131073
      CaptionStyle    =   1
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewPrefs.frx":2F4E
      Caption         =   "Close Window"
      Alignment       =   4
      PictureAlignment=   9
      BevelWidth      =   0
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdSettings 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   240
      TabIndex        =   7
      Top             =   4170
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1058
      _Version        =   131073
      CaptionStyle    =   1
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewPrefs.frx":3954
      Caption         =   "Report Settings"
      Alignment       =   4
      PictureAlignment=   9
      BevelWidth      =   0
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin Threed.SSPanel lblHeadings 
      Height          =   285
      Index           =   0
      Left            =   780
      TabIndex        =   8
      Top             =   600
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   503
      _Version        =   131073
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Modem Communications Port:"
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   1
   End
   Begin Threed.SSPanel lblHeadings 
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   9
      Top             =   1050
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   131073
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Local Identity:"
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   1
   End
   Begin Threed.SSPanel lblHeadings 
      Height          =   285
      Index           =   2
      Left            =   3810
      TabIndex        =   10
      Top             =   1050
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   503
      _Version        =   131073
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Dial Prefix:"
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   1
   End
   Begin Threed.SSPanel lblHeadings 
      Height          =   315
      Index           =   15
      Left            =   570
      TabIndex        =   11
      Top             =   210
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   131073
      ForeColor       =   8388608
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Fax Preferences"
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   1
   End
   Begin Threed.SSPanel pnlIcon 
      Height          =   525
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   926
      _Version        =   131073
      PictureFrames   =   1
      BackStyle       =   1
      Picture         =   "frmNewPrefs.frx":482E
      BevelOuter      =   0
   End
   Begin Threed.SSPanel pnlIcon 
      Height          =   525
      Index           =   1
      Left            =   210
      TabIndex        =   13
      Top             =   540
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   926
      _Version        =   131073
      PictureFrames   =   1
      BackStyle       =   1
      Picture         =   "frmNewPrefs.frx":6FE0
      BevelOuter      =   0
   End
   Begin Threed.SSCommand cmdThemes 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   240
      TabIndex        =   14
      Top             =   4860
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1058
      _Version        =   131073
      CaptionStyle    =   1
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewPrefs.frx":7902
      Caption         =   "Theme Settings"
      Alignment       =   4
      PictureAlignment=   9
      BevelWidth      =   0
      RoundedCorners  =   0   'False
      Outline         =   0   'False
   End
   Begin Threed.SSPanel lblHeadings 
      Height          =   285
      Index           =   19
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   503
      _Version        =   131073
      ForeColor       =   8388608
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Select  Settings"
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin Threed.SSPanel pnlSettings 
      Height          =   3495
      Index           =   1
      Left            =   6030
      TabIndex        =   24
      Top             =   1680
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   6165
      _Version        =   131073
      BackStyle       =   1
      BevelOuter      =   1
      Begin VB.TextBox txtUpdateQueues 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         TabIndex        =   32
         Tag             =   "Queues|UpdateTime"
         Text            =   "10"
         ToolTipText     =   "Enter here the interval in seconds the queues will be refreshed"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtBotQueue 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         TabIndex        =   31
         Tag             =   "Queues|Bottom"
         Text            =   "3"
         Top             =   390
         Width           =   495
      End
      Begin VB.TextBox txtMaxRetry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         TabIndex        =   30
         Tag             =   "Queues|Max Retries"
         Text            =   "3"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtDialTime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         TabIndex        =   29
         Tag             =   "Modem|Dial Time"
         Text            =   "100"
         Top             =   1020
         Width           =   495
      End
      Begin VB.TextBox txtRings 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         TabIndex        =   28
         Tag             =   "General|Rings"
         Text            =   "3"
         Top             =   2520
         Width           =   495
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   345
         Index           =   20
         Left            =   150
         TabIndex        =   26
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   609
         _Version        =   131073
         ForeColor       =   8388608
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Fax Queues Settings"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   0
         Left            =   2490
         TabIndex        =   33
         Tag             =   "Bottom of queue"
         Top             =   390
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtBotQueue"
         BuddyDispid     =   196622
         OrigLeft        =   2250
         OrigTop         =   180
         OrigRight       =   2490
         OrigBottom      =   495
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   1
         Left            =   2490
         TabIndex        =   34
         Tag             =   "Max Retries"
         Top             =   720
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtMaxRetry"
         BuddyDispid     =   196623
         OrigLeft        =   2250
         OrigTop         =   570
         OrigRight       =   2490
         OrigBottom      =   885
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   2
         Left            =   2490
         TabIndex        =   35
         Tag             =   "Dialing time"
         Top             =   1020
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   60
         BuddyControl    =   "txtDialTime"
         BuddyDispid     =   196624
         OrigLeft        =   2250
         OrigTop         =   900
         OrigRight       =   2490
         OrigBottom      =   1215
         Max             =   200
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   3
         Left            =   2460
         TabIndex        =   36
         Tag             =   "Update Send Queue"
         Top             =   1800
         Width           =   225
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   60
         BuddyControl    =   "txtUpdateQueues"
         BuddyDispid     =   196621
         OrigLeft        =   2970
         OrigTop         =   5430
         OrigRight       =   3165
         OrigBottom      =   5715
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   5
         Left            =   2490
         TabIndex        =   37
         Tag             =   "Number of Rings"
         Top             =   2520
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtRings"
         BuddyDispid     =   196625
         OrigLeft        =   2925
         OrigTop         =   6030
         OrigRight       =   3165
         OrigBottom      =   6315
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin Threed.SSCheck chkReceive 
         Height          =   255
         Left            =   555
         TabIndex        =   38
         Tag             =   "General|Receive"
         ToolTipText     =   "Enables / Disables the receiving of faxes when no faxes are being sent"
         Top             =   2250
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Receive Faxes When Idle"
      End
      Begin Threed.SSCheck chkUpdateQueues 
         Height          =   255
         Left            =   555
         TabIndex        =   39
         Tag             =   "Queues|UpdateQueues"
         ToolTipText     =   "Enables/Disables the Refresh of the queues"
         Top             =   1530
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Refresh Queues Every"
         Value           =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   6
         Left            =   480
         TabIndex        =   40
         Top             =   390
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bottom of Queue:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   7
         Left            =   495
         TabIndex        =   41
         Top             =   720
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Maximum Retries:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   8
         Left            =   870
         TabIndex        =   42
         Top             =   1020
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Dialing Time:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   10
         Left            =   705
         TabIndex        =   43
         Top             =   2520
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Wait for Rings:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   9
         Left            =   1260
         TabIndex        =   44
         Top             =   1800
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Interval:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
   End
   Begin Threed.SSPanel pnlSettings 
      Height          =   3465
      Index           =   2
      Left            =   2070
      TabIndex        =   25
      Top             =   1680
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   6112
      _Version        =   131073
      BackStyle       =   1
      BevelWidth      =   2
      BevelOuter      =   1
      Begin VB.TextBox txtLines 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2910
         TabIndex        =   67
         Tag             =   "FaxImage|Detail Lines"
         Text            =   "64"
         Top             =   2250
         Width           =   375
      End
      Begin VB.TextBox txtOffset 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1170
         TabIndex        =   50
         Tag             =   "FaxImage|OffSet"
         Text            =   "60"
         Top             =   2220
         Width           =   420
      End
      Begin VB.ComboBox cboFont 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         Sorted          =   -1  'True
         TabIndex        =   49
         Tag             =   "FaxImage|FontName"
         Text            =   "Courier New"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboSize 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         TabIndex        =   48
         Tag             =   "FaxImage|FontSize"
         Text            =   "19"
         Top             =   1800
         Width           =   705
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00A5BFC2&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Index           =   21
         Left            =   150
         TabIndex        =   47
         Top             =   1290
         Width           =   3465
         Begin VB.Line linered 
            BorderColor     =   &H80000014&
            BorderStyle     =   6  'Inside Solid
            Index           =   4
            X1              =   30
            X2              =   8000
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00A5BFC2&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Index           =   20
         Left            =   150
         TabIndex        =   46
         Top             =   810
         Width           =   3465
         Begin VB.Line linered 
            BorderColor     =   &H00000000&
            BorderStyle     =   6  'Inside Solid
            Index           =   3
            X1              =   30
            X2              =   8000
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.ComboBox cboPreset 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         TabIndex        =   45
         Top             =   390
         Width           =   2055
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   345
         Index           =   21
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   609
         _Version        =   131073
         ForeColor       =   8388608
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Fax Image Settings"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   4
         Left            =   1590
         TabIndex        =   51
         Tag             =   "Offset"
         Top             =   2220
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   54
         BuddyControl    =   "txtOffset"
         BuddyDispid     =   196627
         OrigLeft        =   2790
         OrigTop         =   1920
         OrigRight       =   3030
         OrigBottom      =   2235
         Max             =   158
         Min             =   40
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin Threed.SSCommand cmdaddPreset 
         Height          =   375
         Left            =   660
         TabIndex        =   52
         Top             =   870
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   131073
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Add to preset list"
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   53
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Preset Font Styles:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   12
         Left            =   315
         TabIndex        =   54
         Top             =   1440
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Font Name:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   13
         Left            =   405
         TabIndex        =   55
         Top             =   1830
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Font Size:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   14
         Left            =   135
         TabIndex        =   56
         Top             =   2220
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Margin Offset:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
      Begin Threed.SSCheck chkBold 
         Height          =   255
         Left            =   1890
         TabIndex        =   57
         Tag             =   "FaxImage|FontBold"
         Top             =   1860
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bold font"
         Alignment       =   1
      End
      Begin Threed.SSCheck chkStretchLogo 
         Height          =   255
         Left            =   300
         TabIndex        =   58
         Tag             =   "FaxImage|StretchLogos"
         ToolTipText     =   "Enable this if you require old logo's to be stretched"
         Top             =   2550
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Stretch Logo's"
      End
      Begin Threed.SSCheck chkHighRes 
         Height          =   255
         Left            =   300
         TabIndex        =   59
         Tag             =   "FaxImage|HighRes"
         ToolTipText     =   "Enable this for High Resolution Fax Images (Takes longer to build and send but quality is much better)"
         Top             =   3060
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Build High Resolution Faxes"
      End
      Begin Threed.SSCheck chkDoubleLines 
         Height          =   255
         Left            =   300
         TabIndex        =   60
         Tag             =   "FaxImage|DoubleWidthLines"
         ToolTipText     =   "Enable this if you require boxes to be drawn at double width"
         Top             =   2790
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   131073
         BackStyle       =   1
         Windowless      =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Draw Double Width Lines"
         Value           =   1
      End
      Begin MSComCtl2.UpDown udControl 
         Height          =   285
         Index           =   6
         Left            =   3285
         TabIndex        =   68
         Tag             =   "Lines"
         Top             =   2250
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   503
         _Version        =   393216
         Value           =   54
         BuddyControl    =   "txtLines"
         BuddyDispid     =   196626
         OrigLeft        =   2790
         OrigTop         =   1920
         OrigRight       =   3030
         OrigBottom      =   2235
         Max             =   158
         Min             =   40
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin Threed.SSPanel lblHeadings 
         Height          =   285
         Index           =   16
         Left            =   1950
         TabIndex        =   69
         Top             =   2250
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detail Lines:"
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   1
      End
   End
   Begin VB.PictureBox FaxFinder 
      Height          =   480
      Left            =   7560
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   70
      Top             =   300
      Width           =   1200
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cFB As New cFormBackground
Private m_cFB_Imp As New cFormBackground
Private bDirty As Boolean

Private Sub chkReceive_Click(Value As Integer)
      bDirty = True
      If chkReceive.Value Then
         txtRings.Enabled = True
         udControl(5).Enabled = True
         lblHeadings(10).Enabled = True
      Else
         txtRings.Enabled = False
         udControl(5).Enabled = False
         lblHeadings(10).Enabled = False
      End If
End Sub

Private Sub cmdOk_Click()
      If bDirty Then
         ' Settings have change ask about saving them
         Select Case MsgBox("Settings have been changed!" & vbCrLf & vbCrLf & "Save Settings First?", vbYesNoCancel + vbQuestion + vbApplicationModal, "enFax Report Settings")
            Case vbYes
               Call SaveSettings
               Unload Me
            Case vbCancel
               Exit Sub
         End Select
      End If
      Me.Hide
      DoEvents
      Unload Me
End Sub

Private Sub chkUpdateQueues_Click(Value As Integer)
        bDirty = True
        If chkUpdateQueues.Value Then
            txtUpdateQueues.Enabled = True
            udControl(3).Enabled = True
            lblHeadings(9).Enabled = True
        Else
            txtUpdateQueues.Enabled = False
            udControl(3).Enabled = False
            lblHeadings(9).Enabled = False
        End If
End Sub


Private Sub cmdBrowseDataBase_Click()
      txtDataBase = fFolderBrowser(Me, ApplicationPath, "Select DataBase Location", Me.txtDataBase)
End Sub

Private Sub cmdSettings_Click()
      If fIsReportFormOpen([Report Settings]) Then
         '
      Else
         frmReportSettings.SettingsType = [Screen Settings]
         Load frmReportSettings
         Call FormStayOnTop(frmReportSettings, True, Me)
         frmReportSettings.Show 1, Me
         frmReportSettings.ZOrder
      End If
End Sub

Private Sub cmdThemes_Click()
      Load frmThemes
      Call FormStayOnTop(frmThemes, True, Me)
      frmThemes.Show 1, Me
      Unload frmThemes
      Set frmThemes = Nothing
      FaxSettings.RetreiveIniFileSettings
      Call SetFormStyle(Me, m_cFB)
      Me.Refresh
      Call frmMDI.RefreshStyle
      frmTrace.RefreshStyle
End Sub

Private Sub lstPanel_Click()
Dim iIndex As Integer
      iIndex = lstPanel.ListIndex
Dim i As Integer

      For i = 0 To pnlSettings.UBound
         pnlSettings(i).Visible = False
      Next
      pnlSettings(iIndex).Visible = True
End Sub

Private Sub txtBotQueue_Validate(Cancel As Boolean)
      Cancel = fValidateUpDown(Me, udControl(iudQueueBottom), txtBotQueue)
End Sub

Private Sub txtMaxRetry_Validate(Cancel As Boolean)
      Cancel = fValidateUpDown(Me, udControl(iudMaxRetry), txtMaxRetry)
End Sub

Private Sub txtDialTime_Validate(Cancel As Boolean)
      Cancel = fValidateUpDown(Me, udControl(iudDialTime), txtDialTime)
End Sub

Private Sub txtOffset_Validate(Cancel As Boolean)
      Cancel = fValidateUpDown(Me, udControl(iudOffset), txtOffset)
End Sub

Private Sub txtUpdateQueues_Validate(Cancel As Boolean)
      Cancel = fValidateUpDown(Me, udControl(iudQueueUpdate), txtUpdateQueues)
End Sub

Private Sub cmdAutoDetect_Click()
      cboComPort = fAutoDetect
End Sub

Private Sub cmdBrowseLogos_Click()
      txtLogoPath = fFolderBrowser(Me, ApplicationPath, "Select Logo Path", txtLogoPath)
End Sub

Private Sub cmdBrowseViewer_Click()
      txtViewerPath = fFolderBrowser(Me, ApplicationPath & "\Viewer", "Select Viewer Location", txtViewerPath)
End Sub

Private Sub cmdSaveSettings_Click()
      If fCheckPath(Me, FaxSettings.Paths_ViewerPath) = False Then
         txtViewerPath.SetFocus
         Exit Sub
      End If
        
      If fCheckPath(Me, FaxSettings.Paths_LogoPath) = False Then
         txtLogoPath.SetFocus
         Exit Sub
      End If
        
      Call SaveSettings
End Sub

Private Sub cmdaddPreset_Click()
      bDirty = True
Dim StrTemp As String

      StrTemp = StrTemp & cboFont.Text & ","
      StrTemp = StrTemp & cboSize.Text & ","
      StrTemp = StrTemp & txtOffset & ","
      StrTemp = StrTemp & Str$(chkBold.Value)
      cboPreset.AddItem StrTemp
        
End Sub

Private Sub cboPreset_Click()

      If cboPreset.Text <> "" Then
        
         cboFont.Text = vExtract(cboPreset.Text, 1, ",")
         cboSize.Text = vExtract(cboPreset.Text, 2, ",")
         txtOffset = vExtract(cboPreset.Text, 3, ",")
         chkBold.Value = Val(vExtract(cboPreset.Text, 4, ","))
            
      End If
        
End Sub
Public Sub PopulateSettings()
Dim i As Integer
Dim iPresets As Integer
Dim sPresets() As String

        ' ----------------------------
        ' Populate Controls With Data
        ' ----------------------------
        For i = 1 To 9
            cboComPort.AddItem Str$(i)
        Next i
        cboComPort.ListIndex = 0
        Call GetFonts(cboFont, cboSize)
        cboPreset.Clear
        iPresets = CInt(Val(fGetIni("Preset", "Count", 2)))
        ReDim sPresets(iPresets) As String
        sPresets(1) = "Courier New,23,60, 1"
        sPresets(2) = "Courier New,15,60, 1"

        For i = 1 To iPresets
            cboPreset.AddItem fGetIni("Preset", "Preset" & i, sPresets(i))
        Next
        Call Load_Settings(Me)

End Sub

Public Sub SaveSettings()
Dim i As Integer
        
      Call Save_Settings(Me)
      Call fPutIni("Preset", "Count", cboPreset.ListCount)
      For i = 1 To cboPreset.ListCount
         Call fPutIni("Preset", "Preset" & i, cboPreset.List(i))
      Next
      bDirty = False
      FaxSettings.RetreiveIniFileSettings
      MsgBox "Settings Saved", vbInformation, "enFax Properties"
End Sub

Public Function fAutoDetect()
Dim i As Integer
Dim sClasses As String
Dim iDevices As Integer
Dim sDeviceName As String
Dim xitem   As ListItem

      frmSearching.Show 0, Me
      DoEvents
      frmPrefs.cboComPort.Clear
      iDevices = FaxFinder.DeviceCount
      DoEvents
      If iDevices > 0 Then
         Load frmMultiPorts
         With frmMultiPorts.lstModems
            .ColumnHeaders.Add , , "Port", 4.4
            .ColumnHeaders.Add , , "Device Name", 45
            .ColumnHeaders.Add , , "Class", 10
            .View = lvwReport
            
            .ListItems.Clear
            For i = 0 To iDevices - 1
               cboComPort.AddItem frmPrefs.FaxFinder.Item(i).Port
               sClasses = ""
               If FaxFinder.Item(i).bClass1 Then sClasses = "1"
               If FaxFinder.Item(i).bClass2 Then sClasses = sClasses & ", 2"
               If FaxFinder.Item(i).bClass20 Then sClasses = sClasses & " & 2.0"
               sCommInput = ""
               frmMDI.MSComm1.CommPort = FaxFinder.Item(i).Port
               sDeviceName = fQueryModem("ATI3")
               sDeviceName = Replace(sDeviceName, vbCr, "")
               sDeviceName = Replace(sDeviceName, vbLf, "")
               Set xitem = .ListItems.Add(, , CInt(frmPrefs.FaxFinder.Item(i).Port))
               xitem.SubItems(1) = CStr(sDeviceName)
               xitem.SubItems(2) = CStr(sClasses)
            Next

         End With
         frmSearching.Hide
         Unload frmSearching
         Set frmSearching = Nothing
         DoEvents
         frmMultiPorts.Show 1, Me
         If frmMultiPorts.Tag > -1 Then
            fAutoDetect = frmMultiPorts.Tag
         End If
         Unload frmMultiPorts
         Set frmMultiPorts = Nothing
      Else
         frmSearching.Hide
         Unload frmSearching
         MessageBox ("Unable to locate any fax capable modems" & vbCrLf & vbCrLf & "Please Check hardware and try again")
      End If
End Function

Private Sub txtViewerPath_Change()
      bDirty = True
End Sub

Private Sub txtRings_Change()
      bDirty = True
End Sub

Private Sub txtUpdateQueues_Change()
      bDirty = True
End Sub

Private Sub txtDataBase_Change()
      bDirty = True
End Sub

Private Sub txtDialPrefix_Change()
      bDirty = True
End Sub

Private Sub txtDialTime_Change()
      bDirty = True
End Sub

Private Sub txtLocalID_Change()
      bDirty = True
End Sub

Private Sub txtLogoPath_Change()
      bDirty = True
End Sub

Private Sub txtMaxRetry_Change()
      bDirty = True
End Sub

Private Sub txtOffset_Change()
      bDirty = True
End Sub

Private Sub cboComPort_Change()
      bDirty = True
End Sub

Private Sub cboFont_Change()
      bDirty = True
End Sub

Private Sub cboSize_Change()
      bDirty = True
End Sub

Private Sub chkBold_Click(Value As Integer)
      bDirty = True
End Sub

Private Sub chkDoubleLines_Click(Value As Integer)
      bDirty = True
End Sub

Private Sub chkFaxHistory_Click(Value As Integer)
      bDirty = True
End Sub

Private Sub chkHighRes_Click(Value As Integer)
      bDirty = True
End Sub

Private Sub chkStartPaused_Click(Value As Integer)
      bDirty = True
End Sub

Private Sub chkStretchLogo_Click(Value As Integer)
      bDirty = True
End Sub

Private Sub txtBotQueue_Change()
      bDirty = True
End Sub

Private Sub cboComPort_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(cboComPort, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboComPort_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(cboComPort, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboFont_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(cboFont, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboFont_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(cboFont, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboPreset_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(cboPreset, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboPreset_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(cboPreset, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboSize_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(cboSize, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub cboSize_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(cboSize, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub txtBotQueue_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtBotQueue, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub txtBotQueue_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtBotQueue, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub txtDataBase_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtDataBase, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub txtDataBase_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtDataBase, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub

Private Sub txtDialPrefix_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtDialPrefix, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtDialPrefix_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtDialPrefix, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtDialTime_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtDialTime, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtDialTime_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtDialTime, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtLocalID_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtLocalID, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtLocalID_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtLocalID, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtLogoPath_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtLogoPath, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtLogoPath_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtLogoPath, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtMaxRetry_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtMaxRetry, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtMaxRetry_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtMaxRetry, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtOffset_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtOffset, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtOffset_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtOffset, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtRings_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtRings, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtRings_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtRings, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtUpdateQueues_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtUpdateQueues, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtUpdateQueues_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtUpdateQueues, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtViewerPath_GotFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlGotFocus(txtViewerPath, True)    'set the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub
Private Sub txtViewerPath_LostFocus()
      ' ***-CodeSmart Focus TagStart | Please Do Not Modify
      Call ControlLostFocus(txtViewerPath, True)   'restore the backcolour
      ' ***-CodeSmart Focus TagEnd | Please Do Not Modify
End Sub


Private Sub Form_Load()
Dim i As Integer

      Me.Width = 6360
      Me.Height = 5955
      Call SetFormStyle(Me, m_cFB)
      Call StartupWindowpos(Me)
      Call PopulateSettings
      For i = 0 To pnlSettings.UBound
         pnlSettings(i).TOp = 1740
         pnlSettings(i).Left = 2100
         pnlSettings(i).Width = 3915
         pnlSettings(i).Height = 3465
         pnlSettings(i).BevelWidth = 4
         pnlSettings(i).Visible = False
      Next
      Call FormStayOnTop(Me, True, frmMDI)
      bDirty = False
      lstPanel.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
      Call FormStayOnTop(Me, False)
      Call UnSetFormStyle(Me, m_cFB)
      Call StartupWindowpos(Me, True)
End Sub
