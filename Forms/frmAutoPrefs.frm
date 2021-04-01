VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAutoPrefs 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto List and Kick Preferences"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   14730
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   14730
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3790
      Index           =   1
      Left            =   9750
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   2910
      Width           =   5000
      Begin VB.CheckBox chkListActive 
         BackColor       =   &H00C0C000&
         Caption         =   "Auto Kick Advertising"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Tag             =   "ActiveLists|AdvertisingKicks"
         Top             =   0
         Width           =   2025
      End
      Begin VB.Frame fraList 
         BackColor       =   &H00C0C000&
         Height          =   3645
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   4845
         Begin VB.CommandButton cmdMerge 
            Caption         =   "Merge"
            Height          =   315
            Index           =   4
            Left            =   3510
            TabIndex        =   45
            ToolTipText     =   "STR|Merge words from a file"
            Top             =   2250
            Width           =   975
         End
         Begin VB.OptionButton optNoBan 
            BackColor       =   &H00C0C000&
            Caption         =   "No Ban"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   14
            Tag             =   "ActiveLists|AdvertiseNoBan"
            ToolTipText     =   "STR|Selecting this will only kick"
            Top             =   3135
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optBan 
            BackColor       =   &H00C0C000&
            Caption         =   "Ban For"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   13
            Tag             =   "ActiveLists|AdvertiseBan"
            ToolTipText     =   "STR|Selecting this will ban the guest"
            Top             =   3135
            Width           =   885
         End
         Begin VB.ComboBox cboBans 
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            ItemData        =   "frmAutoPrefs.frx":0ECA
            Left            =   2340
            List            =   "frmAutoPrefs.frx":0EEE
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "ActiveLists|AdvertiseBanTime"
            ToolTipText     =   "STR|Select the ban length"
            Top             =   3075
            Width           =   2145
         End
         Begin VB.ListBox lstNickNames 
            Height          =   2085
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Tag             =   "AdvertisingKicks.dat"
            Top             =   540
            Width           =   2085
         End
         Begin VB.TextBox txtKickingMessage 
            Height          =   1200
            Index           =   2
            Left            =   2340
            MultiLine       =   -1  'True
            TabIndex        =   10
            Tag             =   "ActiveLists|AdvertisingMessage"
            Text            =   "frmAutoPrefs.frx":0F44
            ToolTipText     =   "STR|Enter the Kick message"
            Top             =   540
            Width           =   2385
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2340
            TabIndex        =   9
            ToolTipText     =   "STR|Add A Word to the list"
            Top             =   2700
            Width           =   975
         End
         Begin VB.TextBox txtNickName 
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   2700
            Width           =   2085
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   3510
            TabIndex        =   7
            ToolTipText     =   "STR|Remove a word from the list"
            Top             =   2700
            Width           =   975
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   315
            Index           =   2
            Left            =   2340
            TabIndex        =   6
            ToolTipText     =   "STR|Clear the Words List"
            Top             =   2250
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Kick Message"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   2340
            TabIndex        =   16
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Word List"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   15
            Top             =   300
            Width           =   780
         End
      End
   End
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3790
      Index           =   3
      Left            =   6300
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   2850
      Width           =   5000
      Begin VB.Frame fraFrame 
         BackColor       =   &H00C0C000&
         Height          =   3645
         Index           =   4
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   4845
         Begin VB.CheckBox chkKickCaps 
            BackColor       =   &H00C0C000&
            Caption         =   "Enable Caps Check"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   71
            Tag             =   "ActiveLists|KickCaps"
            ToolTipText     =   "STR|Enable caps test"
            Top             =   0
            Width           =   1860
         End
         Begin VB.OptionButton optCapsKick 
            BackColor       =   &H00C0C000&
            Caption         =   "Give Warning"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   62
            Tag             =   "ActiveLists|CapsWarn"
            ToolTipText     =   "STR|Will just warn the guest"
            Top             =   930
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton optCapsKick 
            BackColor       =   &H00C0C000&
            Caption         =   "Kick For Caps"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   270
            TabIndex        =   61
            Tag             =   "ActiveLists|CapsKick"
            ToolTipText     =   "STR|Will kick the guest"
            Top             =   1950
            Width           =   1425
         End
         Begin MSComctlLib.Slider sldCaps 
            Height          =   510
            Left            =   270
            TabIndex        =   72
            Tag             =   "ActiveLists|CapsTolerance"
            ToolTipText     =   "STR|Select the tolerance level"
            Top             =   330
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   900
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   3
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Frame fraFrame 
            BackColor       =   &H00C0C000&
            Height          =   1005
            Index           =   5
            Left            =   150
            TabIndex        =   63
            Top             =   900
            Width           =   4545
            Begin VB.TextBox txtCapsMessage 
               Height          =   600
               Index           =   0
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   64
               Tag             =   "ActiveLists|CapsMessage"
               Text            =   "frmAutoPrefs.frx":0F5E
               ToolTipText     =   "STR|Enter the Warning message"
               Top             =   315
               Width           =   4335
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C000&
            Height          =   1605
            Left            =   150
            TabIndex        =   65
            Top             =   1920
            Width           =   4545
            Begin VB.OptionButton optBan 
               BackColor       =   &H00C0C000&
               Caption         =   "Ban For"
               Height          =   225
               Index           =   3
               Left            =   1290
               TabIndex        =   69
               Tag             =   "ActiveLists|CapsBan"
               ToolTipText     =   "STR|Selecting this will ban the guest"
               Top             =   375
               Width           =   885
            End
            Begin VB.ComboBox cboBans 
               Enabled         =   0   'False
               Height          =   345
               Index           =   3
               ItemData        =   "frmAutoPrefs.frx":0F8B
               Left            =   2280
               List            =   "frmAutoPrefs.frx":0FAF
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Tag             =   "ActiveLists|CapsBanTime"
               ToolTipText     =   "STR|Select the ban length"
               Top             =   330
               Width           =   2145
            End
            Begin VB.OptionButton optNoBan 
               BackColor       =   &H00C0C000&
               Caption         =   "No Ban"
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   67
               Tag             =   "ActiveLists|CapsNoBan"
               ToolTipText     =   "STR|Selecting this will only kick"
               Top             =   375
               Value           =   -1  'True
               Width           =   885
            End
            Begin VB.TextBox txtCapsMessage 
               Height          =   600
               Index           =   1
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   66
               Tag             =   "ActiveLists|CapsKickMessage"
               Text            =   "frmAutoPrefs.frx":1005
               ToolTipText     =   "STR|Enter the Kick message"
               Top             =   855
               Width           =   4335
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C000&
               Caption         =   "Kick Message"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   6
               Left            =   120
               TabIndex        =   70
               Top             =   600
               Width           =   1080
            End
         End
         Begin VB.Label lblTolCaps 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tolerance"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1950
            TabIndex        =   73
            Top             =   435
            Width           =   1080
         End
      End
   End
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3790
      Index           =   2
      Left            =   7980
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   60
      Width           =   5000
      Begin VB.Frame fraFrame 
         BackColor       =   &H00C0C000&
         Height          =   3645
         Index           =   1
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   4845
         Begin VB.CheckBox chkKickScrolling 
            BackColor       =   &H00C0C000&
            Caption         =   "Enable Scrolling Check"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   57
            Tag             =   "ActiveLists|KickScrolling"
            ToolTipText     =   "STR|Enable scrolling test"
            Top             =   0
            Width           =   2160
         End
         Begin VB.OptionButton optScrollKick 
            BackColor       =   &H00C0C000&
            Caption         =   "Kick For Scrolling"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   270
            TabIndex        =   50
            Tag             =   "ActiveLists|ScrollKick"
            ToolTipText     =   "STR|Will kick the guest"
            Top             =   1950
            Width           =   1785
         End
         Begin VB.OptionButton optScrollKick 
            BackColor       =   &H00C0C000&
            Caption         =   "Give Warning"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   47
            Tag             =   "ActiveLists|ScrollWarn"
            ToolTipText     =   "STR|Will just warn the guest"
            Top             =   930
            Value           =   -1  'True
            Width           =   1365
         End
         Begin MSComctlLib.Slider sldScrolling 
            Height          =   510
            Left            =   270
            TabIndex        =   58
            Tag             =   "ActiveLists|ScrollTolerance"
            ToolTipText     =   "STR|Select the tolerance level"
            Top             =   330
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   900
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   3
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Frame fraFrame 
            BackColor       =   &H00C0C000&
            Height          =   1005
            Index           =   0
            Left            =   150
            TabIndex        =   48
            Top             =   900
            Width           =   4545
            Begin VB.TextBox txtScrollingMessage 
               Height          =   600
               Index           =   0
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   49
               Tag             =   "ActiveLists|ScrollMessage"
               Text            =   "frmAutoPrefs.frx":1033
               ToolTipText     =   "STR|Enter the Warning message"
               Top             =   315
               Width           =   4335
            End
         End
         Begin VB.Frame fraFrame 
            BackColor       =   &H00C0C000&
            Height          =   1605
            Index           =   2
            Left            =   150
            TabIndex        =   51
            Top             =   1920
            Width           =   4545
            Begin VB.TextBox txtScrollingMessage 
               Height          =   600
               Index           =   1
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   55
               Tag             =   "ActiveLists|ScrollKickMessage"
               Text            =   "frmAutoPrefs.frx":1061
               ToolTipText     =   "STR|Enter the Kick message"
               Top             =   855
               Width           =   4335
            End
            Begin VB.OptionButton optNoBan 
               BackColor       =   &H00C0C000&
               Caption         =   "No Ban"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   54
               Tag             =   "ActiveLists|ScrollNoBan"
               ToolTipText     =   "STR|Selecting this will only kick"
               Top             =   375
               Value           =   -1  'True
               Width           =   885
            End
            Begin VB.ComboBox cboBans 
               Enabled         =   0   'False
               Height          =   345
               Index           =   2
               ItemData        =   "frmAutoPrefs.frx":1086
               Left            =   2280
               List            =   "frmAutoPrefs.frx":10AA
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Tag             =   "ActiveLists|ScrollBanTime"
               ToolTipText     =   "STR|Select the ban length"
               Top             =   330
               Width           =   2145
            End
            Begin VB.OptionButton optBan 
               BackColor       =   &H00C0C000&
               Caption         =   "Ban For"
               Height          =   225
               Index           =   2
               Left            =   1290
               TabIndex        =   52
               Tag             =   "ActiveLists|ScrollBan"
               ToolTipText     =   "STR|Selecting this will ban the guest"
               Top             =   375
               Width           =   885
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C000&
               Caption         =   "Kick Message"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   120
               TabIndex        =   56
               Top             =   600
               Width           =   1080
            End
         End
         Begin VB.Label lblTolScrolling 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tolerance"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1950
            TabIndex        =   59
            Top             =   435
            Width           =   1080
         End
      End
   End
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3790
      Index           =   0
      Left            =   4020
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   17
      Top             =   1980
      Width           =   5000
      Begin VB.CheckBox chkListActive 
         BackColor       =   &H00C0C000&
         Caption         =   "Auto Kick for Profanity"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Tag             =   "ActiveLists|ProfanityKicks"
         Top             =   0
         Width           =   2265
      End
      Begin VB.Frame fraList 
         BackColor       =   &H00C0C000&
         Height          =   3645
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   4845
         Begin VB.ComboBox cboBans 
            Enabled         =   0   'False
            Height          =   345
            Index           =   0
            ItemData        =   "frmAutoPrefs.frx":1100
            Left            =   2340
            List            =   "frmAutoPrefs.frx":1124
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "ActiveLists|ProfanityBanTime"
            ToolTipText     =   "STR|Select the ban length"
            Top             =   3075
            Width           =   2145
         End
         Begin VB.OptionButton optBan 
            BackColor       =   &H00C0C000&
            Caption         =   "Ban For"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   28
            Tag             =   "ActiveLists|ProfanityBan"
            ToolTipText     =   "STR|Selecting this will ban the guest"
            Top             =   3135
            Width           =   885
         End
         Begin VB.OptionButton optNoBan 
            BackColor       =   &H00C0C000&
            Caption         =   "No Ban"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   27
            Tag             =   "ActiveLists|ProfanityNoBan"
            ToolTipText     =   "STR|Selecting this will only kick"
            Top             =   3135
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.CommandButton cmdMerge 
            Caption         =   "Merge"
            Height          =   315
            Index           =   3
            Left            =   3510
            TabIndex        =   26
            ToolTipText     =   "STR|Merge words from a file"
            Top             =   2250
            Width           =   975
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   315
            Index           =   1
            Left            =   2340
            TabIndex        =   25
            ToolTipText     =   "STR|Clear the Words List"
            Top             =   2250
            Width           =   975
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   3510
            TabIndex        =   24
            ToolTipText     =   "STR|Remove a word from the list"
            Top             =   2700
            Width           =   975
         End
         Begin VB.TextBox txtNickName 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   2700
            Width           =   2085
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2340
            TabIndex        =   22
            ToolTipText     =   "STR|Add A Word to the list"
            Top             =   2700
            Width           =   975
         End
         Begin VB.TextBox txtKickingMessage 
            Height          =   1200
            Index           =   1
            Left            =   2340
            MultiLine       =   -1  'True
            TabIndex        =   21
            Tag             =   "ActiveLists|ProfanityMessage"
            Text            =   "frmAutoPrefs.frx":117A
            ToolTipText     =   "STR|Enter the Kick message"
            Top             =   540
            Width           =   2385
         End
         Begin VB.ListBox lstNickNames 
            Height          =   2085
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Tag             =   "ProfanityKicks.dat"
            Top             =   540
            Width           =   2085
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Word List"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   31
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C000&
            Caption         =   "Kick Message"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2340
            TabIndex        =   30
            Top             =   270
            Width           =   1080
         End
      End
   End
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
      Left            =   120
      TabIndex        =   43
      Top             =   6480
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
      Left            =   1650
      TabIndex        =   42
      Top             =   6480
      Width           =   1365
   End
   Begin VB.CheckBox chkListActive 
      BackColor       =   &H00C0C000&
      Caption         =   "Auto Kick List"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   690
      TabIndex        =   33
      Tag             =   "ActiveLists|Kicks"
      Top             =   0
      Width           =   1455
   End
   Begin VB.ListBox lstPanel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      ItemData        =   "frmAutoPrefs.frx":1193
      Left            =   3930
      List            =   "frmAutoPrefs.frx":11A3
      TabIndex        =   32
      ToolTipText     =   "STR|Click an Option to change"
      Top             =   390
      Width           =   3015
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00C0C000&
      Height          =   6255
      Index           =   0
      Left            =   60
      TabIndex        =   34
      Top             =   30
      Width           =   3765
      Begin VB.ListBox lstNickNames 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3120
         Index           =   0
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   40
         Tag             =   "Kicks.dat"
         Top             =   240
         Width           =   3585
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   705
         Index           =   0
         Left            =   120
         Picture         =   "frmAutoPrefs.frx":11ED
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "STR|Add a user to the list"
         Top             =   3870
         Width           =   885
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   705
         Index           =   0
         Left            =   1410
         Picture         =   "frmAutoPrefs.frx":1AB7
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "STR|Remove a user from the list"
         Top             =   3870
         Width           =   885
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   705
         Index           =   0
         Left            =   2760
         Picture         =   "frmAutoPrefs.frx":2381
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "STR|Clear the users from the list"
         Top             =   3870
         Width           =   885
      End
      Begin VB.TextBox txtNickName 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   3420
         Width           =   3585
      End
      Begin VB.TextBox txtKickingMessage 
         Height          =   900
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Tag             =   "ActiveLists|KicksMessage"
         Text            =   "frmAutoPrefs.frx":2C4B
         Top             =   5280
         Width           =   3555
      End
      Begin VB.Label lblListed 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   74
         Top             =   4650
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Kick Message"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   5010
         Width           =   1080
      End
   End
   Begin VB.PictureBox pctSettings 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   3790
      Index           =   4
      Left            =   4980
      ScaleHeight     =   3795
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   6090
      Width           =   5000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Select Option"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3930
      TabIndex        =   44
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frmAutoPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Dim sNames(1) As String
Dim sTols(2) As String
Const iProfanity As Integer = 0
Const iAdvertise As Integer = 1
Const iScrolling As Integer = 2
Const iCaps As Integer = 3
Private m_objTooltip    As cTooltip

Public Sub chkKickCaps_Click()

      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If chkKickCaps.Value = 1 Then
         cboBans(3).Enabled = True
         optBan(3).Enabled = True
         optNoBan(3).Enabled = True
         sldScrolling.Enabled = True
         Me.txtCapsMessage(0).Enabled = True
         Me.txtCapsMessage(1).Enabled = True
         Me.optCapsKick(0).Enabled = True
         Me.optCapsKick(1).Enabled = True
         If Me.optNoBan(3).Value Then
            Me.optNoBan_Click (3)
         Else
            Me.optBan_Click (3)
         End If
         If Me.optCapsKick(0).Value Then
            Me.optCapsKick_Click (0)
         Else
            Me.optCapsKick_Click (1)
         End If
      Else
         cboBans(3).Enabled = False
         optBan(3).Enabled = False
         optNoBan(3).Enabled = False
         sldScrolling.Enabled = False
         Me.txtCapsMessage(0).Enabled = False
         Me.txtCapsMessage(1).Enabled = False
         Me.optCapsKick(0).Enabled = False
         Me.optCapsKick(1).Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub

Public Sub chkKickScrolling_Click()
      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If chkKickScrolling.Value = 1 Then
         cboBans(2).Enabled = True
         optBan(2).Enabled = True
         optNoBan(2).Enabled = True
         sldScrolling.Enabled = True
         Me.txtScrollingMessage(0).Enabled = True
         Me.txtScrollingMessage(1).Enabled = True
         Me.optScrollKick(0).Enabled = True
         Me.optScrollKick(1).Enabled = True
         If Me.optNoBan(2).Value Then
            Me.optNoBan_Click (2)
         Else
            Me.optBan_Click (2)
         End If
         If Me.optScrollKick(0).Value Then
            Me.optScrollKick_Click (0)
         Else
            Me.optScrollKick_Click (1)
         End If
      Else
         cboBans(2).Enabled = False
         optBan(2).Enabled = False
         optNoBan(2).Enabled = False
         sldScrolling.Enabled = False
         Me.txtScrollingMessage(0).Enabled = False
         Me.txtScrollingMessage(1).Enabled = False
         Me.optScrollKick(0).Enabled = False
         Me.optScrollKick(1).Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub



Public Sub chkListActive_Click(Index As Integer)
      On Error Resume Next
      If chkListActive(Index).Value Then
         fraList(Index).Enabled = True
         cmdClear(Index).Enabled = True
         txtNickName(Index).Enabled = True
         lstNickNames(Index).Enabled = True
         Me.txtKickingMessage(Index).Enabled = True
         cmdMerge(Index).Enabled = True
         Me.optBan(Index - 3).Enabled = True
         Me.optNoBan(Index - 3).Enabled = True
      Else
         fraList(Index).Enabled = False
         cmdAdd(Index).Enabled = False
         cmdRemove(Index).Enabled = False
         cmdClear(Index).Enabled = False
         txtNickName(Index).Enabled = False
         lstNickNames(Index).Enabled = False
         Me.txtKickingMessage(Index).Enabled = False
         cmdMerge(Index).Enabled = False
         Me.optBan(Index - 3).Enabled = False
         Me.optNoBan(Index - 3).Enabled = False
      End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
      If FindInList(lstNickNames(Index), Trim(txtNickName(Index))) > -1 Then
         Call MsgBox("Already in List", vbOKOnly, "Add User")
      Else
         lstNickNames(Index).AddItem Trim(txtNickName(Index))
         txtNickName(Index) = ""
      End If
      lblListed.Caption = lstNickNames(0).ListCount & " Names in list"

End Sub

Private Sub cmdCancel_Click()
      Me.Hide
End Sub

Private Sub cmdClear_Click(Index As Integer)
      If lstNickNames(Index).ListCount > 0 Then
         If MsgBox("Are You Sure", vbYesNo, "Clear NickNames List") = vbYes Then
            lstNickNames(Index).Clear
            txtNickName(Index).Text = ""
         End If
      End If
      lblListed.Caption = lstNickNames(0).ListCount & " Names in list"

End Sub

Private Sub cmdOk_Click()
      Call SaveSettings
      Me.Hide
End Sub
Private Sub cmdRemove_Click(Index As Integer)
      If lstNickNames(Index).ListIndex > -1 Then
         lstNickNames(Index).RemoveItem lstNickNames(Index).ListIndex
         txtNickName(Index).Text = ""
         cmdAdd(Index).Enabled = False
         cmdRemove(Index).Enabled = False
      End If
      lblListed.Caption = lstNickNames(0).ListCount & " Names in list"

End Sub
Private Sub Form_Load()
Dim i As Integer
Dim iFile As Integer
Dim sFilePath As String
Dim l As Integer
Dim sTemp As String

      LoadToolTips Me, m_objTooltip
      For i = 0 To cboBans.UBound
         Call BuildBanList(cboBans(i))
      Next
      ' Load_Settings Me
      Me.Width = 9675
      sTols(2) = "Very Tolerent (10+)"
      sTols(1) = "Tolerent (5+)"
      sTols(0) = "Intolerent (2+)"

      ' sNames(0) = "Owners.dat"
      ' sNames(1) = "Hosts.dat"
      
      For i = 0 To cmdAdd.UBound
         cmdAdd(i).Enabled = False
         cmdRemove(i).Enabled = False
      Next
      
      ' For l = 0 To UBound(sNames)
      ' Call LoadList(lstNickNames(l), lstNickNames(l).Tag)
      ' Next
      
      Call PopulateSettings
      ' if lstnicknames(0).ListCount >
      lblListed.Caption = lstNickNames(0).ListCount & " Names in list"
      
      TestCheck
      For i = 0 To pctSettings.UBound
         pctSettings(i).BackColor = Me.BackColor
         If i > 0 Then
            pctSettings(i).Top = pctSettings(0).Top
            pctSettings(i).Left = pctSettings(0).Left
         End If
         pctSettings(i).Visible = False
      Next
      Me.lstPanel.ListIndex = 0
      Me.lblTolScrolling.Caption = sTols(Me.sldScrolling.Value - 1)
      Me.lblTolCaps.Caption = sTols(Me.sldCaps.Value - 1)
      Me.chkKickScrolling_Click
      Me.chkKickCaps_Click
      For i = 0 To chkListActive.UBound
         Me.chkListActive_Click (i)
      Next
End Sub
Public Sub PopulateSettings()

      Me.cboBans(iProfanity).ListIndex = KickSettings.Profanity_BanTime
      Me.cboBans(iScrolling).ListIndex = KickSettings.Scroll_BanTime
      Me.cboBans(iAdvertise).ListIndex = KickSettings.Advertise_BanTime
      Me.cboBans(iCaps).ListIndex = KickSettings.Caps_BanTime
      Me.chkKickCaps.Value = Abs(KickSettings.Caps_Active)
      Me.chkKickScrolling.Value = Abs(KickSettings.Scroll_Active)
      
      Me.chkListActive(0).Value = Abs(KickSettings.KickList_Active)
      Me.chkListActive(1).Value = Abs(KickSettings.Profanity_Active)
      Me.chkListActive(2).Value = Abs(KickSettings.Advertise_Active)
      Me.txtCapsMessage(0) = KickSettings.Caps_Message
      Me.txtCapsMessage(1) = KickSettings.Caps_KickMessage
      Me.txtKickingMessage(0) = KickSettings.KickList_Message
      Me.txtKickingMessage(1) = KickSettings.Profanity_KickMessage
      Me.txtKickingMessage(2) = KickSettings.Advertise_KickMessage
      Me.txtScrollingMessage(0) = KickSettings.Scroll_Message
      Me.txtScrollingMessage(1) = KickSettings.Scroll_KickMessage
      Me.optBan(iProfanity).Value = KickSettings.Profanity_Ban
      Me.optBan(iScrolling).Value = KickSettings.Scroll_Ban
      Me.optBan(iAdvertise).Value = KickSettings.Advertise_Ban
      Me.optBan(iCaps).Value = KickSettings.Caps_Ban
      Me.optNoBan(iProfanity).Value = KickSettings.Profanity_NoBan
      Me.optNoBan(iScrolling).Value = KickSettings.Scroll_NoBan
      Me.optNoBan(iAdvertise).Value = KickSettings.Advertise_NoBan
      Me.optNoBan(iCaps).Value = KickSettings.Caps_NoBan
      Me.optCapsKick(0).Value = KickSettings.Caps_Warning
      Me.optCapsKick(1).Value = KickSettings.Caps_Kick
      Me.optScrollKick(0).Value = KickSettings.Scroll_Warning
      Me.optScrollKick(1).Value = KickSettings.Scroll_Kick
      Me.sldCaps.Value = KickSettings.Caps_Tolerance
      Me.sldScrolling.Value = KickSettings.Scroll_Tolerance
      Call KickSettings.FillControl(lstNickNames(0), lKickList)
      Call KickSettings.FillControl(lstNickNames(1), lProfanity)
      Call KickSettings.FillControl(lstNickNames(2), lAdvert)
End Sub
Private Sub SaveSettings()
      KickSettings.Profanity_BanTime = Me.cboBans(iProfanity).ListIndex
      KickSettings.Scroll_BanTime = Me.cboBans(iScrolling).ListIndex
      KickSettings.Advertise_BanTime = Me.cboBans(iAdvertise).ListIndex
      KickSettings.Caps_BanTime = Me.cboBans(iCaps).ListIndex
      KickSettings.Caps_Active = Me.chkKickCaps.Value
      KickSettings.Scroll_Active = Me.chkKickScrolling.Value
      KickSettings.KickList_Active = Me.chkListActive(0).Value
      KickSettings.Profanity_Active = Me.chkListActive(1).Value
      KickSettings.Advertise_Active = Me.chkListActive(2).Value
      KickSettings.Caps_Message = Me.txtCapsMessage(0)
      KickSettings.Caps_KickMessage = Me.txtCapsMessage(1)
      KickSettings.KickList_Message = Me.txtKickingMessage(0)
      KickSettings.Profanity_KickMessage = Me.txtKickingMessage(1)
      KickSettings.Advertise_KickMessage = Me.txtKickingMessage(2)
      KickSettings.Scroll_Message = Me.txtScrollingMessage(0)
      KickSettings.Scroll_KickMessage = Me.txtScrollingMessage(1)
      KickSettings.Profanity_Ban = Me.optBan(iProfanity).Value
      KickSettings.Scroll_Ban = Me.optBan(iScrolling).Value
      KickSettings.Advertise_Ban = Me.optBan(iAdvertise).Value
      KickSettings.Caps_Ban = Me.optBan(iCaps).Value
      KickSettings.Profanity_NoBan = Me.optNoBan(iProfanity).Value
      KickSettings.Scroll_NoBan = Me.optNoBan(iScrolling).Value
      KickSettings.Advertise_NoBan = Me.optNoBan(iAdvertise).Value
      KickSettings.Caps_NoBan = Me.optNoBan(iCaps).Value
      KickSettings.Caps_Warning = Me.optCapsKick(0).Value
      KickSettings.Caps_Kick = Me.optCapsKick(1).Value
      KickSettings.Scroll_Warning = Me.optScrollKick(0).Value
      KickSettings.Scroll_Kick = Me.optScrollKick(1).Value
      KickSettings.Caps_Tolerance = Me.sldCaps.Value
      KickSettings.Scroll_Tolerance = Me.sldScrolling.Value
      KickSettings.SaveControl lstNickNames(0), lKickList
      KickSettings.SaveControl lstNickNames(1), lProfanity
      KickSettings.SaveControl lstNickNames(2), lAdvert
      KickSettings.SavePrefs
End Sub

Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
End Sub

Private Sub lstNickNames_Click(Index As Integer)
      If lstNickNames(Index).ListIndex <> -1 Then
         txtNickName(Index) = lstNickNames(Index).List(lstNickNames(Index).ListIndex)
         cmdRemove(Index).Enabled = True
         cmdAdd(Index).Enabled = False
      End If
End Sub
Private Sub lstPanel_Click()
Dim iIndex As Integer
Dim i As Integer

      iIndex = lstPanel.ListIndex

      For i = 0 To pctSettings.UBound
         pctSettings(i).Visible = False
      Next
      pctSettings(iIndex).Visible = True
End Sub
Public Sub optBan_Click(Index As Integer)
      If Me.optBan(Index).Value Then
         Me.cboBans(Index).Enabled = True
      End If
End Sub
Public Sub optCapsKick_Click(Index As Integer)
      If Index = 1 Then
         Me.optBan(3).Enabled = True
         Me.optNoBan(3).Enabled = True
         Me.cboBans(3).Enabled = True
         Me.txtCapsMessage(1).Enabled = True
         Me.txtCapsMessage(0).Enabled = False
         If Me.optNoBan(3).Value Then
            Me.optNoBan_Click (3)
         Else
            Me.optBan_Click (3)
         End If
      Else
         Me.optBan(3).Enabled = False
         Me.optNoBan(3).Enabled = False
         Me.cboBans(3).Enabled = False
         Me.txtCapsMessage(1).Enabled = False
         Me.txtCapsMessage(0).Enabled = True
      End If
End Sub
Public Sub optNoBan_Click(Index As Integer)
      If Me.optNoBan(Index).Value Then
         Me.cboBans(Index).Enabled = False
      End If
End Sub
Public Sub optScrollKick_Click(Index As Integer)
      If Index = 1 Then
         Me.optBan(2).Enabled = True
         Me.optNoBan(2).Enabled = True
         Me.cboBans(2).Enabled = True
         Me.txtScrollingMessage(1).Enabled = True
         Me.txtScrollingMessage(0).Enabled = False
         If Me.optNoBan(2).Value Then
            Me.optNoBan_Click (2)
         Else
            Me.optBan_Click (2)
         End If
      Else
         Me.optBan(2).Enabled = False
         Me.optNoBan(2).Enabled = False
         Me.cboBans(2).Enabled = False
         Me.txtScrollingMessage(1).Enabled = False
         Me.txtScrollingMessage(0).Enabled = True
      End If
End Sub
Private Sub sldCaps_Scroll()
      Me.lblTolCaps.Caption = sTols(Me.sldCaps.Value - 1)
End Sub
Private Sub sldScrolling_Scroll()
      Me.lblTolScrolling.Caption = sTols(Me.sldScrolling.Value - 1)
End Sub

Private Sub txtNickName_Change(Index As Integer)
      If txtNickName(Index) <> "" Then
         cmdAdd(Index).Enabled = True
         cmdRemove(Index).Enabled = False
      Else
         cmdAdd(Index).Enabled = False
         cmdRemove(Index).Enabled = False
      End If
End Sub
Private Sub TestCheck()
Dim i As Integer
      For i = 0 To 2
         If chkListActive(i).Value Then
            fraList(i).Enabled = True
            cmdClear(i).Enabled = True
            txtNickName(i).Enabled = True
            lstNickNames(i).Enabled = True
         Else
            fraList(i).Enabled = False
            cmdAdd(i).Enabled = False
            cmdRemove(i).Enabled = False
            cmdClear(i).Enabled = False
            txtNickName(i).Enabled = False
            lstNickNames(i).Enabled = False
         End If
      Next
End Sub
