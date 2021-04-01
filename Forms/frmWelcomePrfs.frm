VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmWelcomePrefs 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome Message Properties"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWelcomePrfs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWelcomeAway 
      BackColor       =   &H00C0C000&
      Caption         =   "Welcome Back Aways"
      Height          =   285
      Left            =   210
      TabIndex        =   16
      Tag             =   "Welcome|Away"
      Top             =   3540
      Width           =   2025
   End
   Begin VB.CheckBox chkWelcome 
      BackColor       =   &H00C0C000&
      Caption         =   "Welcome Guests"
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Tag             =   "Welcome|Active"
      Top             =   360
      Width           =   1635
   End
   Begin VB.Frame fraWelcome 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   60
      TabIndex        =   2
      Top             =   390
      Width           =   6645
      Begin VB.OptionButton optWelcomeStyle 
         BackColor       =   &H00C0C000&
         Caption         =   "Welcome Message Via Main Screen (All Users Can see)"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Tag             =   "Welcome|RoomMessage"
         Top             =   2760
         Width           =   4875
      End
      Begin VB.OptionButton optWelcomeStyle 
         BackColor       =   &H00C0C000&
         Caption         =   "Message to Guests Own Screen (Only Guest Can See)"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Tag             =   "Welcome|PrivateMessage"
         Top             =   2460
         Value           =   -1  'True
         Width           =   4875
      End
      Begin VB.OptionButton optWelcomeStyle 
         BackColor       =   &H00C0C000&
         Caption         =   "Welcome Message Via Whisper"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Tag             =   "Welcome|Whisper"
         Top             =   2160
         Width           =   4875
      End
      Begin VB.CheckBox chkFontItalic 
         BackColor       =   &H00C0C000&
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4950
         TabIndex        =   8
         Tag             =   "Welcome|FontItalic"
         Top             =   1890
         Width           =   765
      End
      Begin VB.CheckBox chkFontBold 
         BackColor       =   &H00C0C000&
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Tag             =   "Welcome|FontBold"
         Top             =   1890
         Width           =   765
      End
      Begin VB.ComboBox cboFont 
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
         Index           =   0
         Left            =   2910
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Welcome|Font"
         Top             =   1440
         Width           =   2805
      End
      Begin VB.TextBox txtWelcome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Welcome|Message"
         Top             =   540
         Width           =   6105
      End
      Begin MSComctlLib.ImageCombo cboColour 
         Height          =   375
         Index           =   0
         Left            =   930
         TabIndex        =   9
         Tag             =   "Welcome|Colour"
         Top             =   1440
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "Images"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Left            =   180
         TabIndex        =   12
         Top             =   1530
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Left            =   2385
         TabIndex        =   11
         Top             =   1530
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Left            =   3300
         TabIndex        =   10
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblWelcomeMessage 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Welcome Message"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   1455
      End
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
      Left            =   1680
      TabIndex        =   1
      Top             =   6180
      Width           =   1365
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
      Left            =   60
      TabIndex        =   0
      Top             =   6180
      Width           =   1365
   End
   Begin VB.Frame fraAway 
      BackColor       =   &H00C0C000&
      Height          =   2565
      Left            =   60
      TabIndex        =   17
      Top             =   3540
      Width           =   6645
      Begin VB.ComboBox cboFont 
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
         Index           =   1
         Left            =   3030
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Tag             =   "Welcome|AwayFont"
         Top             =   1740
         Width           =   2805
      End
      Begin VB.CheckBox chkFontBold 
         BackColor       =   &H00C0C000&
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   23
         Tag             =   "Welcome|AwayFontBold"
         Top             =   2190
         Width           =   765
      End
      Begin VB.CheckBox chkFontItalic 
         BackColor       =   &H00C0C000&
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4710
         TabIndex        =   22
         Tag             =   "Welcome|AwayFontItalic"
         Top             =   2190
         Width           =   765
      End
      Begin VB.OptionButton optAwayStyle 
         BackColor       =   &H00C0C000&
         Caption         =   "Message to Guests Own Screen (Only Guest Can See)"
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Tag             =   "Welcome|AwayPrivateMessage"
         Top             =   990
         Value           =   -1  'True
         Width           =   4875
      End
      Begin VB.OptionButton optAwayStyle 
         BackColor       =   &H00C0C000&
         Caption         =   "Message Via Main Screen (All Users Can see)"
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   20
         Tag             =   "Welcome|AwayRoomMessage"
         Top             =   1290
         Width           =   4875
      End
      Begin VB.TextBox txtAwayMessage 
         Height          =   360
         Left            =   180
         TabIndex        =   18
         Tag             =   "Welcome|AwayMessage"
         Text            =   "Welcome Back %n."
         Top             =   570
         Width           =   6375
      End
      Begin MSComctlLib.ImageCombo cboColour 
         Height          =   375
         Index           =   1
         Left            =   1050
         TabIndex        =   25
         Tag             =   "Welcome|AwayColour"
         Top             =   1740
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "Images"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   5
         Left            =   3060
         TabIndex        =   28
         Top             =   2220
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   4
         Left            =   2505
         TabIndex        =   27
         Top             =   1830
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   0
         Left            =   300
         TabIndex        =   26
         Top             =   1830
         Width           =   615
      End
      Begin VB.Label lblAwayMessage 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Welcome Back Message"
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   1905
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Enter a %n for nick name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   30
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Enter a %r for Room name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2130
      TabIndex        =   29
      Top             =   90
      Width           =   1995
   End
End
Attribute VB_Name = "frmWelcomePrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkWelcome_Click()
    frmMain.chkWelcome.Value = chkWelcome.Value
    ' ***-CodeSmart Linker TagStart | Please Do Not Modify
    If chkWelcome.Value = 1 Then
        cboColour(0).Enabled = True
        cboFont(0).Enabled = True
        chkFontBold(0).Enabled = True
        txtWelcome.Enabled = True
        fraWelcome.Enabled = True
        chkFontItalic(0).Enabled = True
        optWelcomeStyle(1).Enabled = True
        optWelcomeStyle(2).Enabled = True
        optWelcomeStyle(0).Enabled = True
    Else
        cboColour(0).Enabled = False
        cboFont(0).Enabled = False
        chkFontBold(0).Enabled = False
        txtWelcome.Enabled = False
        fraWelcome.Enabled = False
        chkFontItalic(0).Enabled = False
        optWelcomeStyle(1).Enabled = False
        optWelcomeStyle(2).Enabled = False
        optWelcomeStyle(0).Enabled = False
    End If
    ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub
Private Sub chkWelcomeAway_Click()
    frmMain.chkWelcomeAway.Value = chkWelcomeAway.Value
    ' ***-CodeSmart Linker TagStart | Please Do Not Modify
    If chkWelcomeAway.Value = 1 Then
        chkFontBold(1).Enabled = True
        cboFont(1).Enabled = True
        cboColour(1).Enabled = True
        chkFontItalic(1).Enabled = True
        lblAwayMessage.Enabled = True
        optAwayStyle(0).Enabled = True
        optAwayStyle(1).Enabled = True
        txtAwayMessage.Enabled = True
    Else
        chkFontBold(1).Enabled = False
        cboFont(1).Enabled = False
        cboColour(1).Enabled = False
        chkFontItalic(1).Enabled = False
        lblAwayMessage.Enabled = False
        optAwayStyle(0).Enabled = False
        optAwayStyle(1).Enabled = False
        txtAwayMessage.Enabled = False
    End If
    ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub
Private Sub cmdCancel_Click()
    Me.Hide
End Sub
Private Sub cmdOk_Click()
Dim i As Integer
Dim iFile As Integer
Dim sFilePath As String
     
    Save_Settings Me
    Me.Hide
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim sFilePath As String
Dim iFile As Integer
Dim sTemp As String
Dim ii As Integer

    For ii = 0 To 1
        Call BuildFontList(cboFont(ii))
        Call BuildColourList(Me.cboColour(ii))
    Next
   
    Load_Settings Me
    chkWelcome_Click
    chkWelcomeAway_Click
End Sub
