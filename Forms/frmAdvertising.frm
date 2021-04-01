VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAdvertising 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advertising Preferences"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdvertising.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   9270
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2850
      TabIndex        =   16
      Top             =   8400
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
      Left            =   1320
      TabIndex        =   15
      Top             =   8400
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   5430
      Width           =   9195
      Begin VB.CheckBox chkAdvertItalic 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Italic"
         Height          =   255
         Index           =   2
         Left            =   7800
         TabIndex        =   36
         Tag             =   "Advertise|3Italic"
         Top             =   1710
         Width           =   765
      End
      Begin VB.CheckBox chkAdvertBold 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bold"
         Height          =   255
         Index           =   2
         Left            =   6930
         TabIndex        =   35
         Tag             =   "Advertise|3Bold"
         Top             =   1710
         Width           =   765
      End
      Begin VB.ComboBox cboAdvertFont 
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
         Index           =   2
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "Advertise|3Font"
         Top             =   1290
         Width           =   2805
      End
      Begin VB.CheckBox chkAction 
         BackColor       =   &H00C0E0FF&
         Caption         =   "As Action"
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
         Index           =   2
         Left            =   4410
         TabIndex        =   19
         Tag             =   "Advertise|3Action"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtAdvertMessage 
         Height          =   1200
         Index           =   2
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   12
         Tag             =   "Advertise|3Message"
         Top             =   390
         Width           =   5625
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Advertise This"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Tag             =   "Advertise|3"
         Top             =   0
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldAdvertise 
         Height          =   510
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Tag             =   "Advertise|3Interval"
         Top             =   1680
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.ImageCombo cboAdvertColour 
         Height          =   360
         Index           =   2
         Left            =   6120
         TabIndex        =   37
         Tag             =   "Advertise|3Colour"
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         Index           =   8
         Left            =   6120
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   7
         Left            =   6120
         TabIndex        =   39
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   6
         Left            =   6120
         TabIndex        =   38
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Tag             =   "0"
         Top             =   2190
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2700
      Width           =   9195
      Begin VB.CheckBox chkAdvertItalic 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Italic"
         Height          =   255
         Index           =   1
         Left            =   7800
         TabIndex        =   29
         Tag             =   "Advertise|2Italic"
         Top             =   1770
         Width           =   765
      End
      Begin VB.CheckBox chkAdvertBold 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bold"
         Height          =   255
         Index           =   1
         Left            =   6930
         TabIndex        =   28
         Tag             =   "Advertise|2Bold"
         Top             =   1770
         Width           =   765
      End
      Begin VB.ComboBox cboAdvertFont 
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
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "Advertise|2Font"
         Top             =   1350
         Width           =   2805
      End
      Begin VB.CheckBox chkAction 
         BackColor       =   &H00C0E0FF&
         Caption         =   "As Action"
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
         Index           =   1
         Left            =   4410
         TabIndex        =   18
         Tag             =   "Advertise|2Action"
         Top             =   2250
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Advertise This"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Tag             =   "Advertise|2"
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtAdvertMessage 
         Height          =   1200
         Index           =   1
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   7
         Tag             =   "Advertise|2Message"
         Top             =   390
         Width           =   5625
      End
      Begin MSComctlLib.Slider sldAdvertise 
         Height          =   510
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Tag             =   "Advertise|2Interval"
         Top             =   1680
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.ImageCombo cboAdvertColour 
         Height          =   360
         Index           =   1
         Left            =   6120
         TabIndex        =   30
         Tag             =   "Advertise|2Colour"
         Top             =   660
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         Index           =   5
         Left            =   6120
         TabIndex        =   33
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   4
         Left            =   6120
         TabIndex        =   32
         Top             =   1110
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   3
         Left            =   6120
         TabIndex        =   31
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2190
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.CheckBox chkAdvertItalic 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Italic"
         Height          =   255
         Index           =   0
         Left            =   7800
         TabIndex        =   22
         Tag             =   "Advertise|1Italic"
         Top             =   1770
         Width           =   765
      End
      Begin VB.CheckBox chkAdvertBold 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bold"
         Height          =   255
         Index           =   0
         Left            =   6930
         TabIndex        =   21
         Tag             =   "Advertise|1Bold"
         Top             =   1770
         Width           =   765
      End
      Begin VB.ComboBox cboAdvertFont 
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
         Left            =   6120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "Advertise|1Font"
         Top             =   1350
         Width           =   2805
      End
      Begin MSComctlLib.ImageCombo cboAdvertColour 
         Height          =   360
         Index           =   0
         Left            =   6120
         TabIndex        =   23
         Tag             =   "Advertise|1Colour"
         Top             =   660
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.CheckBox chkAction 
         BackColor       =   &H00C0E0FF&
         Caption         =   "As Action"
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
         Left            =   4410
         TabIndex        =   17
         Tag             =   "Advertise|1Action"
         Top             =   2250
         Width           =   1335
      End
      Begin MSComctlLib.Slider sldAdvertise 
         Height          =   510
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Tag             =   "Advertise|1Interval"
         Top             =   1680
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
      Begin VB.TextBox txtAdvertMessage 
         Height          =   1200
         Index           =   0
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   2
         Tag             =   "Advertise|1Message"
         Top             =   390
         Width           =   5625
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Advertise This"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Tag             =   "Advertise|1"
         Top             =   0
         Width           =   1455
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
         Index           =   0
         Left            =   6120
         TabIndex        =   26
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Index           =   1
         Left            =   6120
         TabIndex        =   25
         Top             =   1110
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
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
         Left            =   6120
         TabIndex        =   24
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Tag             =   "0"
         Top             =   2190
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAdvertising"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Check1_Click(Index As Integer)
      ' ***-CodeSmart Linker TagStart | Please Do Not Modify
      If Check1(Index).Value = 1 Then
         txtAdvertMessage(Index).Enabled = True
         sldAdvertise(Index).Enabled = True
         cboAdvertColour(Index).Enabled = True
         cboAdvertFont(Index).Enabled = True
         chkAdvertBold(Index).Enabled = True
         chkAdvertItalic(Index).Enabled = True
      Else
         txtAdvertMessage(Index).Enabled = False
         sldAdvertise(Index).Enabled = False
         cboAdvertColour(Index).Enabled = False
         cboAdvertFont(Index).Enabled = False
         chkAdvertBold(Index).Enabled = False
         chkAdvertItalic(Index).Enabled = False
      End If
      ' ***-CodeSmart Linker TagEnd | Please Do Not Modify
End Sub
Private Sub cmdCancel_Click()
      Me.Hide
End Sub
Private Sub cmdOk_Click()
      Call Save_Settings(Me)
      Me.Hide
End Sub
Private Sub Form_Load()
Dim i As Integer
On Error GoTo 0

      For i = 0 To Check1.UBound
         Call BuildFontList(Me.cboAdvertFont(i))
         Call BuildColourList(Me.cboAdvertColour(i))
      Next
      Load_Settings Me
      For i = 0 To Check1.UBound
         Check1_Click (i)
         Me.sldAdvertise_Scroll (i)
      Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      If UnloadMode = 0 Then
         Me.Hide
         Cancel = True
      End If
End Sub
Public Sub sldAdvertise_Scroll(Index As Integer)
      Me.lblInterval(Index) = "Interval - " & Me.sldAdvertise(Index).Value & " Minutes"
End Sub
