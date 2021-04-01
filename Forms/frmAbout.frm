VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enigma Ware's - IRCDominator"
   ClientHeight    =   7155
   ClientLeft      =   5880
   ClientTop       =   3450
   ClientWidth     =   5790
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4938.51
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.107
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShouts 
      Caption         =   "Shouts"
      Height          =   345
      Left            =   270
      TabIndex        =   18
      Top             =   6690
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysInfo 
      Cancel          =   -1  'True
      Caption         =   "&System Info.."
      Default         =   -1  'True
      Height          =   345
      Left            =   3090
      TabIndex        =   14
      Top             =   6690
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4350
      TabIndex        =   2
      Top             =   6690
      Width           =   1260
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   2
      Left            =   1050
      TabIndex        =   17
      Top             =   5310
      Width           =   4545
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H008080FF&
      BorderWidth     =   16
      Height          =   945
      Index           =   1
      Left            =   210
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
      Left            =   570
      TabIndex        =   5
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
      Left            =   1380
      TabIndex        =   10
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
      Left            =   4500
      TabIndex        =   16
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By Åß§ølµ†€•G€ñïµ§ 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   4560
      Width           =   3915
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1020
      Picture         =   "frmAbout.frx":0A16
      Top             =   4380
      Width           =   720
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
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   2865
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1080
      TabIndex        =   13
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   957.832
      X2              =   5253.992
      Y1              =   2339.839
      Y2              =   2339.839
   End
   Begin VB.Image imgIcon 
      Height          =   630
      Left            =   360
      Picture         =   "frmAbout.frx":18E0
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
      Left            =   2670
      TabIndex        =   8
      Top             =   1020
      Width           =   2265
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001 Absolute Genius"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1050
      MouseIcon       =   "frmAbout.frx":21AA
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Tag             =   "http://www.mrenigma.co.uk"
      Top             =   3930
      Width           =   3120
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.mrenigma.co.uk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1050
      MouseIcon       =   "frmAbout.frx":24B4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "http://www.mrenigma.co.uk"
      Top             =   4110
      Width           =   3870
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Enigma Wares Home Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "frmAbout.frx":27BE
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "http://www.mrenigma.co.uk"
      Top             =   3630
      Width           =   2355
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":2AC8
      ForeColor       =   &H00000000&
      Height          =   690
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Width           =   4260
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "At Last a real chat room manager that will really keep your chat room in order. Auto kicks for swearing - scrolling and links."
      ForeColor       =   &H00000000&
      Height          =   720
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   1860
      Width           =   4260
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FF8080&
      Height          =   225
      Index           =   4
      Left            =   1080
      TabIndex        =   12
      Top             =   540
      Width           =   4065
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   780
      Index           =   3
      Left            =   1080
      TabIndex        =   11
      Top             =   840
      Width           =   4065
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
      Left            =   3690
      TabIndex        =   4
      Top             =   60
      Width           =   2235
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   2
      Left            =   4860
      Top             =   180
      Width           =   555
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   0
      Left            =   4860
      Top             =   1200
      Width           =   555
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objTooltip    As cTooltip

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
      KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
      KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


'Private Sub cmdSubscribe_Click()
'      Subscribe Me.hwnd
'End Sub

Private Sub cmdSysInfo_Click()
      Call StartSysInfo
End Sub

Private Sub cmdOk_Click()
      Unload Me
End Sub



Private Sub Form_DblClick()
Dim sAns As String
      sAns = InputBox("Hello Are you Absolute Genius", "Owner ???")
      If sAns = "lalala" Then
         frmUnlock.sPassword = "GeniusIsGOD"
         Load frmUnlock
         If Not (bActivated) Then
            frmUnlock.Show 1
         End If
         If frmUnlock.Tag = "UNLOCKED" Then
            bActivated = True
            ' Unlock Owner stuff here
            UnlockMe
         End If
         Unload frmUnlock

      Else
         MsgBox "Oh now your not", vbCritical + vbOKOnly, "HAHA"
      End If
End Sub

Private Sub Form_Load()
      LoadToolTips Me, m_objTooltip
      Me.Caption = "About " & App.Title
      Me.Icon = frmMain.Icon
      lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Sub StartSysInfo()
      On Error GoTo SysInfoErr
  
Dim rc As Long
Dim SysInfoPath As String
    
      ' Try To Get System Info Program Path\Name From Registry...
      If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
         ' Try To Get System Info Program Path Only From Registry...
      ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
         ' Validate Existance Of Known 32 Bit File Version
         If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
         Else
            GoTo SysInfoErr
         End If
         ' Error - Registry Entry Can Not Be Found...
      Else
         GoTo SysInfoErr
      End If
    
      Call Shell(SysInfoPath, vbNormalFocus)
    
      Exit Sub
SysInfoErr:
      MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
Dim i As Long                                           ' Loop Counter
Dim rc As Long                                          ' Return Code
Dim hKey As Long                                        ' Handle To An Open Registry Key
Dim hDepth As Long                                      '
Dim KeyValType As Long                                  ' Data Type Of A Registry Key
Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
      ' ------------------------------------------------------------
      ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
      ' ------------------------------------------------------------
      rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
      If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
      tmpVal = String$(1024, 0)                             ' Allocate Variable Space
      KeyValSize = 1024                                       ' Mark Variable Size
    
      ' ------------------------------------------------------------
      ' Retrieve Registry Key Value...
      ' ------------------------------------------------------------
      rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
      KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
      If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
      If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
      tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
   Else                                                    ' WinNT Does NOT Null Terminate String...
      tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
   End If
   ' ------------------------------------------------------------
   ' Determine Key Value Type For Conversion...
   ' ------------------------------------------------------------
   Select Case KeyValType                                  ' Search Data Types...
      Case REG_SZ                                             ' String Registry Key Data Type
         KeyVal = tmpVal                                     ' Copy String Value
      Case REG_DWORD                                          ' Double Word Registry Key Data Type
         For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
         Next
         KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
   End Select
    
   GetKeyValue = True                                      ' Return Success
   rc = RegCloseKey(hKey)                                  ' Close Registry Key
   Exit Function                                           ' Exit
    
GetKeyError: ' Cleanup After An Error Has Occured...
   KeyVal = ""                                             ' Set Return Val To Empty String
   GetKeyValue = False                                     ' Return Failure
   rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Paint()
      ' Dim sCompany As String
      ' Dim sName As String
      ' Dim sSerial As String

      ' GetRegInfo sCompany, sName, sSerial
      ' lblRegCompany = sCompany
      ' lblRegName = sName
      ' lblRegSerial = sSerial
End Sub

Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
End Sub

Private Sub Image1_DblClick()
'      frmUnlock.sPassword = "SpecialAccess"
'      Load frmUnlock
'      If Not (bSpecialActivated) Then
'         frmUnlock.Show 1
'      End If
'      If frmUnlock.Tag = "UNLOCKED" Then
'         bSpecialActivated = True
'         ' Unlock Secret stuff here
'         UnlockSpecial
'      End If
'      Unload frmUnlock
End Sub

Private Sub imgIcon_DblClick()
'      frmUnlock.sPassword = "SpecialSpecial"
'      Load frmUnlock
'      If Not (bSpecialSpecialActivated) Then
'         frmUnlock.Show 1
'      End If
'      If frmUnlock.Tag = "UNLOCKED" Then
'         bSpecialSpecialActivated = True
'         frmControl.chkClone.Visible = True
'         frmControl.cmdGetGold.Visible = True
'         UnlockSpecial
'      End If
'      Unload frmUnlock

End Sub

Private Sub lblCopyright_Click()
      pShell lblCopyright.Tag, Me
End Sub

Private Sub lblProduct_Click()
      pShell lblProduct.Tag, Me
End Sub

Private Sub lblURL_Click()
      pShell lblURL.Tag, Me
End Sub
