Attribute VB_Name = "modDeclares"
Option Explicit

' Global Const bOwner As Boolean = True
Global Const bOwner As Boolean = False

Public Const AMT_CAPS_TOLERATED As Byte = 5
Global bShowSysTray     As Boolean

Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SendMessageBynum Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TYPE_TEXTMETRIC) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwnd As Long, ByVal nFolder As Long, Pidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (Pidl As Long, ByVal FolderPath As String) As Long
    


Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400
Public Const WM_NOTIFY = &H4E
Public Const WM_LBUTTONDOWN = &H201
Public Const EM_GETEVENTMASK = WM_USER + 59
Public Const EM_GETTEXTRANGE = WM_USER + 75
Public Const ACM_OPEN = WM_USER + 100&
Public Const EM_SETEVENTMASK = WM_USER + 69
Public Const EM_AUTOURLDETECT = WM_USER + 91
Public Const EN_LINK = &H70B
Public Const ENM_LINK = &H4000000
Public Const SW_SHOWNORMAL = 1

Type tagNMHDR
   hwndFrom As Long
   idFrom   As Long
   code     As Long
End Type

Type CHARRANGE
   cpMin As Long
   cpMax As Long
End Type

Type ENLINK
   nmhdr  As tagNMHDR
   msg    As Long
   wParam As Long
   lParam As Long
   chrg   As CHARRANGE
End Type

Type TEXTRANGE
   chrg      As CHARRANGE
   lpstrText As Long
End Type

Public Const SPI_GETWORKAREA = 48
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETRECT = &HB2
Public Const WM_GETFONT = &H31
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB

Global iCounter As Integer
Public Enum eCols
   CBlack = 0
   CWhite = 15
   CMaroon = 4
   CGreen = 2
   CNavy = 1
   COlive = 6
   CPurple = 5
   CTeal = 3
   CSilver = 7
   CGray = 8
   CRed = 12
   CLime = 10
   CBlue = 9
   CYellow = 14
   CFuschia = 13
   CAqua = 11
End Enum

Type TYPE_TEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
End Type

Global Const sRoot As String = "SOFTWARE\Classes\CLSID"
Global Const sPassRoot As String = "SOFTWARE\MICROSOFT\MSNCHAT"
Global Const sSubKeyV3 As String = "{81361155-FAF9-11d3-B0D3-00C04F612FF1}"
Global Const sSubKeyV4 As String = "{29c13b62-b9f7-4cd3-8cef-0a58a1a99441}"
Global Const sKey1  As String = "{E113C6A6-D44A-4639-A40E-3B6DE32A1A40}"
Public Const sCols  As String = "0,15,4,2,1,6,5,3,7,8,12,10,9,14,13,11"
Public Enum ePrefs
   pChat = 0
   pWhisper = 1
   pWelcome = 2
   pAway = 3
   pAdvert1 = 4
   pAdvert2 = 5
   pAdvert3 = 6
End Enum
Public Enum eSystemFolder
   fld_Desktop = 0
   fld_StartMenu_Programs = 2
   fld_My_Documents = 5
   fld_Favorites = 6
   fld_Startup = 7
   fld_Recent = 8
   fld_SentTo = 9
   fld_Start_Menu = 11
   fld_Windows_Desktop = 16
   fld_Network_Neighborhood = 19
   fld_Fonts = 20
   fld_ShellNew = 21
   fld_StartMenu = 22
   fld_StartMenuPrograms = 23
   fld_StartMenuStartUp = 24
   fld_AllUsers_Desktop = 25
   fld_DefApplicationData = 26
   fld_Printhood = 27
   fld_ApplicationData = 28
   fld_Favourites = 31
   fld_DefTemporaryInternetFiles = 32
   fld_DefCookies = 33
   fld_DefHistory = 34
   fld_AllUserApplicationData = 35
   fld_Windows = 36
   fld_SystemFolder = 37
   fld_ProgramFiles = 38
   fld_MyPictures = 39
   fld_DefSettingsRoot = 40
   fld_SystemFolder2 = 41
   fld_ProgramFilesCommon = 43
   fld_AllUsersTemplates = 45
   fld_AllUsersDocuments = 46
   fld_AllUsersStartProgramsAdminTools = 47
End Enum

Public bError           As Boolean
Public iPort            As Integer
Public bConnecting      As Boolean
Public bStarted         As Boolean
Global sWaitedFor       As String
Global sNamesList       As String
Global bConnected       As Boolean
Global sRoomMode        As String
Global bRefresh         As Boolean
Global sRoomJoined      As String
Global sNickJoined      As String
Global Whispers()       As New frmWhisper
Global IniFile          As New cIniFile
Global sIncommingData   As String
Global bJoined          As Boolean
Global bTryJoin         As Boolean
Global bAdvanced        As Boolean
Global StatusColour     As Integer
Global bActivated       As Boolean
Global bSpecialActivated As Boolean
Global bSpecialSpecialActivated As Boolean
Global sEnteredPass     As String
Global bScollingChat    As Boolean
Global bScollingTrace   As Boolean
Global bLoadingApp      As Boolean
Global iAliveTimer      As Integer
Global iAdvertTimer     As Integer
Global bAdvertise       As Boolean
Global Const sChatCaption As String = "Main Chat Window"
Global bWaiting As Boolean
Global sWaitKeyWords As String
Global Const sInts As String = "1,2,5,15,60,1440,2880,0"
Global Const sBans As String = "1 Minute,2 Minutes,5 Minutes,15 Minutes,1 Hour,24 Hours,48 Hours,Indefinite"
Global bShutDown  As Boolean


Global GeneralSettings As New clsPrefsGeneral
Global KickSettings As New clsPrefsKicks
Global HostLists As New clsPrefsLists
Global UsersList As New clsUserList
Public glnglpOriginalWndProc As Long
Public glngOriginalhWnd As Long
Global iAuthCount As Integer

Public Sub Main()
Dim sCode As String
Dim sKey As String
Dim sVal As String
Dim iLoop As Long

      On Error Resume Next

100   sKey = Chr(46) & Mid$(sAlphaBet, 19, 1) & Mid$(sAlphaBet, 8, 1) & Mid$(sAlphaBet, 20, 1) & Mid$(sAlphaBet, 13, 1) & Mid$(sAlphaBet, 12, 1)
101   sVal = Chr(104) & Chr(116) & Chr(109) & Chr(108)
102   sCode = ReadRegistry(HKEY_CLASSES_ROOT, sKey, "Content Type")
103   If sCode = sVal Then
104      If UnlockApp = False Then
105         End
         End If
      End If

'108   If Dir(App.Path & "\msnchat30.ocx") = "" Then
'         sKey = "I cannot find the msnchat30.ocx file in the IRCDominator Installed Directory"
'         sKey = sKey & vbCrLf & "Please follow the instructions provided on my web site on how to fix this problem"
'         sKey = sKey & vbCrLf & vbCrLf & "If this problem is not dealt with you may find that you experience problems"
'         sKey = sKey & vbCrLf & "with the guest list.."
'         MsgBox sKey, vbCritical + vbOKOnly, "IRCDominator OCX Check"
'
'      End If

      ' Shell "regsvr32.exe " & App.Path & "\msnchat30.ocx /s"
      ' If Dir(App.Path & "\var.dll") = "" Then
      ' Shell "regsvr32.exe /s " & App.Path & "\var.dll"
      ' MsgBox "I have just updated some files - please re-start IRCDominator", vbCritical, "IRCDominator Update"
      ' End
      ' End If



117   Load frmCheckLatest
118   DoEvents
120   bLoadingApp = True
121   Load frmSplash
122   DoEvents
123   frmSplash.Show
124   DoEvents
119   frmCheckLatest.Check
    
125   IniFile.Path = App.Path & "\" & App.EXEName & ".dat"

      IniFile.Section = "AutoUpdater"
      IniFile.Key = "Version"
      If IniFile.Value <> "1.6.1" Then

         ExtractUpdater
         IniFile.Value = "1.6.1"

      End If

126   GeneralSettings.LoadPrefs
127   KickSettings.LoadPrefs
128   HostLists.LoadPrefs
129   Load MDIMain
130   frmSplash.Hide
131   bLoadingApp = False
132   DoEvents
133   Unload frmSplash
134   MDIMain.Show
135   iAliveTimer = 0
      If GeneralSettings.ShowMOTD = True Then
         frmMOTD.Show 1
      End If
End Sub

Private Sub ExtractUpdater()
Dim iFileNumber As Integer
Dim DllBuffer() As Byte

      DllBuffer = LoadResData(101, "CUSTOM")
      iFileNumber = FreeFile
      Kill App.Path & "\AutoUpdate.exe"
      Open App.Path & "\AutoUpdater.exe" For Binary Access Write As #iFileNumber
      Put #iFileNumber, , DllBuffer
      Close #iFileNumber

End Sub

' Private Sub CreateDll()
' Dim iFileNumber As Integer
' Dim DllBuffer() As Byte
'
' DllBuffer = LoadResData(102, "CUSTOM")
' iFileNumber = FreeFile
' Open App.Path & "\var.dll" For Binary Access Write As #iFileNumber
' Put #iFileNumber, , DllBuffer
' Close #iFileNumber
'
' End Sub

Public Function TestLink(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim udtNMHDR               As tagNMHDR
Dim udtENLINK              As ENLINK
Dim udtTEXTRANGE           As TEXTRANGE
Dim strBuffer              As String * 128
Dim strOperation           As String
Dim strFileName            As String
Dim strDefaultDirectory    As String
Dim tt As Long
Dim T  As Long


      RtlMoveMemory udtNMHDR, ByVal lParam, Len(udtNMHDR)
      If udtNMHDR.hwndFrom = frmMain.txtChat.hwnd And udtNMHDR.code = EN_LINK Then
         RtlMoveMemory udtENLINK, ByVal lParam, Len(udtENLINK)
         If udtENLINK.msg = WM_LBUTTONDOWN Then
            strBuffer = ""
            With udtTEXTRANGE
               .chrg.cpMin = udtENLINK.chrg.cpMin
               .chrg.cpMax = udtENLINK.chrg.cpMax
               .lpstrText = StrPtr(strBuffer)
            End With

            With frmMain.txtChat
               T = SendMessage(.hwnd, EM_GETTEXTRANGE, 0, udtTEXTRANGE)
            End With

            RtlMoveMemory ByVal strBuffer, ByVal udtTEXTRANGE.lpstrText, Len(strBuffer)
            strOperation = "open"
            strFileName = strBuffer
            tt = ShellExecuteForExplore(frmMain.hwnd, strOperation, strFileName, vbNullString, strDefaultDirectory, SW_SHOWNORMAL)

         End If
      End If
      '
      ' RichTextBoxSubProc = CallWindowProc(glnglpOriginalWndProc, hwnd, uMsg, wParam, lParam)

End Function

Public Sub Testing()
Dim i As Integer
      For i = 20 To 255
         ' Debug.Print PP.ConvertedString(0, 1, Chr(I), 0);
      Next
End Sub
