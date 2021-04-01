VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum ttIconType
   TTNoIcon = 0
   TTIconInfo = 1
   TTIconWarning = 2
   TTIconError = 3
End Enum

Public Enum ttStyleEnum
   TTStandard = 0
   TTBalloon = 1
End Enum

Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hwnd        As Long
    uId         As Long
    cRect       As RECT
    hinst       As Long
    lpszText    As String
End Type


Private m_lngHwnd          As Long
Private m_lngMaxWidth      As Long
Private m_hFont            As Long
Private mvarIcon           As ttIconType
Private mvarCentered       As Boolean
Private mvarStyle          As ttStyleEnum
Private mvarTitle          As String
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"
Private ti As TOOLINFO

' ====================================================================='
' Properties
' ====================================================================='
Public Property Let Title(ByVal vData As String)
   mvarTitle = vData
   If m_lngHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
      SendMessage m_lngHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
End Property
Public Property Get Title() As String
   Title = ti.lpszText
End Property

Public Property Let Icon(ByVal vData As ttIconType)
   mvarIcon = vData
   If m_lngHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
      SendMessage m_lngHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
End Property
Public Property Get Icon() As ttIconType
   Icon = mvarIcon
End Property
Public Property Let TStyle(ByVal vData As ttStyleEnum)
   mvarStyle = vData
End Property
Public Property Get TStyle() As ttStyleEnum
   TStyle = mvarStyle
End Property
Public Property Let Centered(ByVal vData As Boolean)
   mvarCentered = vData
End Property
Public Property Get Centered() As Boolean
   Centered = mvarCentered
End Property
Public Property Get MaxWidth() As Long
   MaxWidth = m_lngMaxWidth
End Property
Public Property Let MaxWidth(lngMaxWidth As Long)
   m_lngMaxWidth = lngMaxWidth
   SendMessageLong m_lngHwnd, TTM_SETMAXTIPWIDTH, 0, m_lngMaxWidth
End Property
Public Property Get VisibleTime() As Long
   VisibleTime = SendMessageLong(m_lngHwnd, TTM_GETDELAYTIME, TTDT_AUTOPOP, 0)
End Property
Public Property Let VisibleTime(lngTime As Long)
   If lngTime > 32767 Then lngTime = 32767
   If lngTime < 0 Then lngTime = 0
   SendMessageLong m_lngHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, lngTime
End Property
Public Property Get DelayTime() As Long
   DelayTime = SendMessageLong(m_lngHwnd, TTM_GETDELAYTIME, TTDT_INITIAL, 0)
End Property
Public Property Let DelayTime(lngTime As Long)
   If lngTime > 32767 Then lngTime = 32767
   If lngTime < 0 Then lngTime = 0
   SendMessageLong m_lngHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, lngTime
End Property
Public Property Set Font(objFont As StdFont)
Dim hFont       As Long
Dim lHeight     As Long
Dim lItalic     As Long
Dim lBold       As Long
Dim lUnderline  As Long
Dim lStrikeOut  As Long
Dim ReturnVal   As Long
Dim WndRect     As RECT
Dim sCodePage   As String * 6
Dim lCodePage   As Long
Dim cSetInfo    As CHARSETINFO
Dim lRetVal     As Long
Dim hOldFont    As Long
Dim lPixelsY    As Long
    
   If objFont Is Nothing Then Exit Property
   lHeight = -MulDiv(objFont.Size, GetDeviceCaps(GetDC(m_lngHwnd), LOGPIXELSY), 72)
   If objFont.Italic = True Then
      lItalic = 1
   Else
      lItalic = 0
   End If
   If objFont.Bold = True Then
      lBold = 1
   Else
      lBold = 0
   End If
   If objFont.Underline = True Then
      lUnderline = 1
   Else
      lUnderline = 0
   End If
   If objFont.Strikethrough = True Then
      lStrikeOut = 1
   Else
      lStrikeOut = 0
   End If

' Determine the character set for the current locale
   GetLocaleInfo GetSystemDefaultLCID(), LOCALE_IDEFAULTANSICODEPAGE, sCodePage, 6
   sCodePage = Left$(sCodePage, InStr(sCodePage, Chr$(0)) - 1)
   lCodePage = CLng(sCodePage)
   lRetVal = TranslateCharsetInfo(lCodePage, cSetInfo, TCI_SRCCODEPAGE)
   hFont = CreateFont(lHeight, 0, 0, 0, objFont.Weight, lItalic, lUnderline, lStrikeOut, cSetInfo.ciCharset, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH, objFont.Name)
   ReturnVal = SendMessageLong(m_lngHwnd, WM_SETFONT, hFont, 1&)

   If m_hFont <> 0 Then
      DeleteObject m_hFont
   End If
   m_hFont = hFont
End Property

' ====================================================================='
' Methods
' ====================================================================='
Public Sub Destroy()
      DestroyWindow m_lngHwnd
      If m_hFont <> 0 Then
         DeleteObject m_hFont
      End If
End Sub
Public Sub Create(lngHwndParent As Long)
Dim lWinStyle As Long
    
      lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
      ' 'create baloon style if desired
      If mvarStyle = TTBalloon Then
         lWinStyle = lWinStyle Or TTS_BALLOON
      End If
      m_hFont = 0
      ' m_lngHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, TTS_NOPREFIX Or TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, lngHwndParent, 0, App.hInstance, 0)
      m_lngHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, lngHwndParent, 0&, App.hInstance, 0&)
      ' 'make our tooltip window a topmost window
      SetWindowPos m_lngHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
      ' 'if we want a title or we want an icon
      If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
         SendMessage m_lngHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
      End If
        
      ' If mvarForeColor <> Empty Then
      ' SendMessage lHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
      ' End If
      '
      ' If mvarBackColor <> Empty Then
      ' SendMessage lHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
      ' End If

      SendMessageLong m_lngHwnd, TTM_SETMAXTIPWIDTH, 0, m_lngMaxWidth
End Sub
Public Sub AddControl(ctlTool As Object, strCaption As String, Optional blnCenterTip As Boolean = False)
Dim udtToolInfo As TOOLINFO
   
      With udtToolInfo
         GetClientRect ctlTool.hwnd, .cRect
         .hwnd = ctlTool.hwnd
         If mvarCentered Then
            .uFlags = TTF_IDISHWND Or TTF_SUBCLASS Or TTF_CENTERTIP
         Else
            .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
         End If
         ' If blnCenterTip Then
         ' .uFlags = .uFlags Or TTF_CENTERTIP
         ' End If
         .uId = ctlTool.hwnd
         .lpszText = strCaption
         .cbSize = Len(udtToolInfo)
      End With
      SendMessage m_lngHwnd, TTM_ADDTOOL, 0, udtToolInfo
End Sub

' ====================================================================='
' Events
' ====================================================================='
Private Sub Class_Initialize()
      InitCommonControls
      m_lngMaxWidth = 300
End Sub
