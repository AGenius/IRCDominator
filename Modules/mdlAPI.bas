Attribute VB_Name = "zmodlAPI"
' ============================================================='
' Module Name       : mdlAPI
' Written By        : Gordon Robinson
' Date              : 08/05/2000
' Comments          :
'
' ============================================================='

Option Explicit

' ============================================================='
' Constants
' ============================================================='

Public Const TTS_ALWAYSTIP = &H1
Public Const TTS_NOPREFIX = &H2

Public Const CW_USEDEFAULT = &H80000000

Public Const WS_POPUP = &H80000000

Public Const WM_SETFONT = &H30
Public Const WM_USER = &H400

Public Const TTM_ADDTOOL = WM_USER + 4
Public Const TTM_SETMAXTIPWIDTH = WM_USER + 24
Public Const TTM_SETDELAYTIME = WM_USER + 3
Public Const TTM_GETDELAYTIME = WM_USER + 21

Public Const TTDT_AUTOMATIC = 0
Public Const TTDT_RESHOW = 1
Public Const TTDT_AUTOPOP = 2
Public Const TTDT_INITIAL = 3

Public Const TTF_SUBCLASS = &H10
Public Const TTF_IDISHWND = &H1
Public Const TTF_CENTERTIP = &H2

Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1

Public Const LOGPIXELSY = 90

Public Const OEM_CHARSET = 255
Public Const EASTEUROPE_CHARSET = 238

Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_IDEFAULTANSICODEPAGE = &H1004  'default ansi code page

Public Const TCI_SRCCODEPAGE = 2
Public Const TCI_SRCCHARSET = 1

' ============================================================='
' Types
' ============================================================='

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type FONTSIGNATURE
    fsUsb(4) As Long
    fsCsb(2) As Long
End Type

Public Type CHARSETINFO
    ciCharset   As Long
    ciACP       As Long
    fs          As FONTSIGNATURE
End Type


' ============================================================='
' API Functions
' ============================================================='

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
    (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    lpParam As Any) _
    As Long

Public Declare Function DestroyWindow Lib "user32" _
    (ByVal hwnd As Long) _
    As Long

Public Declare Function GetClientRect Lib "user32" _
    (ByVal hwnd As Long, _
    lpRect As RECT) _
    As Long
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'    (ByVal hwnd As Long, _
'    ByVal wMsg As Long, _
'    ByVal wParam As Long, _
'    lParam As Any) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) _
    As Long

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
    (ByVal H As Long, _
    ByVal W As Long, _
    ByVal E As Long, _
    ByVal O As Long, _
    ByVal W As Long, _
    ByVal i As Long, _
    ByVal u As Long, _
    ByVal S As Long, _
    ByVal C As Long, _
    ByVal OP As Long, _
    ByVal CP As Long, _
    ByVal Q As Long, _
    ByVal PAF As Long, _
    ByVal F As String) _
    As Long

Public Declare Function MulDiv Lib "kernel32" _
    (ByVal nNumber As Long, _
    ByVal nNumerator As Long, _
    ByVal nDenominator As Long) _
    As Long

Public Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hdc As Long, _
    ByVal nIndex As Long) _
    As Long
    
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
    (ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String, _
    ByVal cchData As Long) _
    As Long
    
Public Declare Function TranslateCharsetInfo Lib "gdi32" _
    (ByVal lpSrc As Long, _
    lpcs As CHARSETINFO, _
    ByVal dwFlags As Long) _
    As Long

Public Declare Function GetSystemDefaultLCID Lib "kernel32" _
    () _
    As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) _
    As Long

