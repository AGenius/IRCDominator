VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Sys Tray Interface"
   ClientHeight    =   1920
   ClientLeft      =   5610
   ClientTop       =   3360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Begin VB.Menu mnuSystray 
         Caption         =   "&Open IRCDominator"
         Index           =   0
         Tag             =   "OPEN"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Preferences"
         Index           =   2
         Begin VB.Menu mnuSystray1 
            Caption         =   "General"
            Index           =   3
            Tag             =   "PrefsGen"
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "Passport"
            Enabled         =   0   'False
            Index           =   4
            Tag             =   "PrefsPassport"
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "Welcome"
            Index           =   5
            Tag             =   "PrefsWelcome"
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "Auto Kick"
            Index           =   6
            Tag             =   "PrefsKick"
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "Auto Host / Owner"
            Index           =   7
            Tag             =   "PrefsHost"
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "Trace"
            Index           =   8
            Tag             =   "PrefsTrace"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuSystray1 
            Caption         =   "Advertising"
            Index           =   10
            Tag             =   "PrefsAdvert"
         End
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "&Enigma Wares Home Page"
         Index           =   11
         Tag             =   "WEB"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "&About..."
         Index           =   12
         Tag             =   "ABOUT"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "&Shutdown IRCDominator"
         Index           =   14
         Tag             =   "SHUTDOWN"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' frmSysTray.
' Steve McMahon
' based on code supplied from Ben Baird:

' Author:
' Ben Baird <psyborg@cyberhighway.com>
' Copyright (c) 1997, Ben Baird
'
' Purpose:
' Demonstrates setting an icon in the taskbar's
' system tray without the overhead of subclassing
' to receive events.

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private m_objTooltip    As cTooltip

Private Const WM_MOUSEMOVE = &H200
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const MAX_TOOLTIP As Integer = 64
Private Type NOTIFYICONDATA
      cbSize As Long
      hwnd As Long
      uID As Long
      uFlags As Long
      uCallbackMessage As Long
      hIcon As Long
      szTip As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)

Private m_bAddedMenuItem As Boolean
Private m_iDefaultIndex As Long

Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long

      sTip = nfIconData.szTip
      iPos = InStr(sTip, Chr$(0))
      If (iPos <> 0) Then
         sTip = Left$(sTip, iPos - 1)
      End If
      ToolTip = sTip
End Property
Public Property Let ToolTip(ByVal sTip As String)
      If (sTip & Chr$(0) <> nfIconData.szTip) Then
         nfIconData.szTip = sTip & Chr$(0)
         nfIconData.uFlags = NIF_TIP
         Shell_NotifyIcon NIM_MODIFY, nfIconData
      End If
End Property
Public Property Get IconHandle() As Long
      IconHandle = nfIconData.hIcon
End Property
Public Property Let IconHandle(ByVal hIcon As Long)
      If (hIcon <> nfIconData.hIcon) Then
         nfIconData.hIcon = hIcon
         nfIconData.uFlags = NIF_ICON
         Shell_NotifyIcon NIM_MODIFY, nfIconData
      End If
End Property
Public Function AddMenuItem(ByVal sCaption As String, Optional ByVal sKey As String = "", Optional ByVal bDefault As Boolean = False) As Long
Dim iIndex As Long
      If Not (m_bAddedMenuItem) Then
         iIndex = 0
         m_bAddedMenuItem = True
      Else
         iIndex = mnuSystray.UBound + 1
         Load mnuSystray(iIndex)
      End If
      mnuSystray(iIndex).Visible = True
      mnuSystray(iIndex).Tag = sKey
      mnuSystray(iIndex).Caption = sCaption
      If (bDefault) Then
         m_iDefaultIndex = iIndex
      End If
      AddMenuItem = iIndex
End Function
Private Function ValidIndex(ByVal lIndex As Long) As Boolean
      ValidIndex = (lIndex >= mnuSystray.LBound And lIndex <= mnuSystray.UBound)
End Function
Public Sub EnableMenuItem(ByVal lIndex As Long, ByVal bState As Boolean)
      If (ValidIndex(lIndex)) Then
         mnuSystray(lIndex).Enabled = bState
      End If
End Sub
Public Function RemoveMenuItem(ByVal iIndex As Long) As Long
Dim i As Long
      If ValidIndex(iIndex) Then
         If (iIndex = 0) Then
            mnuSystray(0).Caption = ""
         Else
            ' remove the item:
            For i = iIndex + 1 To mnuSystray.UBound
               mnuSystray(iIndex - 1).Caption = mnuSystray(iIndex).Caption
               mnuSystray(iIndex - 1).Tag = mnuSystray(iIndex).Tag
            Next i
            Unload mnuSystray(mnuSystray.UBound)
         End If
      End If
End Function
Public Property Get DefaultMenuIndex() As Long
      DefaultMenuIndex = m_iDefaultIndex
End Property
Public Property Let DefaultMenuIndex(ByVal lIndex As Long)
      If (ValidIndex(lIndex)) Then
         m_iDefaultIndex = lIndex
      Else
         m_iDefaultIndex = 0
      End If
End Property
Public Function ShowMenu()
      SetForegroundWindow Me.hwnd
      If (m_iDefaultIndex > -1) Then
         Me.PopupMenu mnuPopup, 0, , , mnuSystray(m_iDefaultIndex)
      Else
         Me.PopupMenu mnuPopup, 0
      End If
End Function

Private Sub Form_Load()
      LoadToolTips Me, m_objTooltip
      ' Add the icon to the system tray...
      With nfIconData
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.Handle
         .szTip = App.FileDescription & Chr$(0)
         .cbSize = Len(nfIconData)
      End With
      Shell_NotifyIcon NIM_ADD, nfIconData
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lX As Long
      ' VB manipulates the x value according to scale mode:
      ' we must remove this before we can interpret the
      ' message windows was trying to send to us:
      lX = ScaleX(X, Me.ScaleMode, vbPixels)
      Select Case lX
         Case WM_MOUSEMOVE
            RaiseEvent SysTrayMouseMove
         Case WM_LBUTTONUP
            RaiseEvent SysTrayMouseDown(vbLeftButton)
         Case WM_LBUTTONUP
            RaiseEvent SysTrayMouseUp(vbLeftButton)
         Case WM_LBUTTONDBLCLK
            RaiseEvent SysTrayDoubleClick(vbLeftButton)
         Case WM_RBUTTONDOWN
            RaiseEvent SysTrayMouseDown(vbRightButton)
         Case WM_RBUTTONUP
            RaiseEvent SysTrayMouseUp(vbRightButton)
         Case WM_RBUTTONDBLCLK
            RaiseEvent SysTrayDoubleClick(vbRightButton)
      End Select

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub
Private Sub Form_Unload(Cancel As Integer)
      m_objTooltip.Destroy
End Sub
Private Sub mnuSysTray_Click(Index As Integer)
      RaiseEvent MenuClick(Index, mnuSystray(Index).Tag)
End Sub
Private Sub mnuSystray1_Click(Index As Integer)
      RaiseEvent MenuClick(Index, mnuSystray1(Index).Tag)
End Sub
