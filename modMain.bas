Attribute VB_Name = "modMain"
Option Explicit

Global IniFile As New cIniFile
Global IRCIniFile As New cIniFile
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const WM_USER = &H400
Public Const ACM_OPEN = WM_USER + 100&
Public Function InIDE() As Boolean
      On Error GoTo InIDEError
      InIDE = False
      Debug.Print 1 / 0
      Exit Function

InIDEError:
      InIDE = True
      Exit Function
End Function

Function fGetIni(sSection As String, sKeyWord As String, sDefault As String) As Variant
      With IniFile
         .Section = sSection
         .Key = sKeyWord
         .Default = sDefault
         fGetIni = .Value
      End With
End Function

Public Sub PutIni(sSection As String, sKeyWord As String, sValue As String)
      With IniFile
         .Section = sSection
         .Key = sKeyWord
         .Value = sValue
      End With
End Sub

Function Locate(sData As String, sFind As String) As Long
      Locate = InStr(1, sData, sFind, 1)
End Function

Public Sub LoadResAVI(pCtrlAnimation As Animation, iResid As Integer)
      SendMessage pCtrlAnimation.hwnd, ACM_OPEN, ByVal App.hInstance, ByVal iResid
End Sub
Public Sub ClearAnim(pCtrlAnimation As Animation)

      On Error Resume Next
    
      ' clear previous animation
      With pCtrlAnimation
         .AutoPlay = False
         .Close
         .AutoPlay = True
      End With
End Sub

Sub Main()
      IniFile.Path = App.Path & "\" & App.EXEName & ".dat"
      IRCIniFile.Path = App.Path & "\IRCDominator.dat"
      frmUpdate.Show
End Sub

Public Sub StayOnTop(hwnd As Long, Stay As Boolean)

      If Stay Then
         SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
      Else
         SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
      End If
End Sub

