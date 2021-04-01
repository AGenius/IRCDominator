Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Global bSkip As Boolean


Private Sub Main()
Dim ii As Long
Dim i As Long


      ' If Dir(App.Path & "\IRCDominator.ex_") = "" Then
      ' MsgBox "Sorry there appears to be no update to process", vbCritical + vbOKOnly, "IRCDominator Update"
      ' End
      ' End If
      Load frmUpdater
      DoEvents
      frmUpdater.Show
      frmUpdater.PB.Max = 10000
      For i = 1 To 10000
         For ii = 1 To 1000
            DoEvents
         Next
         frmUpdater.PB.Value = i
         DoEvents
         If bSkip Then Exit For
      Next
      On Error Resume Next
      Kill App.Path & "\AutoUpdate.exe"
      MoveFile App.Path & "\AutoUpdate.ex_", App.Path + "\AutoUpdate.exe"
      Kill App.Path & "\IRCDominator.exe"
      MoveFile App.Path & "\IRCDominator.ex_", App.Path + "\IRCDominator.exe"
      ' Directory of File to open and rename   'place to movefile and name of file here
      Unload frmUpdater
      Set frmUpdater = Nothing
      Shell "IRCDominator.exe"
End Sub

Public Sub StayOnTop(hwnd As Long, Stay As Boolean)

      If Stay Then
         SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
      Else
         SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
      End If
End Sub

