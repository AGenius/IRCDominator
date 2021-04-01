Attribute VB_Name = "modBits"
Option Explicit

Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public Sub FlashTitle(ByRef oForm As Form)
        FlashWindow oForm.hwnd, 1
End Sub

