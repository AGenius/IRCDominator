Attribute VB_Name = "modDecode"
Option Explicit
Global PP As New MSNChatProtocolCtl


Public Function ConvertFromUTF(ByVal sText As String) As String
On Error Resume Next

      ConvertFromUTF = sText
      On Error Resume Next
      ConvertFromUTF = PP.ConvertedString(1, 0, sText, 0)
End Function


'
' Public Function ConvertNickname(ByVal sNickName As String) As String
' Dim EndOfNick As Integer
' Dim ConvertUTF8 As Integer
' Dim onechar As String
' Dim EachChar(255) As String
' Dim Counter As Integer
' Dim Counter2 As Integer
' Dim ConvFromUtf8(255) As String
' Dim sFromString As String
' Dim sTooString As String
'
'
' sFromString = ""
' sTooString = ""
'
' Counter = 0
'
' EndOfNick = Len(sNickName)
'
' Do While Counter <> EndOfNick
' Counter = Counter + 1
' EachChar(Counter) = Mid$(sNickName, Counter, 1)
' Loop
'
' Counter = 1
' Counter2 = 1
'
' Do While Counter <> EndOfNick + 1
'
' Select Case EachChar(Counter)
' Case "�"
' Select Case EachChar(Counter + 1)
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case Else
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case "�"
' Select Case EachChar(Counter + 1)
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case Else
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case "�"
' Select Case EachChar(Counter + 1)
' Case "�"
' Select Case EachChar(Counter + 2)
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case "�"
' Select Case EachChar(Counter + 2)
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case "�"
' Select Case EachChar(Counter + 1)
' Case "�"
' Select Case EachChar(Counter + 2)
' Case "�"
' ConvFromUtf8(Counter2) = "(": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = ")": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = ":": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "*": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "%": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = ".": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "/": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = ";": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = " ": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "'": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "?": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "<": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "=": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = ">": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case "�"
' Select Case EachChar(Counter + 2)
' Case "�"
' ConvFromUtf8(Counter2) = "[": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "]": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "^": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "\": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "�"
' ConvFromUtf8(Counter2) = "@": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = "�": Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Loop
' FinishDecoding:
'
' EndOfNick = Counter2
'
' For ConvertUTF8 = 1 To EndOfNick
' ConvertNickname = ConvertNickname + ConvFromUtf8(ConvertUTF8)
' Next ConvertUTF8
' End Function
