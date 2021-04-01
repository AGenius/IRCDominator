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
' Case "√"
' Select Case EachChar(Counter + 1)
' Case "•"
' ConvFromUtf8(Counter2) = "Â": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "°"
' ConvFromUtf8(Counter2) = "‡": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¢"
' ConvFromUtf8(Counter2) = "‚": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "£"
' ConvFromUtf8(Counter2) = "„": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "§"
' ConvFromUtf8(Counter2) = "‰": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¶"
' ConvFromUtf8(Counter2) = "Ê": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ß"
' ConvFromUtf8(Counter2) = "Á": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "®"
' ConvFromUtf8(Counter2) = "Ë": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "©"
' ConvFromUtf8(Counter2) = "È": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "™"
' ConvFromUtf8(Counter2) = "Í": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "´"
' ConvFromUtf8(Counter2) = "Î": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¨"
' ConvFromUtf8(Counter2) = "Ï": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "≠"
' ConvFromUtf8(Counter2) = "Ì": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Æ"
' ConvFromUtf8(Counter2) = "Ì": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ø"
' ConvFromUtf8(Counter2) = "Ô": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∞"
' ConvFromUtf8(Counter2) = "": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "±"
' ConvFromUtf8(Counter2) = "Ò": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "≤"
' ConvFromUtf8(Counter2) = "Ú": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "≥"
' ConvFromUtf8(Counter2) = "Û": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¥"
' ConvFromUtf8(Counter2) = "Ù": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "µ"
' ConvFromUtf8(Counter2) = "ı": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∂"
' ConvFromUtf8(Counter2) = "ˆ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∑"
' ConvFromUtf8(Counter2) = "˜": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∏"
' ConvFromUtf8(Counter2) = "¯": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "π"
' ConvFromUtf8(Counter2) = "˘": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∫"
' ConvFromUtf8(Counter2) = "˙": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ª"
' ConvFromUtf8(Counter2) = "˚": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "º"
' ConvFromUtf8(Counter2) = "¸": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ω"
' ConvFromUtf8(Counter2) = "˝": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "æ"
' ConvFromUtf8(Counter2) = "˛": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ø"
' ConvFromUtf8(Counter2) = "ˇ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ä"
' ConvFromUtf8(Counter2) = "¿": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Å"
' ConvFromUtf8(Counter2) = "¡": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ç"
' ConvFromUtf8(Counter2) = "¬": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "É"
' ConvFromUtf8(Counter2) = "√": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ñ"
' ConvFromUtf8(Counter2) = "ƒ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ö"
' ConvFromUtf8(Counter2) = "≈": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ü"
' ConvFromUtf8(Counter2) = "∆": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "á"
' ConvFromUtf8(Counter2) = "«": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "à"
' ConvFromUtf8(Counter2) = "»": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "â"
' ConvFromUtf8(Counter2) = "…": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ä"
' ConvFromUtf8(Counter2) = " ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ã"
' ConvFromUtf8(Counter2) = "À": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "å"
' ConvFromUtf8(Counter2) = "Ã": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ç"
' ConvFromUtf8(Counter2) = "Õ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "é"
' ConvFromUtf8(Counter2) = "Œ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "è"
' ConvFromUtf8(Counter2) = "œ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ê"
' ConvFromUtf8(Counter2) = "–": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ë"
' ConvFromUtf8(Counter2) = "—": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "í"
' ConvFromUtf8(Counter2) = "“": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ì"
' ConvFromUtf8(Counter2) = "”": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "î"
' ConvFromUtf8(Counter2) = "‘": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ï"
' ConvFromUtf8(Counter2) = "’": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ñ"
' ConvFromUtf8(Counter2) = "÷": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ó"
' ConvFromUtf8(Counter2) = "◊": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ò"
' ConvFromUtf8(Counter2) = "ÿ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ô"
' ConvFromUtf8(Counter2) = "Ÿ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ö"
' ConvFromUtf8(Counter2) = "⁄": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "õ"
' ConvFromUtf8(Counter2) = "€": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ú"
' ConvFromUtf8(Counter2) = "‹": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ù"
' ConvFromUtf8(Counter2) = "›": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "û"
' ConvFromUtf8(Counter2) = "ﬁ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ü"
' ConvFromUtf8(Counter2) = "ﬂ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case Else
' ConvFromUtf8(Counter2) = "√": Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case "¬"
' Select Case EachChar(Counter + 1)
' Case "°"
' ConvFromUtf8(Counter2) = "°": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¢"
' ConvFromUtf8(Counter2) = "¢": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "£"
' ConvFromUtf8(Counter2) = "£": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "§"
' ConvFromUtf8(Counter2) = "§": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "•"
' ConvFromUtf8(Counter2) = "•": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¶"
' ConvFromUtf8(Counter2) = "¶": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ß"
' ConvFromUtf8(Counter2) = "ß": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "®"
' ConvFromUtf8(Counter2) = "®": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "©"
' ConvFromUtf8(Counter2) = "©": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "™"
' ConvFromUtf8(Counter2) = "™": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "´"
' ConvFromUtf8(Counter2) = "´": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¨"
' ConvFromUtf8(Counter2) = "¨": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "≠"
' ConvFromUtf8(Counter2) = "≠": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Æ"
' ConvFromUtf8(Counter2) = "Æ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ø"
' ConvFromUtf8(Counter2) = "Ø": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∞"
' ConvFromUtf8(Counter2) = "∞": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "±"
' ConvFromUtf8(Counter2) = "±": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "≤"
' ConvFromUtf8(Counter2) = "≤": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "≥"
' ConvFromUtf8(Counter2) = "≥": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "¥"
' ConvFromUtf8(Counter2) = "¥": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "µ"
' ConvFromUtf8(Counter2) = "µ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∂"
' ConvFromUtf8(Counter2) = "∂": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∑"
' ConvFromUtf8(Counter2) = "∑": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∏"
' ConvFromUtf8(Counter2) = "∏": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "π"
' ConvFromUtf8(Counter2) = "π": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "∫"
' ConvFromUtf8(Counter2) = "∫": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ª"
' ConvFromUtf8(Counter2) = "ª": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "º"
' ConvFromUtf8(Counter2) = "º": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "Ω"
' ConvFromUtf8(Counter2) = "Ω": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "æ"
' ConvFromUtf8(Counter2) = "æ": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "ø"
' ConvFromUtf8(Counter2) = "ø": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case "•"
' ConvFromUtf8(Counter2) = "•": Counter2 = Counter2 + 1: Counter = Counter + 2
' Case Else
' ConvFromUtf8(Counter2) = "¬": Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case "‚"
' Select Case EachChar(Counter + 1)
' Case "Ñ"
' Select Case EachChar(Counter + 2)
' Case "¢"
' ConvFromUtf8(Counter2) = "ô": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case "Ç"
' Select Case EachChar(Counter + 2)
' Case "¨"
' ConvFromUtf8(Counter2) = "Ä": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 1
' End Select
' Case "Ô"
' Select Case EachChar(Counter + 1)
' Case "Ä"
' Select Case EachChar(Counter + 2)
' Case "®"
' ConvFromUtf8(Counter2) = "(": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "©"
' ConvFromUtf8(Counter2) = ")": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "∫"
' ConvFromUtf8(Counter2) = ":": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "™"
' ConvFromUtf8(Counter2) = "*": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "•"
' ConvFromUtf8(Counter2) = "%": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "Æ"
' ConvFromUtf8(Counter2) = ".": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "Ø"
' ConvFromUtf8(Counter2) = "/": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "ª"
' ConvFromUtf8(Counter2) = ";": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "†"
' ConvFromUtf8(Counter2) = " ": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "ß"
' ConvFromUtf8(Counter2) = "'": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "ø"
' ConvFromUtf8(Counter2) = "?": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "º"
' ConvFromUtf8(Counter2) = "<": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "Ω"
' ConvFromUtf8(Counter2) = "=": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "æ"
' ConvFromUtf8(Counter2) = ">": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "Ä"
' ConvFromUtf8(Counter2) = "Å": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = EachChar(Counter): Counter2 = Counter2 + 1: Counter = Counter + 2
' ConvFromUtf8(Counter2) = EachChar(Counter + 1)
' End Select
' Case "Å"
' Select Case EachChar(Counter + 2)
' Case "õ"
' ConvFromUtf8(Counter2) = "[": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "ù"
' ConvFromUtf8(Counter2) = "]": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "û"
' ConvFromUtf8(Counter2) = "^": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "ú"
' ConvFromUtf8(Counter2) = "\": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case "Ä"
' ConvFromUtf8(Counter2) = "@": Counter2 = Counter2 + 1: Counter = Counter + 3
' Case Else
' ConvFromUtf8(Counter2) = "Å": Counter2 = Counter2 + 1: Counter = Counter + 2
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
