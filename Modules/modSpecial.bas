Attribute VB_Name = "modSpecial"
Option Explicit


Public Function CreateSerialNo(sPassword As String) As String
Dim lSince2000 As Long
Dim sName As String
Dim i As Long

Dim SerialStr As String
      ' Find No Days Since 2000
      lSince2000 = DateTime.DateDiff("d", "01/01/2000", DateTime.Date)
    
      sName = MDIMain.Svr1.LocalHostName
      
      SerialStr = Hex(12345) & "-"
      SerialStr = SerialStr & sPassword & "-"
      For i = 1 To Len(sName)
         SerialStr = SerialStr & Right("00" & Hex(CLng(Asc(Mid$(sName, i, 1)))), 2)
      Next i
    
      CreateSerialNo = ""
      ' Debug.Print SerialStr
      ' Debug.Print
      For i = 1 To Len(SerialStr) Step 2
         CreateSerialNo = CreateSerialNo & Mid(SerialStr, i + 1, 1)
      Next i
      CreateSerialNo = CreateSerialNo & "$"
      For i = 0 To Len(SerialStr) Step 2
         CreateSerialNo = CreateSerialNo & Mid(SerialStr, i + 1, 1)
      Next i
    
End Function
Public Function DecodeSerialNo(sERIALnO As String, GenIPAddress As String) As Boolean
      ' Dim SerialStr As String
Dim StrTemp As String
Dim lSince2000 As Long
Dim IpString As String
Dim IpSplit() As String
Dim SerialComponent() As String
Dim i As Long
Dim TmpDate As Date

Dim SerialStr As String

      On Error GoTo ErrDecode
      DecodeSerialNo = False
      IpSplit = Split(sERIALnO, "$")

      If UBound(IpSplit()) = 1 Then
         For i = 1 To Len(IpSplit(1))
            SerialStr = SerialStr & Mid(IpSplit(1), i, 1)
            SerialStr = SerialStr & Mid(IpSplit(0), i, 1)
         Next i
      End If

      SerialComponent = Split(SerialStr, "-")

      ' Decode Machine Name
      For i = 1 To Len(SerialComponent(1)) Step 2
         GenIPAddress = GenIPAddress & CLng("&H" & Mid(SerialComponent(1), i, 2)) & "."
      Next i
      DecodeSerialNo = True
      Exit Function

ErrDecode:
      DecodeSerialNo = False
End Function

Function ScrambleString(StringToScramble As String) As String
Dim StrTemp As String
Dim i As Long
Dim IntTemp As Long
Dim ScrambleKey As Byte

      ScrambleKey = Rnd(15) * 15
      StrTemp = StringToScramble

      For i = 1 To Len(StrTemp)
         IntTemp = Asc(Mid(StrTemp, i, 1)) - 32
         IntTemp = (IntTemp + ScrambleKey) Mod 96
         ScrambleString = ScrambleString & Chr((96 - IntTemp) + 32)
      Next i
      ScrambleString = ScrambleString & Hex(ScrambleKey)
End Function
Function UnScrambleString(ScrambledString As String) As String
Dim StrTemp As String
Dim i As Long
Dim IntTemp As Long
Dim ScrambleKey As Byte
   
      ScrambleKey = CByte("&H" & Right(ScrambledString, 1))
    
      For i = 1 To Len(ScrambledString) - 1
         IntTemp = Asc(Mid(ScrambledString, i, 1)) - 32
         IntTemp = 96 - IntTemp
         IntTemp = (IntTemp - ScrambleKey) Mod 96
         UnScrambleString = UnScrambleString & Chr(IntTemp + 32)
      Next i
      Debug.Print
End Function

Public Function IsUnlocked(sPassword As String) As Boolean
Dim sPath As String
Dim sKey As String
Dim sTemp As String

      sPath = "Software\EnigmaWare\Dominator"
      sKey = "UnlockCode"
   
      sTemp = ReadRegistry(HKEY_LOCAL_MACHINE, sPath, sKey)
   
      If CreateSerialNo(sPassword) = UnScrambleString(sTemp) Then
         IsUnlocked = True
      End If
   
End Function
Public Function CheckUnlocked(sPassword As String, sUnlockCode As String) As Boolean
   
      If CreateSerialNo(sPassword) = UnScrambleString(sUnlockCode) Then
         CheckUnlocked = True
      End If
   
End Function
Public Sub WriteUnlockKey(sUnlock As String)
Dim sPath As String
Dim sKey As String

      sPath = "Software\EnigmaWare\Dominator"
      sKey = "UnlockCode"
         
      Call WriteRegistry(HKEY_LOCAL_MACHINE, sPath, sKey, ValString, sUnlock)
End Sub
