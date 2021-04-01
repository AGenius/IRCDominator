Attribute VB_Name = "DO_Routines"
Option Explicit

Public Sub DoTestContent(sText As String, ByVal sNick As String)
Dim i As Integer
Dim sWord As String
Dim iUser As Integer
Dim bCapsOver As Boolean
Dim bScrollingOver As Boolean
Dim sTemp As String
Dim lCurTime As Long
Dim User As clsUser

      sNick = TestNick(sNick, True)
      ' Test Content of incomming chat text
      Set User = UsersList.Item(sNick)
      If User.Status > 0 And User.Status < 3 Then
         Exit Sub
      End If
      
      If KickSettings.Scroll_Active = True Then
         iUser = FindUser(sNick)
         If iUser = 0 Then
            ' Me
            Exit Sub
         End If
      
         ' Count Scrolling
         If User.ScrollCounter = 0 Then
            User.ScrollTime = Replace(Time, ":", "")
         End If
         If User.LastSentence = sText Then
            lCurTime = Val(Replace(Time, ":", ""))
            If lCurTime - Val(User.ScrollTime) < 10 Then
               User.ScrollCounter = User.ScrollCounter + 1
               User.ScrollTime = lCurTime
            Else
               User.ScrollCounter = 0
               User.LastSentence = sText
               User.ScrollTime = lCurTime
            End If
         Else
            User.LastSentence = sText
         End If
  
         Select Case KickSettings.Scroll_Tolerance
            Case 3
               ' High
               If User.ScrollCounter >= 10 Then bScrollingOver = True
            Case 2
               ' Tolerant
               If User.ScrollCounter >= 5 Then bScrollingOver = True
            Case 1
               ' Intolerance Tolerance
               If User.ScrollCounter >= 2 Then bScrollingOver = True
         End Select
      
         If bScrollingOver Then
            ' Do option
            If KickSettings.Scroll_Warning Then
               ' Just Do Warning Message
               User.ScrollCounter = 0
               sTemp = Replace(KickSettings.Scroll_Message, "%n", TestNick(sNick, False))
               sTemp = Replace(sTemp, vbCrLf, " ")
               Call PRIVMSG(sTemp, "", True, True, pChat, True)
               Chat "   " & TestNick(sNickJoined, False) & " : ", False, eCols.CBlack, 2
               Chat sTemp, True, FindColour(GeneralSettings.Chat_Colour), GetStyle(False), GeneralSettings.Chat_Font
            Else
               ' Do a kick
               If KickSettings.Scroll_NoBan Then
                  Call DoKick(sNick, KickSettings.Scroll_KickMessage, False)
               Else
                  Call DoKick(sNick, KickSettings.Scroll_KickMessage, True, True, KickSettings.GetBanTime(KickSettings.Scroll_BanTime), KickSettings.GetBanText(KickSettings.Scroll_BanTime))
               End If
            End If
         End If
      End If
      
      ' ------------------------------------------------------------------
      
      ' Count Caps
      
      If KickSettings.Caps_Active Then
         If IsCaps(sText) Then
            User.CapsCounter = User.CapsCounter + 1
         Else
            User.CapsCounter = 0
         End If
      
         Select Case KickSettings.Caps_Tolerance
            Case 3
               ' High
               If User.CapsCounter >= 10 Then bCapsOver = True
            Case 2
               ' Tolerant
               If User.CapsCounter >= 5 Then bCapsOver = True
            Case 1
               ' Intolerance Tolerance
               If User.CapsCounter >= 2 Then bCapsOver = True
         End Select
      
         If bCapsOver Then
            ' Do option
            If KickSettings.Caps_Warning Then
               ' Just Do Warning Message
               User.CapsCounter = 0
               sTemp = Replace(KickSettings.Caps_Message, "%n", TestNick(sNick, False))
               sTemp = Replace(sTemp, vbCrLf, " ")
               Call PRIVMSG(sTemp, "", True, True, pChat, True)
               Chat "   " & TestNick(sNickJoined, False) & " : ", False, eCols.CBlack, 2
               Chat sTemp, True, FindColour(GeneralSettings.Chat_Colour), GetStyle(False), GeneralSettings.Chat_Font
            Else
               ' Do a kick
               If KickSettings.Caps_NoBan Then
                  Call DoKick(sNick, KickSettings.Caps_Message, False)
               Else
                  Call DoKick(sNick, KickSettings.Caps_KickMessage, True, True, KickSettings.GetBanTime(KickSettings.Caps_BanTime), KickSettings.GetBanText(KickSettings.Caps_BanTime))
               End If
            End If
         End If
      End If
      
      ' ------------------------------------------------------------------
      
      If KickSettings.Profanity_Active Then
         ' Profanity test
         If KickSettings.IsInList(sText, lProfanity) Then
            ' Found do kick
            If KickSettings.Profanity_NoBan Then
               Call DoKick(sNick, KickSettings.Profanity_KickMessage, False)
            Else
               Call DoKick(sNick, KickSettings.Profanity_KickMessage, True, True, KickSettings.GetBanTime(KickSettings.Profanity_BanTime), KickSettings.GetBanText(KickSettings.Profanity_BanTime))
            End If
         End If
      End If
      If KickSettings.Advertise_Active Then
         ' Advertising Test
         
         If KickSettings.IsInList(sText, lAdvert) Then
            ' Found do kick
            If KickSettings.Advertise_NoBan Then
               Call DoKick(sNick, KickSettings.Advertise_KickMessage, False)
            Else
               Call DoKick(sNick, KickSettings.Advertise_KickMessage, True, True, KickSettings.GetBanTime(KickSettings.Advertise_BanTime), KickSettings.GetBanText(KickSettings.Advertise_BanTime))
            End If
         End If
      End If
      
End Sub

Public Sub DoKick(ByVal sNick As String, sMessage As String, bBan As Boolean, Optional bBanPassport As Boolean, Optional iTime As Integer, Optional sTime As String, Optional bSilent As Boolean)
Dim sTemp As String
Dim sKickString As String
Dim sNewNick As String
      ' Dim sGate As String
Dim i As Integer
Dim User As clsUser

      sNick = TestNick(sNick, True)
      i = FindUser(sNick)
      
      Set User = UsersList.Item(sNick)
      
      sKickString = "KICK %#" & FixSpaces(sRoomJoined, True) & " " & sNick & " :" & Replace(Trim(sMessage), vbCrLf, "")

      If bBan Then
         If bBanPassport Then
            If User.GapeKeeperID = "" Then
               sTemp = SendGetResponse("WHOIS " & sNick, "End of /WHOIS list")
               If sTemp <> "" Then
                  sNewNick = GetLine(sTemp, " 311 ")
                  User.GapeKeeperID = Split(sNewNick, " ")(4)
               End If
            End If
            sNewNick = "*!" & User.GapeKeeperID & "*$*"
            If Not (bSilent) Then
               sTemp = sKickString & " " & Replace(LoadResString(154), "%s", sTime)
               SendServer2 sTemp
               DoEvents
            End If
            sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " ADD DENY " & sNewNick & " " & iTime & " :" & sMessage & " "
            sTemp = sTemp & Replace(LoadResString(154), "%s", sTime)
            sTemp = sTemp & " - " & sNick
            SendServer2 sTemp
            DoEvents
         Else
            ' Ban Nick Name
            If Not (bSilent) Then
               sTemp = sKickString & " " & Replace(LoadResString(154), "%s", sTime)
               SendServer2 sTemp
               DoEvents
            End If
            sTemp = "ACCESS %#" & FixSpaces(sRoomJoined, True) & " ADD DENY " & sNick & " " & iTime & " :" & sMessage & " "
            sTemp = sTemp & Replace(LoadResString(154), "%s", sTime)
            SendServer2 sTemp
            DoEvents
         End If
      Else
         SendServer2 sKickString
         DoEvents
      End If
End Sub


Public Sub DoWelcome(ByVal sUserNick As String, Optional bOveride As Boolean)
Dim sTemp As String
Dim iStyle As Integer
Dim i As Integer
    
      For i = 0 To 2
         If frmWelcomePrefs.optWelcomeStyle(i).Value Then
            iStyle = i
            Exit For
         End If
      Next
    
      If bOveride Then
         iStyle = 1
      End If

      sUserNick = TestNick(sUserNick, True)
      If sUserNick <> sNickJoined Then
         sTemp = frmWelcomePrefs.txtWelcome.Text
         sTemp = Replace(sTemp, "%n", TestNick(sUserNick, False))
         sTemp = Replace(sTemp, "%r", FixSpaces(sRoomJoined, False))
         Select Case iStyle
            Case 0
               ' Whisper Welcome
               Call WHISPER(sTemp, sUserNick, True, pWelcome)
            Case 1
               ' Private Message Welcome
               Call PRIVMSG(sTemp, sUserNick, True, True, pWelcome, False, False)
            Case 2
               ' Welcome on Main Screen
               Call PRIVMSG(sTemp, "", True, True, pWelcome, False, False)
         End Select
      End If

End Sub

Public Sub DoTopic(sData As String)
Dim sTemp As String

      sTemp = GetLine(sData, " 332 ")
    
      If sTemp <> "" Then
         sTemp = GetAfter(sTemp, " :")
         If Left$(sTemp, 1) = "%" Then
            sTemp = Mid$(sTemp, 2, Len(sTemp))
         End If
         Chat LoadResString(122), False, eCols.CTeal
         Chat FixSpaces(sTemp, False), True, eCols.CBlack
         Chat "", True
      End If
End Sub

Public Sub DoNames()
Dim asNames() As String
Dim sNames() As String
Dim i As Integer
Dim iNames As Integer
Dim asNewNames() As String
Dim sName As String
Dim sTemp As String
Dim iCount As Integer
      
      If sNamesList = "" Then
         Exit Sub
      End If
      asNames = Split(sNamesList, vbCrLf)
      For i = 0 To UBound(asNames)
         If Locate(asNames(i), " 353 ") Then
            asNames(i) = Mid$(asNames(i), Locate(asNames(i), " :") + 2, Len(asNames(i)))
            ReDim Preserve asNewNames(iNames) As String
            asNewNames(iNames) = asNames(i)
            iNames = iNames + 1
         End If
      Next
      sNamesList = Join(asNewNames, " ")
      asNames = Split(sNamesList, " ")
      frmMain.tUsers.ListItems.Clear
      frmMain.tMe.ListItems.Clear
      frmMain.tUsers.Visible = False
      Set UsersList = Nothing
      Set UsersList = New clsUserList
      For i = 0 To UBound(asNames)
         sName = Split(asNames(i), ",")(3)
         If Split(asNames(i), ",")(1) <> "U" Then
            ' MSN
            sName = "^" & Mid$(sName, 2, Len(sName))
         End If
         Call UserJOIN(sName, False)
         If Left$(asNames(i), 1) = "G" Then
            ' User is away
            If Left$(sName, 1) = "@" Or Left$(sName, 1) = "." Or Left$(sName, 1) = "+" Then
               sTemp = Mid$(sName, 2, Len(sName))
            Else
               sTemp = sName
            End If
            Call UserAway(sTemp, True, False)
         End If
      Next
      frmMain.tUsers.Visible = True
      frmMain.tUsers.SortOrder = lvwAscending
      frmMain.tUsers.SortKey = 1
      On Error Resume Next
      sNamesList = ""
      Call JoinedRoom

End Sub

Public Sub DoWhisper(sMessage As String, sGate As String)
      ' Incomming Whisper
Dim sTemp As String
Dim sFontName As String
Dim asTemp() As String
Dim iFontColour As Integer
Dim iFontStyle As Integer
Dim sNick As String
Dim sMSG As String
Dim iWindow As Integer
Dim bErrNoWhisper As Boolean
Dim User As clsUser
Static bWorking As Boolean

Retest:
      If bWorking Then
         DoEvents
         GoTo Retest:
      End If
      bWorking = True
      sTemp = Mid$(sMessage, 2, Len(sMessage))
      sNick = GetBefore(sTemp, "!")
      Set User = UsersList.Item(sNick)
      If sGate <> "" Then
         User.GapeKeeperID = sGate
      End If
      If Locate(sMessage, Chr(1) & "ERR NOUSERWHISPER" & Chr(1)) Then
         bErrNoWhisper = True
      End If
      If GeneralSettings.Whisper_NoWhispers Then
         If bErrNoWhisper Then
            GoTo ExitMe:
         End If
         ' No Whispers
         If GeneralSettings.Whisper_Message Then
            ' Send User message back
            If GeneralSettings.Whisper_Response <> "" Then
               If GeneralSettings.Whisper_PrivMessage Then
                  Call PRIVMSG(GeneralSettings.Whisper_Response, sNick, True, True, pWhisper, False, True)
               Else
                  Call WHISPER(GeneralSettings.Whisper_Response, sNick, True, pWhisper)
               End If
            End If
         Else
            Call WHISPER(Chr(1) & "ERR NOUSERWHISPER" & Chr(1), sNick, False, pWhisper)
         End If
         If GeneralSettings.Whisper_Notify Then
            Chat vbTab & TestNick(sNick, False) & LoadResString(176), True, , eCols.CRed
         End If
         GoTo ExitMe:
         Exit Sub
      Else
         ' If Locate(sTemp, "@cg WHISPER ") Then
         ' asTemp = Split(sTemp, " :")
         ' Else
         asTemp = Split(sTemp, ":" & Chr(1))
         ' End If
         On Error Resume Next
         Err.Clear
         sTemp = GetAfter(asTemp(1), " ")
         If Err = 0 Then
            sFontName = GetBefore(sTemp, " ")
            sMSG = Replace(GetAfter(sTemp, " "), Chr(1), "")
            sFontName = Replace(sFontName, "\r", Chr(13))
            sFontName = Replace(sFontName, "\n", Chr(10))
            iFontColour = Asc(Mid$(sFontName, 1, 1))
            iFontStyle = Asc(Mid$(sFontName, 2, 1))
            sFontName = Mid$(sFontName, 3, Len(sFontName))
            sFontName = Split(sFontName, " ", 1)(0)
            sFontName = GetBefore(sFontName, ";")
         Else
            iFontColour = 1
            asTemp = Split(sTemp, ":")
            If Locate(sTemp, "@cg WHISPER ") Then
               sTemp = Split(sTemp, ":", 2)(1)
            Else
               sTemp = GetAfter(asTemp(1), " ")
            End If
            sMSG = sTemp
         End If

         sNick = TestNick(sNick, True)
         PlaySound Sound_Whisper
         If GeneralSettings.Whisper_Window Then
            ' Do Whisper window here
            iWindow = FindWhisperWindow(frmWhisperForm, sNick)
            If iWindow < 0 And bErrNoWhisper And frmWelcomePrefs.chkWelcome.Value Then
               If frmWelcomePrefs.optWelcomeStyle(0).Value Then
                  ' Send Welcome as priv mssg
                  Call DoWelcome(sNick, True)
               End If
            End If
            If iWindow < 0 And bErrNoWhisper = False Then
               Call ShowWhisperWindow(sNick, True)
               iWindow = FindWhisperWindow(frmWhisperForm, sNick)
               frmWhisperForm(iWindow).txtNickName.Text = ConvertFromUTF(TestNick(sNick, False))
            End If
            If bErrNoWhisper Then
               Chat LoadResString(140), True, eCols.CRed, , , frmWhisperForm(iWindow).txtMessages
               GoTo ExitMe:
            End If
            Chat "   " & TestNick(sNick, False) & " : ", False, eCols.CNavy, 2, , frmWhisperForm(iWindow).txtMessages
            Chat sMSG, True, FindColour(iFontColour), iFontStyle, sFontName, frmWhisperForm(iWindow).txtMessages
         Else
            ' Whisper on main screen
            If bErrNoWhisper Then
               Chat LoadResString(140), True, eCols.CRed
               GoTo ExitMe:
            End If
            Chat "   " & sNick, False, eCols.CNavy, 2
            Chat LoadResString(102), False, eCols.CPurple, 2, "Arial"
            Chat TestNick(sNickJoined, True), False, eCols.CNavy, 2, "Arial"
            Chat " : " & sMSG, True, FindColour(iFontColour), iFontStyle, "Arial"
         End If
      End If

ExitMe:
      bWorking = False
End Sub
Public Sub DoUserOp(sOp As String)
Dim i As Integer
Dim iState As Integer
Dim sNick As String
Dim sTemp As String
Dim iLoop As Long
Dim sList As String
Dim asList() As String

      On Error GoTo Hell
        
      sList = BuildNameList
      asList = Split(sList, ",")

      On Error Resume Next
      With frmMain
         iState = Val(Mid$(.tMe.ListItems.Item(1).SubItems(1), 6, 1))
         If iState > 0 And iState < 3 Then
            If .tMe.ListItems(1).Selected = True Then
               ' Op / Deop Me
               sNick = .tMe.ListItems.Item(1).Text
               sNick = TestNick(sNick, True)
               sTemp = "MODE %#" & FixSpaces(sRoomJoined, True) & " " & sOp & " " & sNick
               SendServer2 sTemp
               Exit Sub
            End If

            For i = 0 To UBound(asList)
               sNick = asList(i)
               sTemp = "MODE %#" & FixSpaces(sRoomJoined, True) & " " & sOp & " " & sNick
               SendServer2 sTemp
               iLoop = iLoop + 1
               If iLoop > 5 Then
                  For iLoop = 1 To 1000000
                     DoEvents
                  Next
                  iLoop = 0
               End If
            Next
         End If
      End With
        
Hell:
End Sub
Public Sub DoPass()
Dim sPass As String

      Load frmPassword
      frmPassword.Show 1
      sPass = frmPassword.Tag
      Unload frmPassword
      Set frmPassword = Nothing
      If sPass <> "" Then
         sEnteredPass = sPass
         sPass = "MODE " & TestNick(sNickJoined, True) & " +h " & sPass
      Else
         sEnteredPass = "None"
      End If
      SendServer2 sPass
End Sub
Public Sub DoModeChange(sIncomming As String)
Dim sTemp As String
Dim sType As String
Dim iUser As Integer
Dim sUserNick As String
Dim iImage As Integer
Dim sHostNick As String
Dim iState As Integer
Dim sOrigNick As String
Dim iMyState As Integer
Dim iUsersState As Integer
Dim User As clsUser
Dim HostStatus As eHostType

      iMyState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))

      On Error Resume Next
      If GetLine(sIncomming, "@GateKeeper MODE ") <> "" Or GetLine(sIncomming, "@GateKeeperPassport MODE ") <> "" Then
         ' Deal with Mode Changes
         sTemp = GetAfter(sIncomming, " MODE %#" & FixSpaces(sRoomJoined, True) & " ")
         ' On Error Resume Next
         sHostNick = Mid$(GetBefore(sIncomming, "!"), 2, Len(sIncomming))
         sHostNick = TestNick(sHostNick, False)
         With frmMain.tUsers
            If sTemp = "+m" Then
               MDIMain.mnuRoomModerated.Checked = True
               ' Spec all
               iState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))
               If iState = 0 Then
                  frmMain.tMe.ListItems.Item(1).SmallIcon = 3
                  frmMain.tMe.ListItems.Item(1).SubItems(1) = "000003_" & frmMain.tMe.ListItems.Item(1).SubItems(2)
               End If
               For iUser = 1 To .ListItems.Count
                  iState = Val(Mid$(.ListItems.Item(iUser).SubItems(1), 6, 1))
                  If .ListItems.Item(iUser).SmallIcon = 0 Then
                     .ListItems.Item(iUser).SmallIcon = 3
                     sOrigNick = .ListItems.Item(iUser).SubItems(2)
                     If Left$(sUserNick, 6) = LoadResString(167) Then
                        .ListItems(iUser).SubItems(1) = "999993_" & sOrigNick
                     Else
                        .ListItems(iUser).SubItems(1) = "888883_" & sOrigNick
                     End If
                  End If
                  .Refresh
               Next
               Exit Sub
            End If
            If sTemp = "-m" Then
               ' UnSpec All
               MDIMain.mnuRoomModerated.Checked = False
               iState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))
               If iState = 3 Then
                  frmMain.tMe.ListItems.Item(1).SmallIcon = 0
                  frmMain.tMe.ListItems.Item(1).SubItems(1) = "000000_" & frmMain.tMe.ListItems.Item(1).Text
               End If
               For iUser = 1 To .ListItems.Count
                  iState = Val(Mid$(.ListItems.Item(iUser).SubItems(1), 6, 1))
                  If .ListItems.Item(iUser).SmallIcon = 3 Then
                     .ListItems.Item(iUser).SmallIcon = 0
                     sOrigNick = .ListItems.Item(iUser).SubItems(2)
                     If Left$(sUserNick, 6) = LoadResString(167) Then
                        .ListItems(iUser).SubItems(1) = "999990_" & sOrigNick
                     Else
                        .ListItems(iUser).SubItems(1) = "888880_" & sOrigNick
                     End If
                  End If
                  .Refresh
               Next
               Exit Sub
            End If
            If Left$(sTemp, 2) = "+l" Then
               ' Room Limit
               Exit Sub
            End If

            sUserNick = Split(sTemp, " ")(1)
            If sUserNick = "" Then Exit Sub
            iUser = FindUser(sUserNick)
            If iUser > 0 Then
               Set User = UsersList.Item(sUserNick)
               iUsersState = User.HostType
            End If
            sOrigNick = sUserNick
            sUserNick = TestNick(ConvertFromUTF(sUserNick), False)
            ' iUsersState = Val(Mid$(frmMain.tUsers.ListItems.Item(iUser).SubItems(1), 6, 1))
            
            If Left$(sTemp, 2) = "-q" Then iImage = 0: User.HostType = Guest    ' Take Owner
            If Left$(sTemp, 2) = "-o" Then iImage = 0: User.HostType = Guest    ' Take Host
            If Left$(sTemp, 2) = "+q" Then iImage = 1: User.HostType = Owner    ' Make Owner
            If Left$(sTemp, 2) = "+o" Then iImage = 2: User.HostType = Host    ' Make Host
            ' Stop
            If Left$(sTemp, 2) = "+v" Then
               ' Make Me Spectator
               If iUser = 0 Then
                  iState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))
                  If iState > 0 And iState <> 3 Then
                     Exit Sub
                  End If
               Else
                  User.HostType = Spectator
                  ' iState = Val(Mid$(.ListItems.Item(iUser).SubItems(1), 6, 1))
                  ' If iState > 0 And iState <> 3 Then
                  ' Exit Sub
                  ' End If
               End If
               iImage = 0
            End If
            If Left$(sTemp, 2) = "-v" Then
               ' Make Participant
               If iUser = 0 Then
                  iState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))
                  If iState > 0 And iState <> 3 Then
                     Exit Sub
                  End If
               Else
                  User.HostType = Guest
                  ' iState = Val(Mid$(.ListItems.Item(iUser).SubItems(1), 6, 1))
                  ' If iState > 0 Then
                  ' Exit Sub
                  ' End If
               End If
               iImage = 3
            End If
            If iUser = 0 And iMyState <> iImage Then
               If sOrigNick <> sNickJoined Then
                  Exit Sub
               End If
               ' My Op Status Changed
               DoMyOpTest False
               frmMain.tMe.ListItems.Item(1).SubItems(1) = "00000" & iImage
               frmMain.tMe.ListItems.Item(1).SmallIcon = iImage
               If iImage > 0 And iImage < 3 Then
                  frmMain.tMe.ListItems.Item(1).Text = sUserNick & LoadResString(165)
               Else
                  frmMain.tMe.ListItems.Item(1).Text = sUserNick
               End If
               sTemp = LoadResString(109) & sHostNick & LoadResString(143) & sUserNick & LoadResString(144 + iImage)
               Call Notify(sTemp, "")
            Else
               If iUsersState <> iImage Then
                  .ListItems.Item(iUser).SmallIcon = iImage
                  .ListItems.Item(iUser).ForeColor = &H0&
                  If iImage <> 0 Then
                     .ListItems(iUser).SubItems(1) = "00000" & iImage
                     If iImage > 0 And iImage < 3 Then
                        .ListItems(iUser).Text = sUserNick & LoadResString(165)
                     Else
                        .ListItems(iUser).Text = sUserNick
                     End If
                  Else
                     If Left$(sUserNick, 6) = LoadResString(167) Then
                        .ListItems(iUser).SubItems(1) = "999990_" & sOrigNick
                     Else
                        .ListItems(iUser).SubItems(1) = "888880_" & sOrigNick
                     End If
                     .ListItems(iUser).Text = sUserNick
                  End If
                  User.DisplayName = .ListItems(iUser).Text
                  User.Status = iImage
                  .Sorted = True
                  sTemp = LoadResString(109) & sHostNick & LoadResString(143) & sUserNick & LoadResString(144 + iImage)
                  Call Notify(sTemp, "")
               End If
            End If
         End With
      End If
End Sub
Public Sub DoMyOpTest(bMe As Boolean)
Dim iMyState As Integer
Dim iUsersState As Integer
Dim bGuest As Boolean
Dim i As Integer

      On Error GoTo Hell

      iMyState = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))
      iUsersState = Val(Mid$(frmMain.tUsers.ListItems.Item(frmMain.tUsers.SelectedItem.Index).SubItems(1), 6, 1))
      MDIMain.mnuProfile.Enabled = True
      frmControl.cmdProfile.Enabled = True
      If Locate(frmMain.tUsers.ListItems.Item(frmMain.tUsers.SelectedItem.Index), LoadResString(167)) Then
         bGuest = True
         MDIMain.mnuProfile.Enabled = False
         frmControl.cmdProfile.Enabled = False
      End If
      If Locate(frmMain.tUsers.ListItems.Item(frmMain.tUsers.SelectedItem.Index), LoadResString(168)) Then
         bGuest = True
         MDIMain.mnuProfile.Enabled = False
         frmControl.cmdProfile.Enabled = False
      End If
      frmControl.cmdTime.Enabled = False
      frmControl.cmdIdent.Enabled = False
      ' frmControl.cmdNuke.Enabled = False
      frmControl.cmdBrown.Enabled = False
      frmControl.cmdOwner.Enabled = False
      frmControl.cmdParticipant.Enabled = False
      frmControl.cmdSpec.Enabled = False
      For i = 0 To frmControl.cmdKick.UBound
         frmControl.cmdKick(i).Enabled = False
      Next
      frmControl.cmdWhisper.Enabled = False
      frmControl.cmdAdd(0).Enabled = False
      frmControl.cmdAdd(1).Enabled = False
      frmControl.cmdAdd(2).Enabled = False
      MDIMain.mnuTime.Enabled = False
      MDIMain.mnuIdent.Enabled = False
      MDIMain.mnuHost.Enabled = False
      MDIMain.mnuOwner.Enabled = False
      MDIMain.mnuParticipant.Enabled = False
      MDIMain.mnuSpectator.Enabled = False
      MDIMain.mnuOver100.Enabled = False
      ' MDIMain.mnuProfile.Enabled = False
      MDIMain.mnuOver100.Enabled = False
      MDIMain.mnuKick.Enabled = False
      MDIMain.mnuWhisper.Enabled = False
      If frmMain.tUsers.ListItems.Count > 0 And frmMain.tUsers.SelectedItem.Index > 0 Then
         ' frmControl.cmdNuke.Enabled = True
         frmControl.cmdWhisper.Enabled = True
         MDIMain.mnuWhisper.Enabled = True
         frmControl.cmdIdent.Enabled = True
         MDIMain.mnuIdent.Enabled = True
         frmControl.cmdTime.Enabled = True
         MDIMain.mnuTime.Enabled = True
         If Not (bGuest) Then
            ' frmControl.cmdProfile.Enabled = True
            MDIMain.mnuPrefs.Enabled = True
         End If
         Select Case iMyState
            Case 1
               For i = 0 To frmControl.cmdKick.UBound
                  frmControl.cmdKick(i).Enabled = True
               Next
               frmControl.cmdAdd(0).Enabled = True
               frmControl.cmdAdd(1).Enabled = True
               frmControl.cmdAdd(2).Enabled = True
               frmControl.cmdOwner.Enabled = True
               MDIMain.mnuOver100.Enabled = True
               MDIMain.mnuOwner.Enabled = True
               frmControl.cmdParticipant.Enabled = True
               frmControl.cmdSpec.Enabled = True
               MDIMain.mnuKick.Enabled = True
               If iUsersState = 1 Then
                  frmControl.cmdOwner.Enabled = False
                  MDIMain.mnuOwner.Enabled = False
                  frmControl.cmdBrown.Enabled = True
                  MDIMain.mnuHost.Enabled = True
                  frmControl.cmdParticipant.Enabled = True
                  MDIMain.mnuParticipant.Enabled = True
               Else
                  If iUsersState = 2 Then
                     frmControl.cmdBrown.Enabled = False
                     MDIMain.mnuHost.Enabled = False
                     frmControl.cmdParticipant.Enabled = True
                     MDIMain.mnuParticipant.Enabled = True
                  Else
                     frmControl.cmdBrown.Enabled = True
                     MDIMain.mnuHost.Enabled = True
                     frmControl.cmdParticipant.Enabled = False
                     MDIMain.mnuParticipant.Enabled = False
                  End If
               End If
               ' MDIMain.mnuHost.Enabled = True
               MDIMain.mnuSpectator.Enabled = True
            Case 2
               For i = 0 To frmControl.cmdKick.UBound
                  frmControl.cmdKick(i).Enabled = True
               Next
               frmControl.cmdAdd(0).Enabled = True
               frmControl.cmdAdd(1).Enabled = True
               ' '               MDIMain.mnuOver100.Enabled = True
               frmControl.cmdAdd(2).Enabled = True
               frmControl.cmdParticipant.Enabled = False
               If iUsersState = 1 Then
                  frmControl.cmdParticipant.Enabled = False
                  frmControl.cmdBrown.Enabled = False
                  frmControl.cmdSpec.Enabled = False
                  MDIMain.mnuKick.Enabled = False
                  MDIMain.mnuParticipant.Enabled = False
                  MDIMain.mnuSpectator.Enabled = False
                  MDIMain.mnuHost.Enabled = False
               Else
                  If iUsersState = 2 Then
                     frmControl.cmdBrown.Enabled = False
                     MDIMain.mnuHost.Enabled = False
                     frmControl.cmdParticipant.Enabled = True
                     MDIMain.mnuParticipant.Enabled = True
                  Else
                     frmControl.cmdBrown.Enabled = True
                     MDIMain.mnuHost.Enabled = True
                     frmControl.cmdParticipant.Enabled = False
                     MDIMain.mnuParticipant.Enabled = False
                  End If
                  frmControl.cmdSpec.Enabled = True
                  ' MDIMain.mnuHost.Enabled = True

                  MDIMain.mnuKick.Enabled = True
                  MDIMain.mnuSpectator.Enabled = True
               End If
         End Select
      End If
      If bMe Then
         ' frmControl.cmdNuke.Enabled = False
         frmControl.cmdTime.Enabled = False
         MDIMain.mnuTime.Enabled = False
         frmControl.cmdIdent.Enabled = False
         MDIMain.mnuIdent.Enabled = False
         For i = 0 To frmControl.cmdKick.UBound
            frmControl.cmdKick(i).Enabled = False
         Next
         frmControl.cmdAdd(0).Enabled = False
         frmControl.cmdAdd(1).Enabled = False
         frmControl.cmdAdd(2).Enabled = False
         MDIMain.mnuKick.Enabled = False
         frmControl.cmdParticipant.Enabled = True
         MDIMain.mnuParticipant.Enabled = True
      End If
      If MDIMain.mnuRoomModerated.Checked = False Then
         frmControl.cmdSpec.Enabled = False
         MDIMain.mnuSpectator.Enabled = False
      End If

Hell:
End Sub
Public Sub DoRemoteKick(sData As String)
Dim sTemp As String
Dim sUserNick As String
Dim sHostName As String
Dim sReason As String
Dim sOrigNick As String
Dim User As clsUser

      ' User was KICKED
            
      sHostName = GetBefore(sData, "!")
      sHostName = TestNick(Replace(Mid$(sHostName, 2, Len(sHostName)), LoadResString(165), ""))
            
      sUserNick = Split(sData, " ")(3)
      
      sOrigNick = TestNick(sUserNick)
      
      ' Set User = UsersList.Item(sUserNick)
         
      sReason = Split(sData, " ", 5)(4)
      sReason = Mid$(sReason, 2, Len(sReason))
            
      Chat LoadResString(109) & sHostName & LoadResString(110) & sOrigNick & LoadResString(112) & sReason, True, eCols.CRed, 2
            
      UserPART sUserNick, True

      If sUserNick = sNickJoined Then
         ' I was kicked
         frmMain.tUsers.ListItems.Clear
         frmMain.tMe.ListItems.Clear
         MDIMain.mnuPart.Enabled = False
         MDIMain.mnuJoin.Enabled = True
         If GeneralSettings.AutoJoinKick Then
            MDIMain.mnuJoin_Click
         End If
      End If
End Sub

Public Sub DoMessages(sMessage As String, Optional bHideTrace As Boolean, Optional bWelcome As Boolean = False, Optional sGate As String)
      ' Incomming message
Dim sTemp As String
Dim sFontName As String
Dim asTemp() As String
Dim iFontColour As Integer
Dim iFontStyle As Integer
Dim sNick As String
Dim sMSG As String
Dim bCommand As Boolean
Dim bAction As Boolean
Dim bNuking As Boolean
Dim i As Integer

      sTemp = Mid$(sMessage, 2, Len(sMessage))
      ' If Locate(sMessage, Chr(1) & "TIME " & Chr(1)) Then
      ' bNuking = True
      ' End If
      ' If Locate(sMessage, Chr(1) & "ACTION " & Chr(1)) Then
      ' bNuking = True
      ' End If
      On Error Resume Next
      If bWelcome Then
         sNick = ""
         asTemp = Split(sTemp, " :")
         sMSG = Replace(asTemp(1), Chr(1), "")
      Else
         sNick = GetBefore(sTemp, "!")
         asTemp = Split(sTemp, ":" & Chr(1))
         bAction = False
         If UBound(asTemp) > 0 Then
            If Left$(asTemp(1), 7) = "ACTION " Then
               ' Action not message
               bAction = True
            End If
         End If
         sTemp = GetAfter(asTemp(1), " ")
         sFontName = GetBefore(sTemp, " ")
         If bAction Then
            sMSG = Replace(sTemp, Chr(1), "")
         Else
            sMSG = Replace(GetAfter(sTemp, " "), Chr(1), "")
         End If
         If sGate <> "" Then
            i = FindUser(sNick)
            If i > 0 Then
               frmMain.tUsers.ListItems(i).SubItems(7) = sGate
            End If
         End If
         ' If sMSG = "" And bNuking Then
         ' If frmNukeKick.chkKickNuke.Value Then
         ' ' Kick?
         ' With frmNukeKick
         '
         ' If bOwner Then
         ' If .chkDisable.Value Then
         ' sTemp = sNotice(sNick) & sRelockString & Chr(1)
         ' SendServer2 sTemp
         ' DoEvents
         ' End If
         ' End If
         '
         ' If .optNoBan.Value Then
         ' Call DoKick(sNick, .txtNukeMessage, False)
         ' Else
         ' Call DoKick(sNick, .txtNukeMessage, True, True, .cboBans.ItemData(.cboBans.ListIndex), .cboBans.List(.cboBans.ListIndex))
         ' End If
         ' If .chkAddKickList.Value Then
         ' .lstNickNames.AddItem sNick
         ' ' Add nuker to kick list
         ' End If
         ' Exit Sub
         ' End With
         ' End If
         ' End If
         sFontName = Replace(sFontName, "\r", Chr(13))
         sFontName = Replace(sFontName, "\n", Chr(10))
         iFontColour = Asc(Mid$(sFontName, 1, 1))
         iFontStyle = Asc(Mid$(sFontName, 2, 1))
         sFontName = Mid$(sFontName, 3, Len(sFontName))
         sFontName = Split(sFontName, " ", 1)(0)
         sFontName = Replace(GetBefore(sFontName, ";"), "\b", " ")
      End If
      bCommand = False
      If sFontName = "" Then
         ' Possible Command or IRC chatter
         Call TestForCommand(sMessage, bHideTrace, bCommand)
         If Not (bCommand) And Not (bAction) And Not (bWelcome) Then
            sMSG = Split(sMessage, " :")(1)
            ' sMSG = sMSG & "   (using irc)"
            iFontStyle = 2
            iFontColour = 1
         End If
      End If
      If Not (bCommand) Then
         sNick = TestNick(sNick, False)
         If sNick = "" Then
            ' Welcome Message
            Chat " ", True
            sMSG = Replace(sMSG, "\c", ",")
            sMSG = Replace(sMSG, "\b", " ")
            Chat Replace(sMSG, Chr(11), " "), True, eCols.CGreen, 2
         Else
'            sNick = TestNick(sNick, False)
            If bAction Then
               Chat "   " & sNick & " " & sMSG, True, eCols.CPurple, 3
            Else
               Chat "   " & sNick, False, eCols.CNavy
               Chat " : ", False, eCols.CBlack
               Chat sMSG, True, FindColour(iFontColour), iFontStyle, sFontName
            End If
         End If
      End If
      If sNick <> "" Then
         Call DoTestContent(sMSG, sNick)
      End If
End Sub

Public Sub DoTime(sNick As String)
Dim sTemp As String

      ' Time Request????
      Chat vbTab & "--> " & TestNick(sNick, False) & " " & LoadResString(175), True, eCols.CRed, 3
      PlaySound Sound_Time
      If GeneralSettings.MaskLocalTime Then
         ' Send Time Reply
         sTemp = "NOTICE " & sNick & " :" & Chr(1) & "TIME " & GeneralSettings.LocalTime & Chr(1)
      Else
         ' Send Real Reply
         sTemp = "NOTICE " & sNick & " :" & Chr(1) & "TIME " & Date & ", " & Time & Chr(1)
      End If
      SendServer2 sTemp
      DoEvents
End Sub
Public Sub DoNickChange(sData As String)
Dim sNewNick As String
Dim sOldNick As String
Dim i As Integer
Dim sTemp As String
Dim User As clsUser
      ' Dim TempUser As New clsUser

      sOldNick = Replace(GetBefore(sData, "!"), ":", "")
      sNewNick = GetAfter(sData, "NICK :")
      
      i = FindUser(sOldNick)
      
      If i > 0 Then
         Set User = UsersList.Item(sOldNick)
         UsersList.Remove sOldNick
         frmMain.tUsers.ListItems(i).SubItems(2) = sNewNick
         User.RealName = sNewNick
         sTemp = Right$(frmMain.tUsers.ListItems(i).Text, Len(LoadResString(165)))
         If sTemp = LoadResString(165) Then
            sNewNick = sNewNick & LoadResString(165)
         End If
         frmMain.tUsers.ListItems(i).Text = TestNick(sNewNick, False)
         User.DisplayName = TestNick(sNewNick, False)
         UsersList.Add User.DisplayName, User.RealName, User.Status, User.CapsCounter, User.ScrollCounter, User.ScrollTime, User.GapeKeeperID, User.SigninTime, User.Away, User.HostType, User.LastSentence, User.RealName
      Else
         sTemp = Right$(frmMain.tMe.ListItems(1).Text, Len(LoadResString(165)))
         sNickJoined = sNewNick
         If sTemp = LoadResString(165) Then
            sNewNick = sNewNick & LoadResString(165)
         End If
         frmMain.tMe.ListItems(1).Text = TestNick(sNewNick, False)
      End If
End Sub
Public Sub DoRoomMode()
Dim sTemp As String

      sWaitedFor = ""
      DoEvents
      sTemp = SendGetResponse("MODE %#" & FixSpaces(sRoomJoined, True), " 324 ")
      If sTemp = "" Then
         sTemp = SendGetResponse("MODE %#" & FixSpaces(sRoomJoined, True), " 324 ")
      End If
      If sTemp = "" Then
         Exit Sub
      End If
      sTemp = GetLine(sTemp, " 324 ")
      sTemp = Mid$(sTemp, Locate(sTemp, "%#"), Len(sTemp))
      sTemp = Replace(Mid$(sTemp, Locate(sTemp, " ") + 1, Len(sTemp)), vbCrLf, "")
      MDIMain.mnuRoomModerated.Checked = False
      MDIMain.mnuRoomSecret.Checked = False
      MDIMain.mnuRoomPrivate.Checked = False
      MDIMain.mnuRoomInvite.Checked = False
      MDIMain.mnuOver100.Checked = False
      MDIMain.mnuRoomGuestWhispers.Checked = False
      MDIMain.mnuRoomNoWhispers.Checked = False
      If Left$(sTemp, 1) = "+" Then
         ' plus
         If InStr(2, sTemp, "m") Then
            ' Moderated Room
            MDIMain.mnuRoomModerated.Checked = True
         End If
         If InStr(2, sTemp, "s") Then
            ' Secret Room
            MDIMain.mnuRoomSecret.Checked = True
         End If
         If InStr(2, sTemp, "p") Then
            ' Private Room
            MDIMain.mnuRoomPrivate.Checked = True
         End If
         If InStr(2, sTemp, "i") Then
            ' Invite Room
            MDIMain.mnuRoomInvite.Checked = True
         End If
         If InStr(2, sTemp, "w") Then
            ' No Guest Whispers Room
            MDIMain.mnuRoomGuestWhispers.Checked = True
         End If
         If InStr(2, sTemp, "W") Then
            ' No Whispers Room
            MDIMain.mnuRoomNoWhispers.Checked = True
         End If
         sTemp = GetAfter(sTemp, "l ")
         ' Room Limit
         MDIMain.mnuRoomLimit.Caption = "Set Room Limit (" & sTemp & ")"
      End If
End Sub
