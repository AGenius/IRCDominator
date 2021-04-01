Attribute VB_Name = "modSoundSupport"
Option Explicit

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Enum eSoundType
   Sound_Arrival = 0
   Sound_Departs = 1
   Sound_HostMessage = 2
   Sound_HostWhisper = 3
   Sound_Invitation = 4
   Sound_Kick = 5
   Sound_TagMessage = 6
   Sound_TagWhisper = 7
   Sound_Whisper = 8
   Sound_Ident = 9
   Sound_Time = 10
End Enum

Public Sub PlayWaveFile(vntWave As Variant, Optional bWait As Boolean = False)
      If vntWave = "" Then
         Exit Sub
      End If
      If Dir(vntWave) <> "" Then
         If bWait Then
            sndPlaySound vntWave, 2
         Else
            sndPlaySound vntWave, 1
         End If
      End If
End Sub


Public Sub InstallSounds()
Dim sRegPath1 As String
Dim sTemp As String

      sRegPath1 = "AppEvents\Schemes\Apps\"
   
      sTemp = ReadRegistry(HKEY_CURRENT_USER, sRegPath1 & "IRCDominator", "")
      
      If UCase(sTemp) = "NOT FOUND" Then
         ' Entries not in registry yet
         Load frmInstalls
         DoEvents
         frmInstalls.Show 1, MDIMain
         DoEvents
         Unload frmInstalls
         Set frmInstalls = Nothing
      End If
End Sub

Public Sub InstallSoundsNow()
Dim sRegPath1 As String
Dim sRegPath2 As String
Dim sTemp As String
Dim sEvents As String
Dim sLabels As String
Dim i As Integer
Dim sKey As String
Dim sValue As String
Dim iCount As Integer

      sRegPath1 = "AppEvents\Schemes\Apps\"
      sRegPath2 = "AppEvents\EventLabels\"
      
      sEvents = GetEvents
      
      sLabels = "IRCDom_Arrival!Member Arrives,"
      sLabels = sLabels & "IRCDom_Departs!Member Departs,"
      sLabels = sLabels & "IRCDom_HostMessage!Message from Host,"
      sLabels = sLabels & "IRCDom_HostWhisper!Whisper from Host,"
      sLabels = sLabels & "IRCDom_Invitation!Invitation from Member,"
      sLabels = sLabels & "IRCDom_Kick!Member is Kicked,"
      sLabels = sLabels & "IRCDom_TagMessage!Message from Tagged Member,"
      sLabels = sLabels & "IRCDom_TagWhisper!Whisper from Tagged Member,"
      sLabels = sLabels & "IRCDom_Whisper!Incoming Whispers,"
      sLabels = sLabels & "IRCDom_Ident!Ident Request,"
      sLabels = sLabels & "IRCDom_Time!Time Request"
      frmInstalls.pgProgress.Max = UBound(Split(sEvents, ",")) + UBound(Split(sLabels, ",")) + 1
      iCount = 0
         
      For i = 0 To UBound(Split(sEvents, ","))
         frmInstalls.pgProgress.Value = iCount
         iCount = iCount + 1
         sKey = Split(Split(sEvents, ",")(i), "!")(0)
         sValue = Split(Split(sEvents, ",")(i), "!")(1)
         Call WriteRegistry(HKEY_CURRENT_USER, sRegPath1 & "IRCDominator\" & sKey, "", ValString, "")
         Call WriteRegistry(HKEY_CURRENT_USER, sRegPath1 & "IRCDominator\" & sKey & "\.Current", "", ValString, sValue)
      Next i
         
      For i = 0 To UBound(Split(sLabels, ","))
         frmInstalls.pgProgress.Value = iCount
         iCount = iCount + 1
         sKey = Split(Split(sLabels, ",")(i), "!")(0)
         sValue = Split(Split(sLabels, ",")(i), "!")(1)
         Call WriteRegistry(HKEY_CURRENT_USER, sRegPath2 & sKey, "", ValString, sValue)
      Next i

      ' [HKEY_CURRENT_USER\AppEvents\EventLabels]
      Call WriteRegistry(HKEY_CURRENT_USER, sRegPath1 & "IRCDominator", "", ValString, "IRCDominator Client")
End Sub
Public Function GetEvents() As String
      GetEvents = "IRCDom_Arrival!ChatJoin.wav,"
      GetEvents = GetEvents & "IRCDom_Departs!,"
      GetEvents = GetEvents & "IRCDom_HostMessage!,"
      GetEvents = GetEvents & "IRCDom_HostWhisper!ChatWhsp.wav,"
      GetEvents = GetEvents & "IRCDom_Invitation!ChatInvt.wav,"
      GetEvents = GetEvents & "IRCDom_Kick!ChatKick.wav,"
      GetEvents = GetEvents & "IRCDom_TagMessage!ChatTag.wav,"
      GetEvents = GetEvents & "IRCDom_TagWhisper!ChatWhsp.wav,"
      GetEvents = GetEvents & "IRCDom_Whisper!ChatWhsp.wav,"
      GetEvents = GetEvents & "IRCDom_Ident!ChatInvt.wav,"
      GetEvents = GetEvents & "IRCDom_Time!ChatInvt.wav"
End Function

Public Sub PlaySound(eSound As eSoundType)
Dim sEvent As String
Dim asEvents() As String
Dim sRegPath1 As String
Dim sMediaPath As String
Dim sSoundPath As String

      If GeneralSettings.PlaySounds Then
         sRegPath1 = "AppEvents\Schemes\Apps\"
         asEvents = Split(GetEvents, ",")
         sEvent = Split(asEvents(eSound), "!")(0)
         sSoundPath = ReadRegistry(HKEY_CURRENT_USER, sRegPath1 & "IRCDominator\" & sEvent & "\.Current", "")
         If UCase(sSoundPath) = "NOT FOUND" Then
            InstallSounds
         End If
         If sSoundPath <> "" Then
            If Locate(sSoundPath, ":") Then
               ' Absolute Path
               sMediaPath = sSoundPath
            Else
               ' Default Media folder
               sMediaPath = FindSystemFolder(fld_Windows) & "\MEDIA\" & sSoundPath
            End If
            If sMediaPath <> "" Then
               PlayWaveFile sMediaPath, False
            End If
         End If
      End If
End Sub
