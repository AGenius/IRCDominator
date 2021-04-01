Attribute VB_Name = "modFunctions"
Option Explicit
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                ' file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                ' path not found
Private Const SE_ERR_OOM = 8                ' out of memory
Private Const SE_ERR_SHARE = 26



Public Function ShellEx( _
      ByVal sFile As String, _
      Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
      Optional ByVal sParameters As String = "", _
      Optional ByVal sDefaultDir As String = "", _
      Optional sOperation As String = "open", _
      Optional Owner As Long = 0) As Boolean
        
Dim lR As Long
Dim lErr As Long, sErr As Long
      If (InStr(UCase$(sFile), ".EXE") <> 0) Then
         eShowCmd = 0
      End If
      On Error Resume Next
      If (sParameters = "") And (sDefaultDir = "") Then
         lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
      Else
         lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
      End If
      If (lR < 0) Or (lR > 32) Then
         ShellEx = True
      Else
         ' raise an appropriate error:
         lErr = vbObjectError + 1048 + lR
         Select Case lR
            Case 0
               lErr = 7: sErr = "Out of memory"
            Case ERROR_FILE_NOT_FOUND
               lErr = 53: sErr = "File not found"
            Case ERROR_PATH_NOT_FOUND
               lErr = 76: sErr = "Path not found"
            Case ERROR_BAD_FORMAT
               sErr = "The executable file is invalid or corrupt"
            Case SE_ERR_ACCESSDENIED
               lErr = 75: sErr = "Path/file access error"
            Case SE_ERR_ASSOCINCOMPLETE
               sErr = "This file type does not have a valid file association."
            Case SE_ERR_DDEBUSY
               lErr = 285: sErr = "The file could not be opened because the target application is busy. Please try again in a moment."
            Case SE_ERR_DDEFAIL
               lErr = 285: sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."
            Case SE_ERR_DDETIMEOUT
               lErr = 286: sErr = "The file could not be opened due to time out. Please try again in a moment."
            Case SE_ERR_DLLNOTFOUND
               lErr = 48: sErr = "The specified dynamic-link library was not found."
            Case SE_ERR_FNF
               lErr = 53: sErr = "File not found"
            Case SE_ERR_NOASSOC
               sErr = "No application is associated with this file type."
            Case SE_ERR_OOM
               lErr = 7: sErr = "Out of memory"
            Case SE_ERR_PNF
               lErr = 76: sErr = "Path not found"
            Case SE_ERR_SHARE
               lErr = 75: sErr = "A sharing violation occurred."
            Case Else
               sErr = "An error occurred occurred whilst trying to open or print the selected file."
         End Select
                
         Err.Raise lErr, , App.EXEName & ".GShell", sErr
         ShellEx = False
      End If

End Function


Public Function FindSystemFolder2(ByVal lngNum As Long) As String

      On Error GoTo FindSystemFolder_Err
    
Dim lpStartupPath As String * 260
Dim Pidl As Long
Dim hResult As Long
    
      ' find if a folder does exist with that number
      hResult = SHGetSpecialFolderLocation(0, lngNum, Pidl)


      If hResult = 0 Then 'there is a result
        
      ' get the actualy directory name
      hResult = SHGetPathFromIDList(ByVal Pidl, lpStartupPath)


      If hResult = 1 Then
         ' strip the string of all miscellaneous and unused characters
            
         lpStartupPath = Left$(Trim$(lpStartupPath), InStr(lpStartupPath, Chr(0)) - 1)
         FindSystemFolder2 = Trim$(lpStartupPath)
      End If
   End If
    
FindSystemFolder_End:
    
   Exit Function
    
FindSystemFolder_Err:

   ' just raise an error is a problem occurs.
   ' note that FindSystemFolder will be vbnullstring
   Err.Raise Err.Number, "FindSystemFolder::" & Err.Source, Err.Description
    
End Function

Public Sub TestFolder()
Dim i As Long
Dim sTemp As String

      For i = 0 To 48
         sTemp = FindSystemFolder2(i)
         If sTemp <> "" Then
            Debug.Print i & " - " & FindSystemFolder2(i)
         End If
      Next

End Sub
Public Function FindSystemFolder(ByVal lngNum As eSystemFolder) As String

      On Error GoTo FindSystemFolder_Err
    
Dim lpStartupPath As String * 260
Dim Pidl As Long
Dim hResult As Long
    
      ' find if a folder does exist with that number
      hResult = SHGetSpecialFolderLocation(0, lngNum, Pidl)


      If hResult = 0 Then 'there is a result
        
      ' get the actualy directory name
      hResult = SHGetPathFromIDList(ByVal Pidl, lpStartupPath)


      If hResult = 1 Then
         ' strip the string of all miscellaneous and unused characters
            
         lpStartupPath = Left$(Trim$(lpStartupPath), InStr(lpStartupPath, Chr(0)) - 1)
         FindSystemFolder = Trim$(lpStartupPath)
      End If
   End If
    
FindSystemFolder_End:
    
   Exit Function
    
FindSystemFolder_Err:

   ' just raise an error is a problem occurs.
   ' note that FindSystemFolder will be vbnullstring
   Err.Raise Err.Number, "FindSystemFolder::" & Err.Source, Err.Description
    
End Function



Public Sub StayOnTop(hwnd As Long, Stay As Boolean)

      If Stay Then
         SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
      Else
         SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
      End If
End Sub
Public Sub pShell(ByVal sWhat As String, fForm As Form)
      On Error Resume Next
      ShellEx sWhat, , , , , fForm.hwnd
      If (Err.Number <> 0) Then
         MsgBox "Sorry, I failed to open '" & sWhat & "' due to an error." & vbCrLf & vbCrLf & "[" & Err.Description & "]", vbExclamation
      End If
End Sub
'
'Public Sub Subscribe(hwnd As Long)
'      Load frmEmailAdd
'      frmEmailAdd.txtSubject = "Subscribe ircdominator"
'      frmEmailAdd.sMailtoString = "mailto:ListServer@mrenigma.btinternet.co.uk?subject=Subscribe ircdominator"
'      frmEmailAdd.txtTO = "ListServer@mrenigma.btinternet.co.uk"
'      frmEmailAdd.Show 1, MDIMain
'      Unload frmEmailAdd
'      Set frmEmailAdd = Nothing
'End Sub


Function Locate(sData As String, sFind As String) As Long
      Locate = InStr(1, sData, sFind, 1)
End Function

Public Function GetVisibleLine(rtfEdit As RichTextBox) As Long
      GetVisibleLine = SendMessageBynum(rtfEdit.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
End Function
Public Function GetLineCount(rtfEdit As RichTextBox) As Long
      GetLineCount = SendMessageBynum(rtfEdit.hwnd, EM_GETLINECOUNT, 0&, 0&)
End Function
Public Function GetVisibleLines(rtfEdit As RichTextBox) As Integer
Dim rc As RECT
Dim hdc As Long
Dim lfont As Long
Dim oldfont As Long
Dim di As Long
Dim lc As Long
Dim tm As TYPE_TEXTMETRIC

      lc = SendMessage(rtfEdit.hwnd, EM_GETRECT, 0, rc)

      lfont = SendMessageBynum(rtfEdit.hwnd, WM_GETFONT, 0, 0&)
    
      hdc = GetDC(rtfEdit.hwnd)

      If lfont <> 0 Then
         oldfont = SelectObject(hdc, lfont)
      End If

      di = GetTextMetrics(hdc, tm)
      If lfont <> 0 Then
         lfont = SelectObject(hdc, oldfont)
      End If
    
      GetVisibleLines% = ((rc.Bottom - rc.Top) / tm.tmHeight) - 1

      di = ReleaseDC(rtfEdit.hwnd, hdc)
        
    
End Function
Public Sub BuildBanList(ccboList As ComboBox)
Dim i As Integer
Dim asBans() As String
Dim asInts() As String

      asBans = Split(sBans, ",")
      asInts = Split(sInts, ",")

      ccboList.Clear
      For i = 0 To UBound(asBans)
         ccboList.AddItem asBans(i)
         ccboList.ItemData(ccboList.NewIndex) = asInts(i)
      Next
End Sub

Public Function SendGetResponse(sSend As String, sUntil As String, Optional bTimeout As Boolean = True) As String
Dim i As Long
Dim sReturnString As String
Dim bFoundIt As Boolean

      sWaitedFor = ""
      bWaiting = True
      SendServer2 sSend, False
Restart:
      For i = 1 To 900000
         
         DoEvents
         DoEvents
         If bShutDown Or bStarted = False Then
            Exit Function
         End If
         sReturnString = sReturnString & sWaitedFor
         sWaitedFor = ""
         If Locate(sReturnString, sUntil) Then
            SendGetResponse = sReturnString
            bWaiting = False
            bFoundIt = True
            sWaitedFor = ""
            Exit For
         End If
      Next
      If bTimeout = False And bFoundIt = False Then
         GoTo Restart:
      End If
      bWaiting = False
End Function

Public Sub NumChatters()
Dim sLabel As String
Dim iNum As Integer

      iNum = frmMain.tUsers.ListItems.Count + frmMain.tMe.ListItems.Count
      If iNum = 1 Then
         sLabel = LoadResString(169)
      Else
         sLabel = LoadResString(137)
         sLabel = Replace(sLabel, "%d", iNum)
      End If
      frmMain.lblUsers = sLabel
      frmMain.lblUsers.Visible = True
End Sub
Public Function BuildNameList() As String
Dim i As Integer
Dim sNick As String
Dim sList As String


      ' Build List up first as user list can change
      For i = 1 To frmMain.tUsers.ListItems.Count
         If frmMain.tUsers.ListItems.Item(i).Selected = True Then
            If frmMain.tUsers.ListItems(i).Selected = True Then
               sNick = frmMain.tUsers.ListItems.Item(i).SubItems(2)
               If sList <> "" Then
                  sList = sList & ","
               End If
               sList = sList & sNick
            End If
         End If
      Next
      BuildNameList = sList
End Function
Public Sub ResetSelected()
Dim i As Integer
        
      On Error Resume Next
      For i = 1 To frmMain.tUsers.ListItems.Count
         frmMain.tUsers.ListItems.Item(i).Selected = False
      Next
      frmMain.tMe.ListItems.Item(1).Selected = False
      i = Val(Mid$(frmMain.tMe.ListItems.Item(1).SubItems(1), 6, 1))

      If i = 0 Or i = 3 Or Err Then
         frmControl.cmdBrown.Enabled = False
         frmControl.cmdOwner.Enabled = False
         frmControl.cmdParticipant.Enabled = False
         frmControl.cmdSpec.Enabled = False
         frmControl.cmdTime.Enabled = False
         frmControl.cmdWhisper.Enabled = False
         frmControl.cmdProfile.Enabled = False
         ' frmControl.cmdNuke.Enabled = False
         frmControl.cmdAdd(0).Enabled = False
         frmControl.cmdAdd(1).Enabled = False
         frmControl.cmdAdd(2).Enabled = False
         MDIMain.mnuHost.Enabled = False
         MDIMain.mnuOwner.Enabled = False
         MDIMain.mnuParticipant.Enabled = False
         MDIMain.mnuSpectator.Enabled = False
         MDIMain.mnuTime.Enabled = False
         MDIMain.mnuWhisper.Enabled = False
         MDIMain.mnuProfile.Enabled = False
         MDIMain.mnuIdent.Enabled = False
         MDIMain.mnuKick.Enabled = False
         For i = 0 To frmControl.cmdKick.UBound
            frmControl.cmdKick(i).Enabled = False
         Next
      End If
      If frmMain.tUsers.ListItems.Count = 0 Then
         frmControl.cmdIdent.Enabled = False
      End If
        
End Sub
Public Sub CloseServer1()
Dim sPath As String

      MDIMain.Svr1.Close
      Unload frmChat
      Set frmChat = Nothing
      DoEvents
      
      sPath = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ActiveX Cache", "0")
      If GeneralSettings.ShowOCXs = True Then
         
         trace "-= ATTEMPTING TO REGISTER MSNChat45.OCX =-"
         If Dir(sPath & "\msnchat45.ocx") <> "" Then
            Shell "regsvr32.exe " & Chr(34) & sPath & "\msnchat45.ocx" & Chr(34)
         Else
            trace "-= The MSNChat45.OCX File is not in " & sPath & " =-"
         End If
   
         If Dir(App.Path & "\msnchat30.ocx") <> "" Then
            Shell "regsvr32.exe " & Chr(34) & App.Path & "\msnchat30.ocx" & Chr(34)
         Else
            trace "-= The MSNChat30.OCX File is not in the IRCDominator Application Directory =-"
         End If
      Else
         If Dir(sPath & "\msnchat45.ocx") <> "" Then
            Shell "regsvr32.exe /s " & Chr(34) & sPath & "\msnchat45.ocx" & Chr(34)
         Else
            trace "-= The MSNChat45.OCX File is not in " & sPath & " =-"
         End If
         
         If Dir(App.Path & "\msnchat30.ocx") <> "" Then
            Shell "regsvr32.exe /s " & Chr(34) & App.Path & "\msnchat30.ocx" & Chr(34)
         Else
            trace "-= The MSNChat30.OCX File is not in the IRCDominator Application Directory =-"
         End If
      End If
End Sub
Public Function FindUser(ByVal sUserNick As String) As Long
      On Error Resume Next
      With frmMain.tUsers
         FindUser = .ListItems("Name_" & sUserNick).Index
      End With
End Function

Public Function BindPort(iPort As Integer) As Integer
Dim iOldPort As Long

      ' On Error Resume Next
RepeatBind:
        
      Err.Clear
      iOldPort = iPort
        
      MDIMain.Svr1.Close
      MDIMain.Svr1.Bind iPort, "127.0.0.1"
      MDIMain.Svr1.LocalPort = iPort
      If Err.Number = "10048" Then
         Err.Clear
         iPort = iPort + 1
         GoTo RepeatBind
      End If
      BindPort = iPort
      If iOldPort <> iPort Then
         trace "Port " & iOldPort & " was already in use trying - now using port (" & iPort & ")"
      End If
End Function

Public Function BuildFontInfo(sMessage As String, PrefsSettings As ePrefs) As String

      Select Case PrefsSettings

         Case ePrefs.pChat
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(GeneralSettings.Chat_Colour) & Chr(GetStyle(PrefsSettings)) & FixSpaces(GeneralSettings.Chat_Font, True) & ";0 " & sMessage & Chr(1)
            Debug.Print GeneralSettings.Chat_Colour
         Case ePrefs.pWhisper
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(GeneralSettings.Whisper_Colour) & Chr(GetStyle(PrefsSettings)) & FixSpaces(frmPrefs.cboWhisperFont, True) & ";0 " & sMessage & Chr(1)
         Case ePrefs.pWelcome
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(frmWelcomePrefs.cboColour(0).SelectedItem.Index) & Chr(GetStyle(PrefsSettings)) & FixSpaces(frmWelcomePrefs.cboFont(0), True) & ";0 " & sMessage & Chr(1)
         Case ePrefs.pAway
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(frmWelcomePrefs.cboColour(1).SelectedItem.Index) & Chr(GetStyle(PrefsSettings)) & FixSpaces(frmWelcomePrefs.cboFont(1), True) & ";0 " & sMessage & Chr(1)
         Case ePrefs.pAdvert1
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(frmAdvertising.cboAdvertColour(0).SelectedItem.Index) & Chr(GetStyle(PrefsSettings)) & FixSpaces(frmAdvertising.cboAdvertFont(0), True) & ";0 " & sMessage & Chr(1)
         Case ePrefs.pAdvert2
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(frmAdvertising.cboAdvertColour(1).SelectedItem.Index) & Chr(GetStyle(PrefsSettings)) & FixSpaces(frmAdvertising.cboAdvertFont(1), True) & ";0 " & sMessage & Chr(1)
         Case ePrefs.pAdvert3
            BuildFontInfo = " :" & Chr(1) & "S " & Chr(frmAdvertising.cboAdvertColour(2).SelectedItem.Index) & Chr(GetStyle(PrefsSettings)) & FixSpaces(frmAdvertising.cboAdvertFont(2), True) & ";0 " & sMessage & Chr(1)

      End Select


      BuildFontInfo = Replace(BuildFontInfo, Chr(13), "\r")
      BuildFontInfo = Replace(BuildFontInfo, Chr(10), "\n")
End Function
Public Function SaveList(ByRef cListBox As ListBox, ByVal sFilePath As String)
Dim iFile As Integer
Dim i As Integer
Dim sNewPath As String

      On Error Resume Next
    
      sNewPath = App.Path & "\Lists\" & App.EXEName & "\" & sFilePath

      iFile = FreeFile
      If cListBox.ListCount > 0 Then
         Open sNewPath For Output As #iFile
        
         If Err > 0 Then
            Err.Clear
            Call MkDir(App.Path & "\Lists")
            Open sNewPath For Output As #iFile
            If Err > 0 Then
               Err.Clear
               Call MkDir(App.Path & "\Lists\" & App.EXEName)
               Open sNewPath For Output As #iFile
               If Err > 0 Then Exit Function
            End If
         End If

         For i = 0 To cListBox.ListCount - 1
            Print #iFile, cListBox.List(i)
         Next
         Close #iFile
      Else
         Open sNewPath For Output As #iFile
         Close #iFile
      End If

End Function
Public Sub LoadList(ByRef cListBox As ListBox, ByVal sFilePath As String)
Dim iFile As Integer
Dim sTemp As String
Dim sNewPath As String
Dim bOldConvert As Boolean

      On Error Resume Next

      ' sNewPath = App.Path & "\" & sFilePath
      sNewPath = App.Path & "\Lists\" & App.EXEName & "\" & sFilePath
      iFile = FreeFile
      ' If Dir(sFilePath) <> "" Then
      Open sNewPath For Input As #iFile
       
      Do While Not (EOF(iFile))
         Line Input #iFile, sTemp
         cListBox.AddItem sTemp
      Loop
      Close #iFile
      If bOldConvert Then
         Call SaveList(cListBox, sFilePath)
         Kill (App.Path & "\" & sFilePath)
      End If
    
      ' End If
      
End Sub
Public Function ConvertFromCollection(ByRef cCollection As Collection) As String
Dim i As Integer

      ' On Error Resume Next
            
      For i = 1 To cCollection.Count
         If ConvertFromCollection <> "" Then
            ConvertFromCollection = ConvertFromCollection & ","
         End If
         ConvertFromCollection = ConvertFromCollection & cCollection.Item(i)
      Next
            
End Function
Public Function ConvertToCollection(ByRef sString As String) As Collection
Dim i As Integer
Dim asTemp() As String
Dim cColTemp As New Collection
            
      
      asTemp = Split(sString, ",")
      
      For i = 0 To UBound(asTemp)
         cColTemp.Add asTemp(i)
      Next
                  
      Set ConvertToCollection = cColTemp
      Set cColTemp = Nothing
      
            
End Function

Public Function LoadListToCollection(ByVal sFilePath As String) As Collection
Dim iFile As Integer
Dim sTemp As String
Dim sNewPath As String
Dim bOldConvert As Boolean
Dim cColTemp As New Collection

      On Error Resume Next
      sNewPath = App.Path & "\Lists\" & App.EXEName & "\" & sFilePath
      iFile = FreeFile
      Open sNewPath For Input As #iFile
      Set LoadListToCollection = Nothing
      If Err > 0 Then
         Set LoadListToCollection = cColTemp
         Exit Function
      End If
      Do While Not (EOF(iFile))
         Line Input #iFile, sTemp
         cColTemp.Add sTemp, sTemp
      Loop
      Close #iFile
      Set LoadListToCollection = cColTemp
      Set cColTemp = Nothing
End Function
Public Function SaveListFromCollection(ByRef cCollection As Collection, ByVal sFilePath As String)
Dim iFile As Integer
Dim i As Integer
Dim sNewPath As String

      On Error Resume Next
    
      sNewPath = App.Path & "\Lists\" & App.EXEName & "\" & sFilePath

      iFile = FreeFile
      If cCollection.Count > 0 Then
         Open sNewPath For Output As #iFile
        
         If Err > 0 Then
            Close iFile
            Err.Clear
            Call MkDir(App.Path & "\Lists")
            Open sNewPath For Output As #iFile
            If Err > 0 Then
               Close iFile
               Err.Clear
               Call MkDir(App.Path & "\Lists\" & App.EXEName)
               Open sNewPath For Output As #iFile
               If Err > 0 Then Exit Function
            End If
         End If

         For i = 0 To cCollection.Count
            Print #iFile, cCollection.Item(i)
         Next
         Close #iFile
      Else
         Open sNewPath For Output As #iFile
         Close #iFile
      End If

End Function

Public Sub BuildColourList(ccboColours As ImageCombo)
Dim i As Integer

      ccboColours.ComboItems.Clear
      ccboColours.ImageList = MDIMain.Colours
      For i = 1 To 16
         ccboColours.ComboItems.Add i, , , i
         If bLoadingApp Then
            frmSplash.pgProgress.Value = frmSplash.pgProgress.Value + 1
            ' frmSplash.pgProgress.Refresh
         End If
      Next

End Sub

Public Sub BuildFontList(ccboList As ComboBox)
Dim i As Integer

      ccboList.Clear
      For i = 0 To Screen.FontCount - 1
         ccboList.AddItem Screen.Fonts(i)
         If bLoadingApp Then
            frmSplash.pgProgress.Value = frmSplash.pgProgress.Value + 1
            ' frmSplash.pgProgress.Refresh
         End If
      Next
End Sub
Public Function GetStyle(PrefsSettings As ePrefs) As Integer
      GetStyle = 1
      Select Case PrefsSettings

         Case ePrefs.pChat
            If GeneralSettings.Chat_StyleBold Then
               GetStyle = GetStyle + 1
            End If
            If GeneralSettings.Chat_StyleItalic Then
               GetStyle = GetStyle + 1
            End If
  
         Case ePrefs.pWhisper
            If GeneralSettings.Whisper_StyleBold Then
               GetStyle = GetStyle + 1
            End If
            If GeneralSettings.Whisper_StyleItalic Then
               GetStyle = GetStyle + 1
            End If
            
         Case ePrefs.pWelcome
            If frmWelcomePrefs.chkFontBold(0).Value Then
               GetStyle = GetStyle + 1
            End If
            If frmWelcomePrefs.chkFontItalic(0).Value Then
               GetStyle = GetStyle + 1
            End If

         Case ePrefs.pAway
            If frmWelcomePrefs.chkFontBold(1).Value Then
               GetStyle = GetStyle + 1
            End If
            If frmWelcomePrefs.chkFontItalic(1).Value Then
               GetStyle = GetStyle + 1
            End If
         Case ePrefs.pAdvert1
            If frmAdvertising.chkAdvertBold(0).Value Then
               GetStyle = GetStyle + 1
            End If
            If frmAdvertising.chkAdvertItalic(0).Value Then
               GetStyle = GetStyle + 1
            End If
         Case ePrefs.pAdvert2
            If frmAdvertising.chkAdvertBold(1).Value Then
               GetStyle = GetStyle + 1
            End If
            If frmAdvertising.chkAdvertItalic(1).Value Then
               GetStyle = GetStyle + 1
            End If
         Case ePrefs.pAdvert2
            If frmAdvertising.chkAdvertBold(2).Value Then
               GetStyle = GetStyle + 1
            End If
            If frmAdvertising.chkAdvertItalic(2).Value Then
               GetStyle = GetStyle + 1
            End If
            
      End Select
End Function
Public Function FindColour(iColour As Integer) As Integer

      FindColour = Split(sCols, ",")(iColour - 1)

End Function

Public Sub RefreshConnection(bChat As Boolean)
      frmMain.tUsers.ListItems.Clear
      frmMain.tMe.ListItems.Clear
      If bChat Then
         frmMain.txtChat.Text = ""
      End If
      frmMain.cmdAway.Tag = "UNAWAY"
      NumChatters
End Sub

Public Sub Load_Settings(fFrom As Form)
Dim i As Integer
Dim sSection As String
Dim sTitle As String
Dim asTemp() As String
Dim sValue As String
      ' *** CodeSmart ErrorHead TagStart | Please Do Not  Modify
      ' Code Added By CodeSmart
      ' =============================================================================
      On Error GoTo Err_Load_Settings:
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      ' *** CodeSmart ErrorHead TagEnd | Please Do Not Modify
        
      For i = 0 To fFrom.Controls.Count - 1
         If TypeOf fFrom.Controls(i) Is CheckBox Or _
         TypeOf fFrom.Controls(i) Is TextBox Or _
         TypeOf fFrom.Controls(i) Is ComboBox Or _
         TypeOf fFrom.Controls(i) Is Slider Or _
         TypeOf fFrom.Controls(i) Is OptionButton Or _
         TypeOf fFrom.Controls(i) Is ImageCombo Then

100         If Locate(fFrom.Controls(i).Tag, "|") Then
101            asTemp = Split(fFrom.Controls(i).Tag, "|")
102            sSection = asTemp(0)
103            sTitle = asTemp(1)
104            sValue = "_"
               ' Debug.Print fFrom.Name, sSection, sTitle

105            If TypeOf fFrom.Controls(i) Is Slider Then
106               fFrom.Controls(i).Value = fGetIni(sSection, sTitle, Trim$(Str$(fFrom.Controls(i).Value)))
               Else
107               If TypeOf fFrom.Controls(i) Is ImageCombo Then
108                  fFrom.Controls(i).ComboItems.Item(Val(fGetIni(sSection, sTitle, 1))).Selected = True
                  Else
109                  If TypeOf fFrom.Controls(i) Is OptionButton Then
110                     fFrom.Controls(i).Value = fGetIni(sSection, sTitle, Trim$(Str$(fFrom.Controls(i).Value)))
                     Else
111                     If TypeOf fFrom.Controls(i) Is CheckBox Then
                           ' Use .Value Property
112                        fFrom.Controls(i).Value = Val(fGetIni(sSection, sTitle, Trim$(Str$(fFrom.Controls(i).Value))))
                        Else
                           ' Use .Text Property
113                        If TypeOf fFrom.Controls(i) Is ComboBox Then
114                           fFrom.Controls(i).ListIndex = FindInCombo(fFrom.Controls(i), fGetIni(sSection, sTitle, Trim$(fFrom.Controls(i).Text)))
                           Else
115                           fFrom.Controls(i).Text = Replace(fGetIni(sSection, sTitle, Trim$(fFrom.Controls(i).Text)), "\r", vbCrLf)
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Next
      ' *** CodeSmart ErrorFoot TagStart | Please Do Not Modify
      ' =============================================================================
      ' =============================================================================
      ' =============================================================================
      Exit Sub
Err_Load_Settings:
      MsgBox ("Error Encounterd in Load_Settings @ " & Erl & " " & Err.Description)
      ' =============================================================================
Exit_Load_Settings:
      ' *** CodeSmart ErrorFoot TagEnd | Please Do Not Modify
End Sub
Public Sub Save_Settings(fFrom As Form)
Dim sSection As String
Dim sTitle As String
Dim asTemp() As String
Dim sValue As String
Dim i As Integer

      For i = 0 To fFrom.Controls.Count - 1
         If TypeOf fFrom.Controls(i) Is CheckBox Or _
         TypeOf fFrom.Controls(i) Is TextBox Or _
         TypeOf fFrom.Controls(i) Is OptionButton Or _
         TypeOf fFrom.Controls(i) Is ImageCombo Or _
         TypeOf fFrom.Controls(i) Is Slider Or _
         TypeOf fFrom.Controls(i) Is ComboBox Then
            
            If Locate(fFrom.Controls(i).Tag, "|") Then
               asTemp = Split(fFrom.Controls(i).Tag, "|")
               sSection = asTemp(0)
               sTitle = asTemp(1)
               sValue = "_"
               If TypeOf fFrom.Controls(i) Is ImageCombo Then
                  sValue = fFrom.Controls(i).SelectedItem.Index
               Else
                  If TypeOf fFrom.Controls(i) Is CheckBox Or TypeOf fFrom.Controls(i) Is OptionButton Or TypeOf fFrom.Controls(i) Is Slider Then
                     ' Use .Value Property
                     sValue = Trim$(Str$(fFrom.Controls(i).Value))
                  Else
                     ' Use .Text Property
                     sValue = Replace(Trim$(fFrom.Controls(i).Text), vbCrLf, "\r")
                  End If
               End If
               If sValue <> "_" Then
                  Call PutIni(sSection, sTitle, sValue)
               End If
            End If
         End If
      Next
End Sub
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
Public Sub LoadToolTips(fFrom As Form, m_objTooltip As cTooltip)
Dim asTemp()    As String
Dim sResId      As String
Dim objFont     As New StdFont
Dim i           As Long
      objFont.Name = "Tahoma"
      objFont.Size = 8
    
      Set m_objTooltip = New cTooltip
      With m_objTooltip
         .TStyle = TTBalloon
         .Icon = TTIconInfo
         .Title = "IRCDominator"
         .Centered = False
         .Create fFrom.hwnd
         .MaxWidth = 400 ' In Pixels
         .VisibleTime = 20000 ' In Milliseconds, 2000 = 2 seconds
         .DelayTime = 1000 ' In Milliseconds
         On Error Resume Next
         For i = 0 To fFrom.Controls.Count
            sResId = fFrom.Controls(i).ToolTipText
            fFrom.Controls(i).ToolTipText = ""
            If GeneralSettings.ShowToolTips = True Then
               If Locate(sResId, "|") Then
                  asTemp = Split(sResId, "|")
                  If asTemp(0) = "RES" Then
                     sResId = LoadResString(asTemp(1))
                  End If
                  If asTemp(0) = "STR" Then
                     sResId = asTemp(1)
                  End If
               End If
               If sResId <> "" Then
                  sResId = Replace(sResId, "\n", vbCrLf)
                  .AddControl fFrom.Controls(i), sResId
               End If
               sResId = ""
            End If
         Next
         Set .Font = objFont
      End With
End Sub
Public Function FindInCombo(ByRef cCombo As ComboBox, sSearch As String) As Long
    
Dim sString As String
Dim id As Integer

      On Error Resume Next
      Err.Clear
    
      For id = 0 To cCombo.ListCount - 1
         sString = UCase(cCombo.List(id))
         If sString = UCase(sSearch) Then
            FindInCombo = id
            Exit For
         End If
      Next

End Function
Public Function FindInList(ByRef cList As ListBox, sSearch As String) As Long
    
Dim sString As String
Dim id As Integer

      On Error Resume Next
      Err.Clear
      FindInList = -1
    
      For id = 0 To cList.ListCount - 1
         sString = UCase(cList.List(id))
         If sString = UCase(sSearch) Then
            FindInList = id
            Exit For
         End If
      Next

End Function
Public Sub UnlockMe()
      MDIMain.mnuPassPort.Enabled = True
      frmControl.CheckActivated
      frmMain.chkPassport.Visible = True
End Sub
Public Sub UnlockSpecial()
      UnlockMe
      MDIMain.mnuPassPort.Enabled = True
      MDIMain.mnuTrace.Visible = True
End Sub
Public Sub UnlockSpecialSpecial()
      UnlockMe
      MDIMain.mnuPassPort.Enabled = True
      MDIMain.mnuTrace.Visible = True
End Sub

Public Function TestNick(sNickName As String, Optional bRemove As Boolean) As String
      TestNick = sNickName
      If bRemove Then
         If Left$(sNickName, Len(LoadResString(167))) = LoadResString(167) Then
            TestNick = Replace(sNickName, LoadResString(167), ">")
         End If
         TestNick = Replace(TestNick, LoadResString(165), "")
      Else
         If Left$(sNickName, 1) = ">" Then
            TestNick = LoadResString(167) & Mid$(sNickName, 2, Len(sNickName))
         End If
      End If
End Function

Public Function GetAfter(sData As String, sFind As String) As String
      On Error Resume Next
      GetAfter = Mid$(sData, Locate(UCase(sData), UCase(sFind)) + Len(sFind), Len(sData))
End Function
Public Function GetBefore(sData As String, sFind As String) As String
      On Error Resume Next
      GetBefore = Left$(sData, Locate(sData, sFind) - 1)
End Function

Public Function FixSpaces(sText As String, bNoSpaces As Boolean) As String
      If bNoSpaces Then
         FixSpaces = Replace(sText, " ", "\b")
      Else
         FixSpaces = Replace(sText, "\b", " ")
      End If
End Function

Public Function GetLine(sData As String, sFind As String) As String
Dim asTemp() As String
Dim i As Integer
        
      asTemp = Split(sData, vbCrLf)
      For i = 0 To UBound(asTemp)
         If Locate(asTemp(i), sFind) Then
            GetLine = asTemp(i)
         End If
      Next
End Function
Public Sub trace(sText As String, Optional bRTF As Boolean)

      iCounter = 0
      If Left$(sText, 1) = ">" Then
         ' Out
         frmMain.imgServer.Picture = MDIMain.ilstServer.ListImages(2).Picture
         MDIMain.SetIcon
      Else
         frmMain.imgServer.Picture = MDIMain.ilstServer.ListImages(3).Picture
         MDIMain.SetIcon
      End If


      If GeneralSettings.ShowTrace Then
         With frmTrace.txtTrace
            If .Text <> "" Then
            
               With frmTraceOptions
                  If .chkAccess.Value = 0 Then
                     If Locate(sText, " 803 ") Or Locate(sText, " 804 ") Or Locate(sText, " 805 ") Then Exit Sub
                  End If
                  If .chkJoins.Value = 0 Then
                     If Locate(sText, " JOIN ") Then Exit Sub
                  End If
                  If .chkParts.Value = 0 Then
                     If Locate(sText, " PART ") Then Exit Sub
                  End If
                  If .chkPRIVMSG.Value = 0 Then
                     If Locate(sText, "PRIVMSG ") Then Exit Sub
                  End If
                  If .chkWhispers.Value = 0 Then
                     If Locate(sText, "WHISPER ") Then Exit Sub
                  End If
                  If .chkAuth.Value = 0 Then
                     If Locate(sText, "AUTH ") Then Exit Sub
                  End If
                  If .chkMODE.Value = 0 Then
                     If Locate(sText, "MODE ") Then Exit Sub
                  End If
               End With
               ' If Locate(sText, " 341 ") Then Exit Sub
               ' If Locate(sText, " 353 ") Then Exit Sub
               If Locate(sText, " 401 ") > 0 And Locate(sText, "¤§tå®_§hïÞ_t®øøÞë®¤") > 0 Then Exit Sub
               If Locate(sText, " 401 ") > 0 And Locate(sText, "hÅß§ølµ†€_G€ñïµ§l") > 0 Then Exit Sub
               If Right$(.Text, 2) <> vbCrLf Then
                  .Text = .Text & vbCrLf
               End If
            End If
            .Text = .Text & sText
            .SelStart = Len(.Text)
            If Len(.Text) > 32000 Then
               .Text = Right$(.Text, 1000)
            End If
            ' .Refresh
         End With
         frmMain.Caption = sChatCaption
      End If
End Sub

Public Sub SendServer2(ByVal sSendText As String, Optional bHide As Boolean)

      ' On Error Resume Next
      Err.Clear
      If sSendText <> "" Then
         If Err.Number = 0 And Not (bHide) Then
            trace ">" & sSendText
         End If
         sSendText = sSendText & vbCrLf
         If MDIMain.Svr2.State = 7 Then
            MDIMain.Svr2.SendData sSendText
            DoEvents
         End If
      End If
End Sub
Public Sub SendServer1(ByVal sSendText As String)

      Err.Clear
      If MDIMain.Svr1.State <> 7 Then Exit Sub
      
      If Right$(sSendText, 2) = vbCrLf Then
         sSendText = Mid$(sSendText, 1, Len(sSendText) - 2)
      End If
        
      sSendText = sSendText & vbCrLf
            
      If MDIMain.Svr1.State = 7 Then
         MDIMain.Svr1.SendData sSendText
      Else
         CloseServer1
      End If
      ' End If
End Sub
Public Function GetPassport1(sEmail As String, sPassword As String, _
      sLocaleInfo As String, _
      sBaseURL As String, _
      sRoomName As String, _
      Optional bAbout As Boolean)
               
Dim sFile As String
Dim iFile As Integer

      If bAbout Then
         sFile = "about: "
      Else
         sFile = ""
      End If
      sFile = sFile & "<html><head></head><body onload=document.form1.submit()>"
      sFile = sFile & "<form target=""_top"" name=""form1"" action=""\lcru=\base/chatroom.msnw%3frhx%3d2523" & sRoomName & "%26rhx1%3d%2525" & sRoomName & "&tw=43200&kv=2&cbid=2208&ts=-166&da=passport.com&kpp=3&ver=3.0.0000.1"" method=""post"">"
      sFile = sFile & "<input type=""text"" name=""login"" value=""\em"">"
      sFile = sFile & "<input type=""text"" name=""passwd"" value=""\ps"">"
      sFile = sFile & "<input type=""submit"" value="" Sign In "" id=""submit1"" name=""submit1"">"
      sFile = sFile & "</body></html>"

      sFile = Replace(sFile, "\em", sEmail)
      sFile = Replace(sFile, "\ps", sPassword)
      sFile = Replace(sFile, "\lc", sLocaleInfo)
      sFile = Replace(sFile, "\base", sBaseURL)
      If bAbout = False Then
         iFile = FreeFile

         Open "C:\GetPassport.html" For Output As #iFile

         Print #iFile, sFile
         Close #iFile
      End If
      GetPassport1 = sFile

End Function


Public Function NewBuildHTML(sRoomName As String, _
      sNick As String, _
      sServerIP As String, _
      sServerPORT As String, _
      Optional bHex As Boolean, _
      Optional bAbout As Boolean) As String
     
Dim sFile As String
Dim sTemp As String
Dim sServerAddress As String
Dim sRoomLabel As String
Dim iFile As Integer
Dim sOCXInfo As String
Dim sClassID As String

      sRoomLabel = "RoomName"
      If bHex Then
         sRoomLabel = "HexRoomName"
      End If
        
      sServerAddress = sServerIP & ":" & Trim(sServerPORT)
      If bAbout Then
         sFile = "about: "
      Else
         sFile = ""
      End If
      sOCXInfo = "http://fdl.msn.com/public/chat/msnchat45.cab#Version=" & GeneralSettings.ChatOCXVersion
      sFile = sFile & "<html><BODY TOPMARGIN=0 LEFTMARGIN=0>"
      sFile = sFile & "<script language=""JavaScript"">" & vbCrLf
      sFile = sFile & "var temp = '<OBJECT ID=""ChatFrame"" CLASSID=""CLSID:" & GeneralSettings.ChatCLASSID & """ width=""100%"" CODEBASE=""" & sOCXInfo & """>'"
      sFile = sFile & ";temp += '<PARAM NAME=\""RoomName\"" VALUE=\""" & FixSpaces(Trim(sRoomName), True) & "\"">'"

      sFile = sFile & ";temp += '<PARAM NAME=""Server"" VALUE=""" & sServerAddress & """>'"
      sFile = sFile & ";temp += '<PARAM NAME=""BaseURL"" VALUE=""http://chat.msn.com/"">'"
      sFile = sFile & ";temp += '<PARAM NAME=""ChatMode"" VALUE=""2"">'"

      If frmMain.chkPassport.Value Then
         If frmControl.chkClone.Value Then
            sFile = sFile & ";temp += '<PARAM NAME=""MSNREGCookie"" VALUE=""" & sNick & """>'"
            sFile = sFile & ";temp += '<PARAM NAME=""NickName"" VALUE=""" & sNick & """>'"
            sFile = sFile & ";temp += '<PARAM NAME=""PassportTicket"" VALUE=""" & Replace(frmPassPort.txtPassport(1), vbCrLf, "") & """>'"
            sFile = sFile & ";temp += '<PARAM NAME=""PassportProfile"" VALUE=""" & Replace(frmPassPort.txtPassport(2), vbCrLf, "") & """>'"
         Else
            sFile = sFile & ";temp += '<PARAM NAME=""NickName"" VALUE=""PASSPORT"">'"
            sFile = sFile & ";temp += '<PARAM NAME=""MSNREGCookie"" VALUE=""" & Replace(frmPassPort.txtPassport(0), vbCrLf, "") & """>'"
            sFile = sFile & ";temp += '<PARAM NAME=""PassportTicket"" VALUE=""" & Replace(frmPassPort.txtPassport(1), vbCrLf, "") & """>'"
            sFile = sFile & ";temp += '<PARAM NAME=""PassportProfile"" VALUE=""" & Replace(frmPassPort.txtPassport(2), vbCrLf, "") & """>'"
         End If
      Else
         sFile = sFile & ";temp += '<PARAM NAME=""NickName"" VALUE=""" & sNick & """>'"
      End If
      ' If bXMode Then
      ' sFile = sFile & "<PARAM NAME=""CreationModes"" VALUE=""x"">"
      ' End If
      ' sFile = sFile & "<PARAM NAME=""Feature"" VALUE=""1"">"
      sFile = sFile & ";temp += '<PARAM NAME=""Category"" VALUE=""UL"">'"
      ' sFile = sFile & "<PARAM NAME=""ChannelLanguage"" VALUE=""ENGLISH"">"
      sFile = sFile & ";temp += '<PARAM NAME=""Market"" VALUE=""EN-GB"">'"
      sFile = sFile & ";temp += '<PARAM NAME=""Topic"" VALUE=""Test"">'"
      sFile = sFile & ";temp += '<PARAM NAME=""WelcomeMsg"" VALUE=""Test"">'"
      sFile = sFile & ";temp += '</OBJECT>';document.write(temp);</script></html>"
      sFile = Replace(sFile, ";", ";" & vbCrLf)
      
      If bAbout = False Then
         iFile = FreeFile
        
         Open "C:\MakeRoom.html" For Output As #iFile
        
         Print #iFile, sFile
         Close #iFile
      End If
      NewBuildHTML = sFile

End Function
'
'
Public Function BuildHTML(sRoomName As String, _
      sNick As String, _
      sServerIP As String, _
      sServerPORT As String, _
      Optional bHex As Boolean, _
      Optional bAbout As Boolean) As String

Dim sFile As String
Dim sTemp As String
Dim sServerAddress As String
Dim sRoomLabel As String
Dim iFile As Integer
Dim sOCXInfo As String
Dim sClassID As String

      sRoomLabel = "RoomName"
      If bHex Then
         sRoomLabel = "HexRoomName"
      End If

      sServerAddress = sServerIP & ":" & Trim(sServerPORT)
      If bAbout Then
         sFile = "about: "
      Else
         sFile = ""
      End If
      ' sOCXInfo = "http://fdl.msn.com/public/chat/msnchat42.cab#Version=" & GeneralSettings.ChatOCXVersion
      sOCXInfo = "http://fdl.msn.com/public/chat/msnchat45.cab#Version=" & GeneralSettings.ChatOCXVersion
      sFile = sFile & "<OBJECT ID=""ChatFrame"" CLASSID=""CLSID:" & GeneralSettings.ChatCLASSID & """ width=""100%"" CODEBASE=""" & sOCXInfo & """>"
      sFile = sFile & "<PARAM NAME=""" & sRoomLabel & """ VALUE=""" & FixSpaces(Trim(sRoomName), True) & """>"
      ' If frmMain.chkPassport.Value Then
      ' If frmControl.chkClone.Value Then
      ' sFile = sFile & "<PARAM NAME=""NickName"" VALUE=""" & sNick & """>"
      ' Else
      ' sFile = sFile & "<PARAM NAME=""NickName"" VALUE=""PASSPORT"">"
      ' End If
      ' Else
      ' sFile = sFile & "<PARAM NAME=""NickName"" VALUE=""" & sNick & """>"
      ' End If
      
      sFile = sFile & "<PARAM NAME=""Server"" VALUE=""" & sServerAddress & """>"
      sFile = sFile & "<PARAM NAME=""BaseURL"" VALUE=""http://chat.msn.com/"">"
      ' If bXMode Then
      ' sFile = sFile & "<PARAM NAME=""ChatMode"" VALUE=""1"">"
      ' Else
      sFile = sFile & "<PARAM NAME=""ChatMode"" VALUE=""2"">"
      sFile = sFile & "<PARAM NAME=""WhisperContent"" VALUE=""http://chat.msn.com/whisper.msnw"" >"
      ' End If
      If frmMain.chkPassport.Value Then
         If frmControl.chkClone.Value Then
            sFile = sFile & "<PARAM NAME=""MSNREGCookie"" VALUE=""" & sNick & """>"
            sFile = sFile & "<PARAM NAME=""NickName"" VALUE=""" & sNick & """>"
            sFile = sFile & "<PARAM NAME=""PassportTicket"" VALUE=""" & Replace(frmPassPort.txtPassport(1), vbCrLf, "") & """ >"
            sFile = sFile & "<PARAM NAME=""PassportProfile"" VALUE=""2AAAAAAAACVBbTmyn3XFScOzMubvd5wtBSFkbzRYVF*nBe6eprRAUUTYfo7Fj0H71u0qVYX5yr7kaoeRdSMbSni5FHVgd8CxvlD4Cf3EEi9EKfQ2cDJS8dRjR2jjR45wh*Cc5TtlyuOWXr9FxGH5tio0QhL7L3MvIU"" >"
         Else
            sFile = sFile & "<PARAM NAME=""NickName"" VALUE=""PASSPORT"">"
            sFile = sFile & "<PARAM NAME=""MSNREGCookie"" VALUE=""" & Replace(frmPassPort.txtPassport(0), vbCrLf, "") & """>"
            sFile = sFile & "<PARAM NAME=""PassportTicket"" VALUE=""" & Replace(frmPassPort.txtPassport(1), vbCrLf, "") & """ >"
            sFile = sFile & "<PARAM NAME=""PassportProfile"" VALUE=""" & Replace(frmPassPort.txtPassport(2), vbCrLf, "") & """ >"
            ' perContent" VALUE="http://chat.msn.com/whisper.msnw">
         End If
      Else
         sFile = sFile & "<PARAM NAME=""NickName"" VALUE=""" & sNick & """>"
      End If
      ' If bXMode Then
      ' sFile = sFile & "<PARAM NAME=""CreationModes"" VALUE=""x"">"
      ' End If
      ' sFile = sFile & "<PARAM NAME=""Feature"" VALUE=""1"">"
      sFile = sFile & "<PARAM NAME=""Category"" VALUE=""UL"">"
      ' sFile = sFile & "<PARAM NAME=""ChannelLanguage"" VALUE=""ENGLISH"">"
      ' sFile = sFile & "<PARAM NAME=""Locale"" VALUE=""EN-GB"">"
      sFile = sFile & "<PARAM NAME=""Market"" VALUE=""EN-GB"">"
      sFile = sFile & "<PARAM NAME=""Topic"" VALUE=""Test"">"
      sFile = sFile & "<PARAM NAME=""WelcomeMsg"" VALUE=""Test"">"
      sFile = Replace(sFile, ">", ">" & vbCrLf)

      If bAbout = False Then
         iFile = FreeFile

         Open "C:\MakeRoom.html" For Output As #iFile

         Print #iFile, sFile
         Close #iFile
      End If
      BuildHTML = sFile

End Function

Public Function Convert2Hex(sString As String) As String
Dim i As Integer
Dim sChr As String * 1
Dim iDec As Integer
      
      For i = 1 To Len(sString)
         sChr = Mid$(sString, i, 1)
         iDec = Asc(sChr)
         Convert2Hex = Convert2Hex & Hex(iDec)
      Next
End Function

Public Function ConvertHex(sHex As String) As String
Dim i As Integer
Dim sChr As String * 1
Dim iDec As Integer

      For i = 1 To Len(sHex) Step 2
         sChr = Chr(Val("&H" & Mid$(sHex, i, 2)))
         ConvertHex = ConvertHex & sChr
      Next
End Function
Public Function ConvertHexURL(SUrl As String) As String
Dim sChr As String * 1
Dim iDec As Integer
Dim iPos As Integer
Dim sHex As String

      Do
         iPos = InStr(1, SUrl, "%")
         If iPos Then
            sHex = Mid$(SUrl, iPos + 1, 2)
            sChr = Chr(Val("&H" & sHex))
            SUrl = Left$(SUrl, iPos - 1) & sChr & Mid$(SUrl, iPos + 3, Len(SUrl))
         End If
      Loop While iPos > 0
      ConvertHexURL = SUrl
End Function

Public Function Encrypt1(sString As String) As String
Dim S As String
Dim i As Integer
Dim sHex As String
Dim sNewHex As String
Dim iDec As Integer
Dim sTmp As String

      ' This method will convert each character in the string
      ' into hex then swap the hex around then turn it back to
      ' decimal and turn it into a char
      
      For i = 1 To Len(sString)
         ' Turn Char into a hex string
       
         sHex = Hex$(Asc(Mid$(sString, i, 1)))
         
         If Len(sHex) = 1 Then
            ' Lets pad the string with zeros
            sHex = "0" & sHex
         End If

         ' Swap Hex awound for example
         ' 6E becomes E6

         sNewHex = Right$(sHex, 1) & Left$(sHex, 1)
         
         ' Convert the new hex into decimal
         
         iDec = Val("&H" & sNewHex)
         
         If iDec > 0 Then
            ' now add the char value to the new string
            
            sTmp = sTmp & Chr(iDec)
            If Len(sTmp) > 50 Then
               ' This increase performance on large strings
               S = S & sTmp
               sTmp = ""
            End If
                        
         End If
      Next
      S = S & sTmp
      Encrypt1 = S
End Function
'
'
'
' Public Function fHexRoom(sString As String) As String
' Dim S As String
' Dim i As Integer
' Dim sHex As String
' Dim sNewHex As String
' Dim iDec As Integer
' Dim sTmp As String
'
' ' This method will convert each character in the string
' ' into hex then swap the hex around then turn it back to
' ' decimal and turn it into a char
'
' For i = 1 To Len(sString) Step 2
' ' Turn Char into a hex string
'
' sHex = Mid$(sString, i, 2)
' iDec = Val("&H" & sHex)
'
' If iDec > 0 Then
' ' now add the char value to the new string
'
' sTmp = sTmp & Chr(iDec)
' If Len(sTmp) > 50 Then
' ' This increase performance on large strings
' S = S & sTmp
' sTmp = ""
' End If
'
' End If
' Next
' S = S & sTmp
' fHexRoom = S
' End Function
' Public Sub CheckOCX()
' On Error Resume Next
' MSNChatOCX.Disconnect
' If MSNChatOCX.ConnectionState <> csDisconnected And MSNChatOCX.ConnectionState <> csDisconnecting Then
' MSNChatOCX.Disconnect
' End If
' End Sub
Public Sub DoConnect(Optional bClearAll As Boolean = False)
Dim bHex As Boolean
Dim SUrl As String
Dim i As Integer
      ' Dim bX As Boolean

      Randomize
      CloseAllWindows (frmWhisperForm)
      MDIMain.tmrActivity.Enabled = False
      MDIMain.tmrJoin.Enabled = False
      bAdvertise = False
      MDIMain.mnuAbort.Enabled = True
      MDIMain.mnuJoin.Enabled = True
      bTryJoin = True
      bJoined = False
      ' sURL = ""
      ' For I = 1 To 16
      ' sURL = sURL & Chr$(Int(Rnd * 25) + 65)
      ' Next
      ' Call WriteRegistry(HKEY_LOCAL_MACHINE, sRoot & "\" & sSubKeyV3, sKey1, ValBinary, sURL)
      ' RegDeleteValue HKEY_LOCAL_MACHINE, sRoot, "{" & GeneralSettings.ChatCLASSID & "}", sKey1
      MDIMain.Svr1.Close
      MDIMain.Svr2.Close
      
      If frmMain.chkHex Then
         bHex = True
      Else
         bHex = False
      End If
      ' If frmMain.Check1 Then
      ' bX = True
      ' Else
      ' bX = False
      ' End If
      On Error Resume Next
      iAuthCount = 1
      bConnecting = False
      iPort = BindPort(iPort)
      iPort = 6667
      bConnected = False
      bStarted = False
      DoEvents
      sNamesList = ""
      If frmMain.txtRoomName <> "" Then
         ' sURL = BuildHTML(Trim(frmMain.txtRoomName), Trim(frmMain.txtNick.Text), "127.0.0.1", Str(iPort), bHex, False)
         SUrl = BuildHTML(Trim(frmMain.txtRoomName), Trim(frmMain.txtNick.Text), GeneralSettings.ServerIP, 6667, bHex, False)
      End If
      Unload frmChat
      Set frmChat = Nothing
      Load frmChat
      frmChat.wb.Navigate "C:\MakeRoom.html"
      DoEvents
      Kill "C:\MakeRoom.html"
      DoEvents
      ' If frmPrefs.chkLeaveChatOpen.Value Then
      ' frmChat.Show
      ' End If
      RefreshConnection bClearAll
      If Not (bStarted) Then
         Chat LoadResString(105), True, eCols.CGreen
         bStarted = True
      End If
      ' MDIMain.Svr2.Connect "207.46.216.29", 6667
      MDIMain.Svr2.Connect GeneralSettings.ServerIP, 6667
      ' CheckOCX
      MDIMain.tmrJoin.Interval = 20000
      MDIMain.tmrJoin.Tag = "CONNECTING"
      MDIMain.tmrJoin.Enabled = True
      ' If frmMain.chkHex Then
      ' sRoomJoined = ConvertHex(frmMain.txtRoomName)
      ' Else
      ' sRoomJoined = FixSpaces(frmMain.txtRoomName, True)
      ' End If
End Sub
Public Sub ProcessData(sData As String)
Dim asData() As String
Dim i As Integer
Dim sLine As String
Dim bHide As Boolean

      asData = Split(sData, vbCrLf)
    
      For i = 0 To UBound(asData)
         sLine = asData(i)
            
         If sLine <> "" Then
            bHide = False
            CheckResponses sLine, bHide
            If Not (bHide) Then
               trace "<" & sLine
            End If
         End If
      Next
End Sub

Public Sub JoinedRoom()
      Call NumChatters
      Call ResetSelected
      bJoined = True
      bTryJoin = False
      MDIMain.mnuAbort.Enabled = True
      MDIMain.mnuRoomOptions.Enabled = True
      MDIMain.mnuJoin.Enabled = False
      MDIMain.mnuPart.Enabled = True
      MDIMain.mnuRejoin.Enabled = True
      MDIMain.mnuPass.Enabled = True
      MDIMain.tmrActivity.Enabled = True
      bAdvertise = True
      MDIMain.tmrJoin.Tag = ""
      bConnected = True
      ' If MSNChatOCX.IsLoggedOnWithPassport Then
      ' DoRoomMode
End Sub
Public Function IsCaps(sChatText As String) As Boolean
Dim sIntText As String

      sIntText = FilterOutNonAlphaNumeric(sChatText, True)
      If Len(sIntText) < AMT_CAPS_TOLERATED Then
         ' If small string then ignore - stuff like 'OK' etc.
         Exit Function
      End If
      If sIntText = UCase(sIntText) Then
         ' Is possible
         ' Return true for suitable warning.
         IsCaps = True
      End If
End Function

' ##############################################################################################################################################
' This function takes out all non alpha-numeric characters and returns the new string. Used in IsCaps
' ##############################################################################################################################################
Public Function FilterOutNonAlphaNumeric(sText As String, bFilterOutSpace As Boolean) As String
      ' Process sIntText to strip any non-letter chars
      ' Chars of 65-90 and 97-122 are valid, all others should be stripped out
Dim bText() As Byte
Dim i As Integer
Dim sIntText As String

      bText() = sText
      For i = 0 To UBound(bText)
         If bFilterOutSpace = True Then
            If Not (bText(i) > 122 Or bText(i) < 65 Or (bText(i) > 90 And bText(i) < 97)) Then
               ' Add to external string.
               sIntText = sIntText & Chr(bText(i))
            End If
         Else
            If Not (bText(i) > 122 Or bText(i) < 65 Or (bText(i) > 90 And bText(i) < 97)) Or bText(i) = 32 Then
               ' If valid OR is a space then
               ' Add to external string. Include spaces.
               sIntText = sIntText & Chr(bText(i))
            End If
         End If
      Next i

      FilterOutNonAlphaNumeric = sIntText
End Function

Public Sub StartWhisper()
Dim i As Integer
Dim sNick As String
Dim sTemp As String
Dim iWindow As Integer

      On Error GoTo Hell
      With frmMain
         If .tMe.ListItems(1).Selected = True Then
            MsgBox (LoadResString(125))
            Exit Sub
         End If

         i = .tUsers.SelectedItem.Index
         If .tUsers.ListItems.Item(i).Selected = True Then
            sNick = .tUsers.ListItems.Item(i).SubItems(2)
            If GeneralSettings.Whisper_Window Then
               ' Do Whisper window here
               iWindow = FindWhisperWindow(frmWhisperForm, sNick)
               If iWindow < 0 Then
                  Call ShowWhisperWindow(sNick, False)
                  iWindow = FindWhisperWindow(frmWhisperForm, sNick)
               End If
               frmWhisperForm(iWindow).txtNickName = TestNick(ConvertFromUTF(sNick), False)
            End If
         End If
      End With
Hell:
        
End Sub
Public Sub MEAway(bAway As Boolean)
Dim i As Integer

      On Error GoTo Hell
        
      With frmMain.tMe
         If bAway Then
            If GeneralSettings.Notify_Aways Then
               Call Notify(LoadResString(116), "")
            End If
            .ListItems.Item(1).SmallIcon = 4
            .ListItems.Item(1).ForeColor = &H808080
         Else
            If GeneralSettings.Notify_Aways Then
               Call Notify(LoadResString(117), "")
            End If
            .ListItems.Item(1).SmallIcon = Val(Mid$(.ListItems.Item(1).SubItems(1), 6, 1))
            .ListItems.Item(1).ForeColor = &H0&
         End If
         .Refresh
      End With
Hell:
End Sub

Public Sub UserAway(ByVal sUserNick As String, bAway As Boolean, Optional bMSG As Boolean = True, Optional sGate As String)
Dim i As Integer
Dim sListNick As String
Dim User As clsUser

      ' sUserNick = TestNick(sUserNick)
        
      i = FindUser(sUserNick)
    
      If i > 0 Then
         
         Set User = UsersList.Item(sUserNick)
         
         If sGate <> "" Then
            User.GapeKeeperID = sGate
         End If
         
         With frmMain.tUsers
            ' If sGate <> "" Then
            ' .ListItems(i).SubItems(7) = sGate
            ' End If
            If bAway Then
               If GeneralSettings.Notify_Aways Then
                  If bMSG Then
                     Call Notify(LoadResString(115), TestNick(sUserNick, False))
                  End If
               End If
               .ListItems.Item(i).SmallIcon = 4
               .ListItems.Item(i).ForeColor = &H808080
               User.Away = True
            Else
               If GeneralSettings.Notify_Aways Then
                  Call Notify(LoadResString(114), TestNick(sUserNick, False))
               End If
               .ListItems.Item(i).SmallIcon = Val((Mid$(.ListItems.Item(i).SubItems(1), 6, 1)))
               .ListItems.Item(i).ForeColor = &H0&
               User.Away = False
               If frmWelcomePrefs.chkWelcomeAway.Value And frmWelcomePrefs.txtAwayMessage <> "" Then
    
                  ' Away Return Message
    
                  sListNick = frmWelcomePrefs.txtAwayMessage.Text
                  sListNick = Replace(sListNick, "%n", TestNick(sUserNick, False))
                  If frmWelcomePrefs.optAwayStyle(0).Value Then
                     ' Private Message Away Return
                     Call PRIVMSG(sListNick, sUserNick, True, True, pAway, False, False)
                  End If
                  If frmWelcomePrefs.optAwayStyle(1).Value Then
                     ' Away Return on Main Screen
                     Call PRIVMSG(sListNick, "", True, True, pAway, False, False)
                  End If
                  
               End If
            End If
            .Refresh
         End With
      End If
            
End Sub

Public Sub UserPART(ByVal sUserNick As String, bKicked As Boolean)
Dim i As Integer
Dim sListNick As String

      On Error Resume Next

      i = FindUser(sUserNick)
      If i > 0 Then
         frmMain.tUsers.ListItems.Remove (i)
            
         UsersList.Remove (sUserNick)
            
         If GeneralSettings.Notify_Leaves Then
            If Not (bKicked) Then
               PlaySound Sound_Departs
               Call Notify(LoadResString(108), TestNick(sUserNick, False))
            End If
         End If
         NumChatters
      End If
      If bKicked Then
         PlaySound Sound_Kick
      End If
End Sub
Public Sub UserJOIN(ByVal sUserNick As String, Optional bMSG As Boolean = True, Optional sGate As String)
Dim iImage As Integer
Dim nNode As ListItem
Dim sLabel As String
Dim sOrigNick As String
Dim HostStatus As eHostType
      
      iImage = 0
      HostStatus = Guest
      If sUserNick = "" Then
         Exit Sub
      End If
      If MDIMain.mnuRoomModerated.Checked Then iImage = 3
      If Left$(sUserNick, 1) = "^" Then
         ' Sysop
         sUserNick = Mid$(sUserNick, 2, Len(sUserNick))
         iImage = 6
         HostStatus = Sysop
      End If
      If Left$(sUserNick, 1) = "." Then
         ' Owner
         sUserNick = Mid$(sUserNick, 2, Len(sUserNick))
         iImage = 1
         HostStatus = Owner
      End If
      If Left$(sUserNick, 1) = "@" Then
         ' Host
         sUserNick = Mid$(sUserNick, 2, Len(sUserNick))
         iImage = 2
         HostStatus = Host
      End If
      If Left$(sUserNick, 1) = "+" Then
         ' Participant
         sUserNick = Mid$(sUserNick, 2, Len(sUserNick))
         iImage = 0
         HostStatus = Guest
      End If
      sOrigNick = sUserNick
      If Left$(sUserNick, 1) = ">" Then
         sUserNick = LoadResString(167) & Mid$(sUserNick, 2, Len(sUserNick))
      End If
      If bMSG = True Then
         If GeneralSettings.Notify_Joins Then
            Call LeaveJoinNotify(LoadResString(107), sUserNick)
         End If
      End If
      If iImage > 0 And iImage < 3 Then
         sUserNick = sUserNick & LoadResString(165)
      End If
      If sOrigNick = TestNick(sNickJoined, True) Then
         ' Myself
         On Error Resume Next
         Set nNode = frmMain.tMe.ListItems.Add(, "Name_" & CStr(iImage) & sUserNick, ConvertFromUTF(sUserNick), iImage, iImage)
         If iImage <> 0 Then
            nNode.SubItems(1) = "00000" & iImage
         Else
            nNode.SubItems(1) = "000000_" & sUserNick
         End If
      Else
         On Error Resume Next
         Set nNode = frmMain.tUsers.ListItems.Add(, "Name_" & CStr(iImage) & sOrigNick, ConvertFromUTF(sUserNick), iImage, iImage)
         
         UsersList.Add ConvertFromUTF(sUserNick), sOrigNick, iImage, 0, 0, "", sGate, Time() & " - " & Date, False, HostStatus, "", sOrigNick
         
         frmMain.tUsers.HoverSelection = False
         nNode.SubItems(2) = sOrigNick
         If iImage <> 0 Then
            nNode.SubItems(1) = "00000" & iImage
         Else
            If Left$(sUserNick, 6) = LoadResString(167) Then
               nNode.SubItems(1) = "999990_" & sOrigNick
            Else
               nNode.SubItems(1) = "888880_" & sOrigNick
            End If
         End If
      End If
      If bMSG = False Then
         Exit Sub
      End If
        
      PlaySound Sound_Arrival
      NumChatters
      If frmControl.chkAutoOwner.Value Then
         sLabel = "MODE %#" & FixSpaces(sRoomJoined, True) & " +q " & TestNick(sUserNick, True)
         SendServer2 sLabel, False
         DoEvents
      End If
      If frmControl.chkAutoHost.Value Then
         sLabel = "MODE %#" & FixSpaces(sRoomJoined, True) & " +o " & TestNick(sUserNick, True)
         SendServer2 sLabel, False
         DoEvents
      End If
      If frmControl.chkAutoV.Value Then
         sLabel = "MODE %#" & FixSpaces(sRoomJoined, True) & " +v " & TestNick(sUserNick, True)
         SendServer2 sLabel, False
         DoEvents
      End If
      If frmControl.chkAutoPart.Value Then
         sLabel = "MODE %#" & FixSpaces(sRoomJoined, True) & " -o " & TestNick(sUserNick, True)
         SendServer2 sLabel, False
         DoEvents
      End If
      If frmControl.chkAutoKick.Value Then
         Call DoKick(TestNick(sUserNick, True), frmControl.txtJoinKick, False)
         DoEvents
      End If
      
      ' Check Auto Owner List
      
      If HostLists.List_Owners_Active Then
         If HostLists.IsInList(TestNick(sUserNick, True), lOwnerList) Then
            sLabel = "MODE %#" & FixSpaces(sRoomJoined, True) & " +q " & TestNick(sUserNick, True)
            SendServer2 sLabel, False
            DoEvents
         End If
      End If
      
      ' Check Auto Host List
      If HostLists.List_Hosts_Active Then
         If HostLists.IsInList(TestNick(sUserNick, True), lHostList) Then
            sLabel = "MODE %#" & FixSpaces(sRoomJoined, True) & " +o " & TestNick(sUserNick, True)
            SendServer2 sLabel, False
            DoEvents
         End If
      End If
      
      ' Check Auto Kick List
      If KickSettings.KickList_Active Then
         If KickSettings.IsInList(TestNick(sUserNick, True), lKickList) Then
            Call DoKick(sUserNick, KickSettings.KickList_Message, False)
            DoEvents
         End If
      End If
      
      ' Check Auto Nuke Kick List
      ' If frmNukeKick.chkKickNuke.Value Then
      ' If FindInList(frmNukeKick.lstNickNames, TestNick(sUserNick, True)) > -1 Then
      ' Call DoKick(sUserNick, frmNukeKick.txtKickingMessage, False)
      ' End If
      ' End If
      
      If frmWelcomePrefs.chkWelcome.Value And frmWelcomePrefs.txtWelcome <> "" Then
         ' Welcome Message
         Call DoWelcome(sUserNick)
         DoEvents
      End If

End Sub
Public Sub CheckResponses(sIncomming As String, bHideTrace As Boolean)
Dim sTemp As String
Dim sNick As String
Dim sHostName As String
Dim sGate As String
Dim sT As String

      sGate = GetBefore(sIncomming, "@")
      sGate = GetAfter(sGate, "!")
      bHideTrace = False
      If Locate(sIncomming, " 422 ") Then
         ' Hide MOTD
         bHideTrace = True
      End If
      
      
      If Locate(sIncomming, " 366 ") Then
         DoRoomMode
         DoNames
         Exit Sub
      End If
      If Locate(sIncomming, " 353 ") Then
         ' Name List Entries
         sNamesList = sNamesList & vbCrLf & sIncomming
         Exit Sub
      End If
      If GetLine(sIncomming, " 433 ") <> "" Then
         ' Nick already in use
         If GeneralSettings.TryJoin Then
            bConnecting = False
            MDIMain.tmrJoin.Interval = 1000 * GeneralSettings.RejoinTimer
            MDIMain.tmrJoin.Tag = ""
            MDIMain.tmrJoin.Enabled = False
            MDIMain.tmrJoin.Enabled = True
            bTryJoin = True
            Exit Sub
         End If
      End If
    
      If GetLine(sIncomming, "@GateKeeper PRIVMSG ") <> "" Or GetLine(sIncomming, "@GateKeeperPassport PRIVMSG ") <> "" Or GetLine(sIncomming, "@cg PRIVMSG ") <> "" Then
         ' Process Messages
         Call DoMessages(sIncomming, bHideTrace, , sGate)
         Exit Sub
      End If
      If GetLine(sIncomming, sRoomJoined & " PRIVMSG ") <> "" Then
         ' Welcome Message
         Call DoMessages(sIncomming, bHideTrace, True, sGate)
         Exit Sub
      End If
      If GetLine(sIncomming, "@GateKeeper NOTICE ") <> "" Or GetLine(sIncomming, "@GateKeeperPassport NOTICE ") <> "" Or GetLine(sIncomming, "@cg NOTICE ") <> "" Then
         ' Process Notice
         Call DoMessages(sIncomming, bHideTrace, , sGate)
         Exit Sub
      End If
      If GetLine(sIncomming, "@GateKeeper WHISPER ") <> "" Or GetLine(sIncomming, "@GateKeeperPassport WHISPER ") <> "" Or GetLine(sIncomming, "@cg WHISPER ") <> "" Then
         ' WHISPER
         Call DoWhisper(sIncomming, sGate)
         Exit Sub
      End If
    
      If GetLine(sIncomming, " 422 ") <> "" Then
         ' Find Nick Used
         sTemp = Split(sIncomming, " ")(2)
         sNickJoined = TestNick(sTemp, True)
         If GeneralSettings.AutoJoin Then
            MDIMain.mnuJoin_Click
         End If
         Exit Sub
      End If
      If Locate(sIncomming, " 913 ") Then
         If Not (bConnected) Then
            ' No Access to room
            CloseServer1
            Chat LoadResString(170), True, eCols.CRed
         End If
         Exit Sub
      End If
      If Locate(sIncomming, " 432 ") Then
         ' Invalid Nick
         CloseServer1
         Chat LoadResString(181), True, eCols.CRed
         Exit Sub
      End If
      If Locate(sIncomming, " 910 ") Then
         ' Auth Failed
         CloseServer1
         Chat LoadResString(170), True, eCols.CRed
         Exit Sub
      End If
      If Locate(sIncomming, " 902 ") Or Locate(sIncomming, " 901 ") Then
         ' No Access to room
         CloseServer1
         Exit Sub
      End If
      If Locate(sIncomming, " 705 ") Then
         ' Room Exists
         CloseServer1
         Chat LoadResString(180), True, eCols.CRed
         Exit Sub
      End If

      If Locate(sIncomming, " 706 ") Then
         ' No Access to room
         CloseServer1
         Chat LoadResString(174), True, eCols.CGray
         Exit Sub
      End If
      If Locate(sIncomming, " 461 ") Then
         ' Insufficient Parameters
         CloseServer1
         Chat "Insufficient Parameters", True, eCols.CGray
         Exit Sub
      End If
      If Locate(sIncomming, " 471 ") Then
         ' Room Limit
         CloseServer1
         Chat LoadResString(126), True, eCols.CRed
         MDIMain.mnuJoin.Enabled = True
         MDIMain.mnuPart.Enabled = False
         Exit Sub
      End If
      If Locate(sIncomming, " 433 ") Then
         ' Nick In Use
         CloseServer1
         Chat LoadResString(129), True, eCols.CRed
         MDIMain.mnuJoin.Enabled = True
         MDIMain.mnuPart.Enabled = False
         Exit Sub
      End If
      If Locate(sIncomming, " 701 ") Or Locate(sIncomming, " 465 ") Then
         CloseServer1
         MDIMain.mnuJoin.Enabled = True
         MDIMain.mnuPart.Enabled = False
         Exit Sub
      End If
        
      If Left$(sIncomming, 5) = "PING " Then
         ' Answer PING
         SendServer2 "PONG"
         trace ">PONG"
         Exit Sub
      End If
        
      If Locate(sIncomming, " 332 ") Then
         ' Topic Entry
         sNamesList = ""
         Chat LoadResString(106), True, eCols.CRed
         DoTopic sIncomming
         Exit Sub
      End If
      If Locate(sIncomming, " 422 ") Then
         ' Room Settings
         bConnected = True
         ' DoRoomMode
         If Not (bJoined) Then
            bJoined = True
            ' sRoomMode = sIncomming
            ' DoNames
            ' If frmPrefs.chkLeaveChatOpen.Value = 0 Then
            ' CloseServer1
            ' End If
         End If
      End If

      ' If LOCATE( sIncomming, " 366 ") And Not (bConnected) Then
      ' DoRoomMode
      ' Exit Sub
      ' End If
      If Locate(sIncomming, " 822 ") Then
         ' User Away
         sNick = GetBefore(sIncomming, "!")
         sNick = Mid$(sNick, 2, Len(sNick))
         UserAway sNick, True, , sGate
         Exit Sub
      End If
      If Locate(sIncomming, " 821 ") Then
         ' User UNAway
         sNick = GetBefore(sIncomming, "!")
         sNick = Mid$(sNick, 2, Len(sNick))
         UserAway sNick, False, , sGate
         Exit Sub
      End If
      If Locate(sIncomming, "@GateKeeper PART ") Or Locate(sIncomming, "@GateKeeperPassport PART ") Or Locate(sIncomming, "@cg PART ") Then
         ' User PART
         sNick = GetBefore(sIncomming, "!")
         sNick = Mid$(sNick, 2, Len(sNick))
         If sNick = ">" & sNickJoined Then
            RefreshConnection False
            sNamesList = ""
         Else
            UserPART sNick, False
         End If
         Exit Sub
      End If
      If Locate(sIncomming, "@GateKeeper KICK ") Or Locate(sIncomming, "@GateKeeperPassport KICK ") Or Locate(sIncomming, "@cg KICK ") Then
         DoRemoteKick sIncomming
      End If
      If Locate(sIncomming, "@GateKeeper JOIN ") Or Locate(sIncomming, "@GateKeeperPassport JOIN ") Or Locate(sIncomming, "@cg JOIN ") Then
         ' User JOIN
         sNick = GetBefore(sIncomming, "!")
         sNick = Mid$(sNick, 2, Len(sNick))
         '
         If sNick <> TestNick(sNickJoined, True) Then

            sTemp = GetBefore(GetAfter(sIncomming, " JOIN "), " :")
            
            If Split(sTemp, ",")(1) <> "U" Then
               ' Sysop
               UserJOIN "^" & sNick
            Else
               On Error Resume Next
               sT = Split(sTemp, ",")(3)
               UserJOIN sT & sNick, True, sGate
            End If
         End If
         Exit Sub
      End If
      If Locate(sIncomming, "@GateKeeper MODE ") Or Locate(sIncomming, "@GateKeeperPassport MODE ") Or Locate(sIncomming, "@cg MODE ") Then
         DoModeChange sIncomming
      End If
      If Locate(sIncomming, "@GateKeeper NICK ") Then
         DoNickChange sIncomming
      End If
      If Locate(sIncomming, " TOPIC ") Then
         ' Topic Change
         sTemp = GetAfter(sIncomming, " :")
         Chat LoadResString(123), False, eCols.CTeal
         Chat FixSpaces(sTemp, False), True, eCols.CBlack
         Chat "", True
         Exit Sub
      End If

End Sub

Public Sub Notify(sMessage As String, sNick As String)
      Chat vbTab & sNick, False, eCols.CGray, 2
      Chat sMessage, True, eCols.CSilver, 2
End Sub
Public Sub LeaveJoinNotify(sMessage As String, sNick As String)
      Chat vbTab & sNick, False, eCols.CGray, 2
      Chat sMessage, True, eCols.CSilver, 2
End Sub

' Public Sub Chat(ByVal sText As String, bCr As Boolean, _
' Optional iColour As Integer = 1, _
' Optional iStyle As Integer = 1, _
' Optional sFontName As String = "Arial", _
' Optional rtTextControl As RichTextBox)
'
' Dim lStart As Long
' Dim sCR As String
' Dim lOldStart As Long
' Dim lOldLen As Long
' Dim lOldTopLine As Long
' Dim lCurrentPos As Long
'
'
' On Error Resume Next
' If TypeOf rtTextControl Is RichTextBox Then
' '
' Else
' rtTextControl = frmMain.txtChat
' End If
'
' If Err > 0 Then
' Set rtTextControl = frmMain.txtChat
' End If
'
' If bCr Then
' sCR = vbCrLf
' End If
' sText = ConvertNickname(sText)
' With frmMain.txtTemp
' lOldStart = .SelStart
' lOldLen = .SelLength
' lOldTopLine = GetVisibleLine(rtTextControl)
' lStart = Len(.Text)
' .SelStart = lStart
' .SelLength = 0
' ' .SelText = sText & sCR
' ' .SelStart = lStart
' lCurrentPos = SendMessageBynum(rtTextControl.hwnd, EM_LINEINDEX, lOldTopLine, 0&)
' '
' .SelText = sText & sCR
' .SelStart = lStart
' .SelLength = Len(sText)
' .SelRTF = sText
' .SelStart = lStart
' .SelLength = Len(sText)
' If frmPrefs.chkNoFormat.Value Then
' iColour = eCols.CBlack
' sFontName = "Arial"
' End If
' .SelColor = QBColor(iColour)
' .SelFontName = sFontName
' Select Case iStyle
' Case 1
' .SelBold = False
' .SelItalic = False
' Case 2
' .SelBold = True
' Case 3
' .SelItalic = True
' Case 4
' .SelBold = True
' .SelItalic = True
' End Select
' If bCr Then
' rtTextControl.SelStart = Len(rtTextControl.Text)
' .SelStart = 0
' .SelLength = Len(.Text)
' rtTextControl.SelRTF = .SelRTF & vbCrLf
' .Text = ""
' .TextRTF = ""
' End If
' If Not (bScollingChat) Then
' .SelStart = Len(.Text)
' Else
' If lOldLen > 0 Then
' .SelStart = lOldStart
' .SelLength = lOldLen
' Else
' .SelStart = lCurrentPos
' End If
' End If
' If Len(.Text) > 52000 Then
' .Text = Right$(.Text, LOCATE( .Text, vbCrLf))
' End If
' End With
' If frmPrefs.chkTestAlive.Value Then
' MDIMain.tmrActivity.Enabled = False
' MDIMain.tmrActivity.Enabled = True
' End If
'
' End Sub

Public Sub Chat(ByVal sText As String, bCr As Boolean, _
      Optional iColour As Integer = 1, _
      Optional iStyle As Integer = 1, _
      Optional sFontName As String = "Arial", _
      Optional rtTextControl As RichTextBox)

Dim lStart As Long
Dim sCR As String
Dim lOldStart As Long
Dim lOldLen As Long
Dim lOldTopLine As Long
Dim lCurrentPos As Long


      On Error Resume Next
      If TypeOf rtTextControl Is RichTextBox Then
         '
      Else
         rtTextControl = frmMain.txtChat
      End If

      If Err > 0 Then
         Set rtTextControl = frmMain.txtChat
      End If

      If bCr Then
         sCR = vbCrLf
      End If
      sText = ConvertFromUTF(sText)

      With rtTextControl
         lOldStart = .SelStart
         lOldLen = .SelLength
         lOldTopLine = GetVisibleLine(rtTextControl)
         lStart = Len(.Text)
         .SelStart = lStart
         .SelText = sText & sCR
         .SelStart = lStart
         .SelLength = Len(sText)
         lCurrentPos = SendMessageBynum(rtTextControl.hwnd, EM_LINEINDEX, lOldTopLine, 0&)
         '
         ' .SelText = sText & sCR
         ' .SelRTF = sText
         .SelStart = lStart
         .SelLength = Len(sText)
         If GeneralSettings.NoFormatting Then
            iColour = eCols.CBlack
            sFontName = "Arial"
         End If
         .SelColor = QBColor(iColour)
         .SelFontName = sFontName
         Select Case iStyle
            Case 1
               .SelBold = False
               .SelItalic = False
            Case 2
               .SelBold = True
            Case 3
               .SelItalic = True
            Case 4
               .SelBold = True
               .SelItalic = True
         End Select
         If Not (bScollingChat) Then
            .SelStart = Len(.Text)
         Else
            If lOldLen > 0 Then
               .SelStart = lOldStart
               .SelLength = lOldLen
            Else
               .SelStart = lCurrentPos
            End If
         End If
         If Len(.Text) > 52000 Then
            .Text = Right$(.Text, Locate(.Text, vbCrLf))
         End If
      End With
      If GeneralSettings.TestAlive Then
         MDIMain.tmrActivity.Enabled = False
         MDIMain.tmrActivity.Enabled = True
      End If

End Sub

Public Sub PRIVMSG(sMessage As String, sNick As String, bSendFontInfo As Boolean, bAddRoomName As Boolean, PrefSettings As ePrefs, Optional bChrs As Boolean = True, Optional bHideTrace As Boolean = False)
Dim sTemp As String
Dim c1 As String

      If bChrs Then
         c1 = Chr(1)
      Else
         c1 = ""
      End If
      If bAddRoomName Then
         sTemp = "PRIVMSG %#" & FixSpaces(sRoomJoined, True)
      Else
         sTemp = "PRIVMSG"
      End If
      If sNick <> "" Then
         sTemp = sTemp & " " & TestNick(sNick, True)
      End If
      If bSendFontInfo Then
         sTemp = sTemp & BuildFontInfo(sMessage, PrefSettings)
      Else
         sTemp = sTemp & " :" & c1 & sMessage & c1
      End If
      SendServer2 sTemp, bHideTrace
End Sub

Public Sub WHISPER(sMessage As String, sNick As String, bSendFontInfo As Boolean, PrefSettings As ePrefs)
Dim sTemp As String
        
      sTemp = "WHISPER %#" & FixSpaces(sRoomJoined, True)
      sTemp = sTemp & " " & TestNick(sNick, True)
      If bSendFontInfo Then
         sTemp = sTemp & BuildFontInfo(sMessage, PrefSettings)
      Else
         sTemp = sTemp & " :" & sMessage
      End If
      SendServer2 sTemp, False
End Sub

Function isVaildEmail(EmailName As String) As Boolean
Dim ipart As Integer, lpart As Integer, Length As Integer
Dim isVaild As Boolean
Dim sEmail As String

      If Len(Trim(EmailName) <= 0) Then isVaildEmail = False
      sEmail = Trim(EmailName)
      ipart = InStr(sEmail, "@")
      lpart = InStr(ipart + 1, sEmail, ".")

      Length = Len(Trim(Mid(sEmail, lpart + 1, 3)))
      If ipart <= 0 Or lpart <= 0 Then
         isVaild = False
      ElseIf Length < 3 Then
         isVaild = False
      ElseIf ipart = 1 Then
         isVaild = False
      ElseIf lpart = Len(sEmail) Then
         isVaild = False
      Else
         isVaild = True
      End If
      isVaildEmail = isVaild
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

Public Function InIDE() As Boolean
      On Error GoTo InIDEError
      InIDE = False
      Debug.Print 1 / 0
      Exit Function

InIDEError:
      InIDE = True
      Exit Function
End Function





' NUM = 1
' If winsock1.State <> sckClosed Then winsock1.Close
' If Winsock2.State <> sckClosed Then Winsock2.Close
' winsock1.Connect "207.68.167.253", 6667
'
'
' in winsock1 data arrival, scodes is my function to get the 613
' If sCodes = "613" Then
' Server2IP = Split(strData, " ")(3)
' Server2IP = Replace(Server2IP, ":", "")
' ' strData = ":TK2CHATWBA08 613 nickname :127.0.0.1 6667" & vbLf
' NUM = 2
' Winsock3.Close
' Winsock3.RemoteHost = Server2IP
' Winsock3.RemotePort = 6667
' Winsock3.Connect
' End If
' If Winsock2.State <> sckClosed Then Winsock2.SendData strData
'
' Private Sub Winsock1_Connect()
' Winsock2.Close
' Winsock2.LocalPort = 6667
' Winsock2.Listen
' End Sub
'
' Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
' If Winsock2.State <> sckClosed Then Winsock2.Close
' Winsock2.Accept requestID
' End Sub
'
' Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
' Winsock2.GetData strData
' Text1.SelText = strData
' If NUM = 1 Then
' If winsock1.State <> sckClosed Then winsock1.SendData strData
' ElseIf NUM = 2 Then
' If Winsock3.State <> sckClosed Then Winsock3.SendData strData
' End If
' End Sub
'
' Private Sub Winsock3_Connect()
' Winsock2.Close
' Winsock2.LocalPort = 6667
' Winsock2.Listen
' End Sub
'
' Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
' Winsock3.GetData strData
' Text1.SelText = strData
' If Winsock2.State = 7 Then Winsock2.SendData strData
' End Sub
'
'
'
'
'
