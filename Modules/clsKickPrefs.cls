VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrefsKicks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eTolerance
   tIntolerent = 0
   tTolerant = 1
   tVerytolerant = 2
End Enum
Public Enum eWordList
   lAdvert = 0
   lProfanity = 1
   lKickList = 2
End Enum
Dim sNames(2) As String

Private mvarCaps_KickMessage As String
Private mvarCaps_Message As String
Private mvarScroll_BanTime As Integer
Private mvarScroll_Kick As Boolean
Private mvarScroll_KickActive As Boolean
Private mvarScroll_KickMessage As String
Private mvarScroll_Message As String
Private mvarScroll_NoBan As Boolean
Private mvarScroll_Tolerance As eTolerance
Private mvarScroll_Warning As Boolean
Private mvarCaps_BanTime As Integer
Private mvarCaps_Ban As Boolean
Private mvarCaps_Active As Boolean
Private mvarCaps_Tolerance As eTolerance
Private mvarAdvertise_Active As Boolean
Private mvarAdvertise_NoBan As Boolean
Private mvarAdvertise_Ban As Boolean
Private mvarAdvertise_BanTime As Integer
Private mvarAdvertise_Message As String
Private mvarProfanity_Active As Boolean
Private mvarProfanity_BanTime As Integer
Private mvarProfanity_Ban As Boolean
Private mvarProfanity_NoBan As Boolean
Private mvarProfanity_Message As String
Private mvarKickList_Message As String
Private mvarCaps_Kick As Boolean
Private mvarCaps_Warn As Boolean
Private mvarAdvertise_WordList As Collection
Private mvarProfanity_WordList As Collection
Private mvarKickList_NickList As Collection
Private mvarKickList_Active As Boolean
Private mvarScroll_Ban As Boolean
Private mvarCaps_NoBan As Boolean

Public Function GetBanTime(iBan As Integer) As Integer
Dim i As Integer
Dim asInts() As String

      asInts = Split(sInts, ",")
      GetBanTime = asInts(iBan)

End Function
Public Function GetBanText(iBan As Integer) As String
Dim i As Integer
Dim asBans() As String

      asBans = Split(sBans, ",")
      GetBanText = asBans(iBan)

End Function
Public Sub SaveControl(ByRef cListControl As ListBox, ByVal lListType As eWordList)
Dim i As Integer

      On Error Resume Next
      Select Case lListType
         Case eWordList.lAdvert
            For i = 1 To mvarAdvertise_WordList.Count
               If Err > 0 Then Exit For
               mvarAdvertise_WordList.Remove (1)
            Next
            For i = 0 To cListControl.ListCount - 1
               mvarAdvertise_WordList.Add cListControl.List(i), cListControl.List(i)
            Next
         Case eWordList.lKickList
            For i = 1 To mvarKickList_NickList.Count
               If Err > 0 Then Exit For
               mvarKickList_NickList.Remove (1)
            Next
            For i = 0 To cListControl.ListCount - 1
               mvarKickList_NickList.Add cListControl.List(i), cListControl.List(i)
            Next
         Case eWordList.lProfanity
            For i = 1 To mvarProfanity_WordList.Count
               If Err > 0 Then Exit For
               mvarProfanity_WordList.Remove (1)
            Next
            For i = 0 To cListControl.ListCount - 1
               mvarProfanity_WordList.Add cListControl.List(i), cListControl.List(i)
            Next
      End Select
ExitMe:
End Sub
Public Sub FillControl(ByRef cListControl As ListBox, ByVal lListType As eWordList)
Dim i As Integer

      On Error GoTo ExitMe
      cListControl.Clear
      Select Case lListType
         Case eWordList.lAdvert
            For i = 1 To mvarAdvertise_WordList.Count
               cListControl.AddItem mvarAdvertise_WordList.Item(i)
            Next
         Case eWordList.lKickList
            For i = 1 To mvarKickList_NickList.Count
               cListControl.AddItem mvarKickList_NickList.Item(i)
            Next
         Case eWordList.lProfanity
            For i = 1 To mvarProfanity_WordList.Count
               cListControl.AddItem mvarProfanity_WordList.Item(i)
            Next
      End Select
ExitMe:
End Sub
Public Sub RemoveFromList(ByVal sWord As String, ByVal lListType As eWordList)

      On Error Resume Next
      
      Select Case lListType
         Case eWordList.lAdvert
            mvarAdvertise_WordList.Remove sWord
         Case eWordList.lKickList
            mvarKickList_NickList.Remove sWord
         Case eWordList.lProfanity
            mvarProfanity_WordList.Remove sWord
      End Select
End Sub
Public Sub ClearList(ByVal lListType As eWordList)
      Select Case lListType
         Case eWordList.lAdvert
            Set mvarAdvertise_WordList = Nothing
         Case eWordList.lKickList
            Set mvarKickList_NickList = Nothing
         Case eWordList.lProfanity
            Set mvarProfanity_WordList = Nothing
      End Select
End Sub
Public Sub AddToList(ByVal sWord As String, ByVal lListType As eWordList)

      On Error Resume Next
      
      Select Case lListType
         Case eWordList.lAdvert
            mvarAdvertise_WordList.Add sWord, sWord
         Case eWordList.lKickList
            mvarKickList_NickList.Add sWord, sWord
         Case eWordList.lProfanity
            mvarProfanity_WordList.Add sWord, sWord
      End Select
End Sub
Public Function IsInList(ByVal sSentence As String, ByVal lListType As eWordList) As Boolean
Dim i As Integer
Dim sWord As String
                        
      On Error Resume Next
      
      i = -1
      Select Case lListType
         Case eWordList.lAdvert
            If mvarAdvertise_WordList.Count > 0 Then
               For i = 1 To mvarAdvertise_WordList.Count
                  sWord = mvarAdvertise_WordList.Item(i)
                  If Locate(UCase(sSentence), sWord) Then
                     IsInList = True
                  End If
               Next
            End If
            
         Case eWordList.lKickList
            If mvarKickList_NickList.Count > 0 Then
            
               sWord = mvarKickList_NickList.Item(sSentence)
               If sWord <> "" Then
                  IsInList = True
               End If
            End If
   
         Case eWordList.lProfanity
            If mvarProfanity_WordList.Count > 0 Then
               For i = 1 To mvarProfanity_WordList.Count
                  sWord = mvarProfanity_WordList.Item(i)
                  If Locate(UCase(sSentence), sWord) Then
                     IsInList = True
                  End If
               Next
            End If
      End Select
End Function
Public Sub SavePrefs()
Dim iTmp As Integer
Dim sTmp As String
Dim bValue As Boolean
Dim sSection As String

      sSection = "Kicks"
      Call PutIni(sSection, "Caps_KickMessage", CStr(mvarCaps_KickMessage))
      Call PutIni(sSection, "Caps_Message", CStr(mvarCaps_Message))
      Call PutIni(sSection, "Caps_BanTime", CStr(mvarCaps_BanTime))
      Call PutIni(sSection, "Caps_Ban", CStr(mvarCaps_Ban))
      Call PutIni(sSection, "Caps_NoBan", CStr(mvarCaps_NoBan))
      Call PutIni(sSection, "Caps_Active", CStr(mvarCaps_Active))
      Call PutIni(sSection, "Caps_Tolerance", CStr(mvarCaps_Tolerance))
      Call PutIni(sSection, "Caps_Kick", CStr(mvarCaps_Kick))
      Call PutIni(sSection, "Caps_Warn", CStr(mvarCaps_Warn))
      
      Call PutIni(sSection, "Scroll_BanTime", CStr(mvarScroll_BanTime))
      Call PutIni(sSection, "Scroll_Kick", CStr(mvarScroll_Kick))
      Call PutIni(sSection, "Scroll_Active", CStr(mvarScroll_KickActive))
      Call PutIni(sSection, "Scroll_KickMessage", CStr(mvarScroll_KickMessage))
      Call PutIni(sSection, "Scroll_Message", CStr(mvarScroll_Message))
      Call PutIni(sSection, "Scroll_NoBan", CStr(mvarScroll_NoBan))
      Call PutIni(sSection, "Scroll_Ban", CStr(mvarScroll_Ban))
      Call PutIni(sSection, "Scroll_Tolerance", CStr(mvarScroll_Tolerance))
      Call PutIni(sSection, "Scroll_Warning", CStr(mvarScroll_Warning))
      
      Call PutIni(sSection, "Advertise_Active", CStr(mvarAdvertise_Active))
      Call PutIni(sSection, "Advertise_NoBan", CStr(mvarAdvertise_NoBan))
      Call PutIni(sSection, "Advertise_Ban", CStr(mvarAdvertise_Ban))
      Call PutIni(sSection, "Advertise_BanTime", CStr(mvarAdvertise_BanTime))
      Call PutIni(sSection, "Advertise_Message", CStr(mvarAdvertise_Message))
      
      Call PutIni(sSection, "Profanity_Active", CStr(mvarProfanity_Active))
      Call PutIni(sSection, "Profanity_BanTime", CStr(mvarProfanity_BanTime))
      Call PutIni(sSection, "Profanity_Ban", CStr(mvarProfanity_Ban))
      Call PutIni(sSection, "Profanity_NoBan", CStr(mvarProfanity_NoBan))
      Call PutIni(sSection, "Profanity_Message", CStr(mvarProfanity_Message))
      
      Call PutIni(sSection, "KickList_Message", CStr(mvarKickList_Message))
      Call PutIni(sSection, "KickList_Active", CStr(mvarKickList_Active))
      Call SaveListFromCollection(mvarKickList_NickList, sNames(0))
      Call SaveListFromCollection(mvarProfanity_WordList, sNames(1))
      Call SaveListFromCollection(mvarAdvertise_WordList, sNames(2))

End Sub

Public Sub LoadPrefs()
Dim iTmp As Integer
Dim sTmp As String
Dim btmp As Boolean
Dim sSection As String

      sSection = "Kicks"

      mvarCaps_KickMessage = fGetIni(sSection, "Caps_KickMessage", "Stop Shouting")
      mvarCaps_Message = fGetIni(sSection, "Caps_Message", "Stop Shouting")
      mvarCaps_BanTime = CInt(fGetIni(sSection, "Caps_BanTime", 3))
      mvarCaps_Ban = CBool(fGetIni(sSection, "Caps_Ban", "True"))
      mvarCaps_NoBan = CBool(fGetIni(sSection, "Caps_NoBan", "False"))
      mvarCaps_Active = CBool(fGetIni(sSection, "Caps_Active", "False"))
      mvarCaps_Tolerance = CInt(fGetIni(sSection, "Caps_Tolerance", 0))
      mvarCaps_Kick = CBool(fGetIni(sSection, "Caps_Kick", "True"))
      mvarCaps_Warn = CBool(fGetIni(sSection, "Caps_Warn", "False"))

      mvarScroll_BanTime = CInt(fGetIni(sSection, "Scroll_BanTime", 3))
      mvarScroll_Kick = CBool(fGetIni(sSection, "Scroll_Kick", "True"))
      mvarScroll_KickActive = CBool(fGetIni(sSection, "Scroll_Active", "False"))
      mvarScroll_KickMessage = fGetIni(sSection, "Scroll_KickMessage", "Stop Scrolling")
      mvarScroll_Message = fGetIni(sSection, "Scroll_Message", "Stop Scrolling")
      mvarScroll_NoBan = CBool(fGetIni(sSection, "Scroll_NoBan", "False"))
      mvarScroll_Ban = CBool(fGetIni(sSection, "Scroll_Ban", "True"))
      mvarScroll_Tolerance = fGetIni(sSection, "Scroll_Tolerance", 0)
      mvarScroll_Warning = CBool(fGetIni(sSection, "Scroll_Warning", "False"))

      mvarAdvertise_Active = CBool(fGetIni(sSection, "Advertise_Active", "False"))
      mvarAdvertise_NoBan = CBool(fGetIni(sSection, "Advertise_NoBan", "False"))
      mvarAdvertise_Ban = CBool(fGetIni(sSection, "Advertise_Ban", "True"))
      mvarAdvertise_BanTime = fGetIni(sSection, "Advertise_BanTime", 3)
      mvarAdvertise_Message = fGetIni(sSection, "Advertise_Message", "Dont Advertise")

      mvarProfanity_Active = CBool(fGetIni(sSection, "Profanity_Active", "False"))
      mvarProfanity_BanTime = fGetIni(sSection, "Profanity_BanTime", 3)
      mvarProfanity_Ban = CBool(fGetIni(sSection, "Profanity_Ban", "True"))
      mvarProfanity_NoBan = CBool(fGetIni(sSection, "Profanity_NoBan", "False"))
      mvarProfanity_Message = fGetIni(sSection, "Profanity_Message", "Dont Swear")

      mvarKickList_Message = fGetIni(sSection, "KickList_Message", "Kick Listed")
      mvarKickList_Active = CBool(fGetIni(sSection, "KickList_Active", "False"))
      
      Set mvarKickList_NickList = LoadListToCollection(sNames(0))
      Set mvarProfanity_WordList = LoadListToCollection(sNames(1))
      Set mvarAdvertise_WordList = LoadListToCollection(sNames(2))
End Sub

Private Sub Class_Initialize()
      sNames(0) = "Kicks.dat"
      sNames(1) = "ProfanityKicks.dat"
      sNames(2) = "AdvertisingKicks.dat"
End Sub
Public Property Let Caps_NoBan(ByVal vData As Boolean)
      mvarCaps_NoBan = vData
End Property
Public Property Get Caps_NoBan() As Boolean
      Caps_NoBan = mvarCaps_NoBan
End Property

Public Property Let Scroll_Ban(ByVal vData As Boolean)
      mvarScroll_Ban = vData
End Property
Public Property Get Scroll_Ban() As Boolean
      Scroll_Ban = mvarScroll_Ban
End Property
Public Property Let KickList_Active(ByVal vData As Boolean)
      mvarKickList_Active = vData
End Property
Public Property Get KickList_Active() As Boolean
      KickList_Active = mvarKickList_Active
End Property

Public Property Let KickList_NickList(ByVal vData As String)
      Set mvarKickList_NickList = ConvertToCollection(vData)
End Property
Public Property Get KickList_NickList() As String
      KickList_NickList = ConvertFromCollection(mvarKickList_NickList)
End Property
Public Property Let Profanity_WordList(ByVal vData As String)
      Set mvarProfanity_WordList = ConvertToCollection(vData)
End Property
Public Property Get Profanity_WordList() As String
      Profanity_WordList = ConvertFromCollection(mvarProfanity_WordList)
End Property
Public Property Let Advertise_WordList(ByVal vData As String)
      Set mvarAdvertise_WordList = ConvertToCollection(vData)
End Property
Public Property Get Advertise_WordList() As String
      Advertise_WordList = ConvertFromCollection(mvarAdvertise_WordList)
End Property
Public Property Let Caps_Warning(ByVal vData As Boolean)
      mvarCaps_Warn = vData
End Property
Public Property Get Caps_Warning() As Boolean
      Caps_Warning = mvarCaps_Warn
End Property
Public Property Let Caps_Kick(ByVal vData As Boolean)
      mvarCaps_Kick = vData
End Property
Public Property Get Caps_Kick() As Boolean
      Caps_Kick = mvarCaps_Kick
End Property
Public Property Let KickList_Message(ByVal vData As String)
      mvarKickList_Message = vData
End Property
Public Property Get KickList_Message() As String
      KickList_Message = mvarKickList_Message
End Property
Public Property Let Profanity_KickMessage(ByVal vData As String)
      mvarProfanity_Message = vData
End Property
Public Property Get Profanity_KickMessage() As String
      Profanity_KickMessage = mvarProfanity_Message
End Property
Public Property Let Profanity_NoBan(ByVal vData As Boolean)
      mvarProfanity_NoBan = vData
End Property
Public Property Get Profanity_NoBan() As Boolean
      Profanity_NoBan = mvarProfanity_NoBan
End Property
Public Property Let Profanity_Ban(ByVal vData As Boolean)
      mvarProfanity_Ban = vData
End Property
Public Property Get Profanity_Ban() As Boolean
      Profanity_Ban = mvarProfanity_Ban
End Property
Public Property Let Profanity_BanTime(ByVal vData As Integer)
      mvarProfanity_BanTime = vData
End Property
Public Property Get Profanity_BanTime() As Integer
      Profanity_BanTime = mvarProfanity_BanTime
End Property
Public Property Let Profanity_Active(ByVal vData As Boolean)
      mvarProfanity_Active = vData
End Property
Public Property Get Profanity_Active() As Boolean
      Profanity_Active = mvarProfanity_Active
End Property
Public Property Let Advertise_KickMessage(ByVal vData As String)
      mvarAdvertise_Message = vData
End Property
Public Property Get Advertise_KickMessage() As String
      Advertise_KickMessage = mvarAdvertise_Message
End Property
Public Property Let Advertise_BanTime(ByVal vData As Integer)
      mvarAdvertise_BanTime = vData
End Property
Public Property Get Advertise_BanTime() As Integer
      Advertise_BanTime = mvarAdvertise_BanTime
End Property
Public Property Let Advertise_Ban(ByVal vData As Boolean)
      mvarAdvertise_Ban = vData
End Property
Public Property Get Advertise_Ban() As Boolean
      Advertise_Ban = mvarAdvertise_Ban
End Property
Public Property Let Advertise_NoBan(ByVal vData As Boolean)
      mvarAdvertise_NoBan = vData
End Property
Public Property Get Advertise_NoBan() As Boolean
      Advertise_NoBan = mvarAdvertise_NoBan
End Property
Public Property Let Advertise_Active(ByVal vData As Boolean)
      mvarAdvertise_Active = vData
End Property
Public Property Get Advertise_Active() As Boolean
      Advertise_Active = mvarAdvertise_Active
End Property
Public Property Let Caps_Tolerance(ByVal vData As eTolerance)
      mvarCaps_Tolerance = vData
End Property
Public Property Get Caps_Tolerance() As eTolerance
      Caps_Tolerance = mvarCaps_Tolerance
End Property
Public Property Let Caps_Active(ByVal vData As Boolean)
      mvarCaps_Active = vData
End Property
Public Property Get Caps_Active() As Boolean
      Caps_Active = mvarCaps_Active
End Property
Public Property Let Caps_Ban(ByVal vData As Boolean)
      mvarCaps_Ban = vData
End Property
Public Property Get Caps_Ban() As Boolean
      Caps_Ban = mvarCaps_Ban
End Property
Public Property Let Caps_BanTime(ByVal vData As Integer)
      mvarCaps_BanTime = vData
End Property
Public Property Get Caps_BanTime() As Integer
      Caps_BanTime = mvarCaps_BanTime
End Property
Public Property Let Scroll_Warning(ByVal vData As Boolean)
      mvarScroll_Warning = vData
End Property
Public Property Get Scroll_Warning() As Boolean
      Scroll_Warning = mvarScroll_Warning
End Property
Public Property Let Scroll_Tolerance(ByVal vData As eTolerance)
      mvarScroll_Tolerance = vData
End Property
Public Property Get Scroll_Tolerance() As eTolerance
      Scroll_Tolerance = mvarScroll_Tolerance
End Property
Public Property Let Scroll_NoBan(ByVal vData As Boolean)
      mvarScroll_NoBan = vData
End Property
Public Property Get Scroll_NoBan() As Boolean
      Scroll_NoBan = mvarScroll_NoBan
End Property
Public Property Let Scroll_Message(ByVal vData As String)
      mvarScroll_Message = vData
End Property
Public Property Get Scroll_Message() As String
      Scroll_Message = mvarScroll_Message
End Property
Public Property Let Scroll_KickMessage(ByVal vData As String)
      mvarScroll_KickMessage = vData
End Property
Public Property Get Scroll_KickMessage() As String
      Scroll_KickMessage = mvarScroll_KickMessage
End Property
Public Property Let Scroll_Active(ByVal vData As Boolean)
      mvarScroll_KickActive = vData
End Property
Public Property Get Scroll_Active() As Boolean
      Scroll_Active = mvarScroll_KickActive
End Property
Public Property Let Scroll_Kick(ByVal vData As Boolean)
      mvarScroll_Kick = vData
End Property
Public Property Get Scroll_Kick() As Boolean
      Scroll_Kick = mvarScroll_Kick
End Property
Public Property Let Scroll_BanTime(ByVal vData As Integer)
      mvarScroll_BanTime = vData
End Property
Public Property Get Scroll_BanTime() As Integer
      Scroll_BanTime = mvarScroll_BanTime
End Property
Public Property Let Caps_Message(ByVal vData As String)
      mvarCaps_Message = vData
End Property
Public Property Get Caps_Message() As String
      Caps_Message = mvarCaps_Message
End Property
Public Property Let Caps_KickMessage(ByVal vData As String)
      mvarCaps_KickMessage = vData
End Property
Public Property Get Caps_KickMessage() As String
      Caps_KickMessage = mvarCaps_KickMessage
End Property
