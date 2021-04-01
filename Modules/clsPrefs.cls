VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrefsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eChatSize
   szSmall = 0
   szMedium = 1
   zsLarge = 2
End Enum
Private mvarPlaySounds As Boolean
Private mvarNotify_Aways As Boolean
Private mvarNotify_Leaves As Boolean
Private mvarNotify_Joins As Boolean
Private mvarNoWhispers As Boolean
Private mvarWhisper_Response As String
Private mvarWhisper_Window As Boolean
Private mvarNoFormatting As Boolean
Private mvarChat_StyleItalic As Boolean
Private mvarChat_StyleBold As Boolean
Private mvarChat_Font As String
Private mvarChat_Colour As eCols
Private mvarChat_DisplaySize As eChatSize
Private mvarAutoJoinKick As Boolean
Private mvarMaskLocalTime As Boolean
Private mvarLocalTime As String
Private mvarWhisper_Notice As Boolean
Private mvarWhisper_Message As Boolean
Private mvarWhisper_Font As String
Private mvarWhisper_StyleBold As Boolean
Private mvarWhisper_StyleItalic As Boolean
Private mvarWhisper_Colour As eCols
Private mvarWhisper_Notify As Boolean
Private mvarWhisper_PrivMessage As Boolean
Private mvarWhisper_Whisper As Boolean
Private mvarWhisper_NoWhispers As Boolean
Private mvarTestAlive As Boolean
Private mvarAliveTime As Integer
Private mvarTryJoin As Boolean
Private mvarRejoinTimer As Integer
Private mvarNotifyTrace As Boolean
Private mvarAutoJoin As Boolean
Private mvarLastRoom As String
Private mvarLastNick As String
Private mvarHexRoom As Boolean
Private mvarWelcome_Active As Boolean
Private mvarWelcome_FontItalic As Boolean
Private mvarWelcome_FontBold As Boolean
Private mvarWelcome_Font As String
Private mvarMessage_Message As String
Private mvarWelcome_Colour As eCols
Private mvarAway_Active As Boolean
Private mvarAway_Message As String
Private mvarAway_Font As String
Private mvarAway_FontBold As Boolean
Private mvarAway_FontItalic As Boolean
Private mvarAway_Colour As eCols
Private mvarWelcome_RoomMessage As Boolean
Private mvarWelcome_PrivateMessage As Boolean
Private mvarWelcome_Whisper As Boolean
Private mvarAway_PrivateMessage As Boolean
Private mvarAway_RoomMessage As Boolean
Private mvarServerIP As String
Private mvarChatOCXVersion As String
Private mvarChatCLASSID As String
Private mvarToolTips As Boolean
Private mvarShowOCXs As Boolean
Private mvarShowMOTD As Boolean
Private mvarChat_ChatX As Boolean

Public Property Let Chat_ChatX(ByVal vData As Boolean)
    mvarChat_ChatX = vData
End Property
Public Property Get Chat_ChatX() As Boolean
    Chat_ChatX = mvarChat_ChatX
End Property




Public Sub SavePrefs()
Dim sSection As String

      sSection = "General"

      Call PutIni(sSection, "AutoJoin", CStr(mvarAutoJoin))
      Call PutIni(sSection, "ShowTrace", CStr(mvarNotifyTrace))
      Call PutIni(sSection, "TryJoin", CStr(mvarTryJoin))
      Call PutIni(sSection, "LocalTime", CStr(mvarLocalTime))
      Call PutIni(sSection, "MaskLocalTime", CStr(mvarMaskLocalTime))
      Call PutIni(sSection, "AutoJoinKick", CStr(mvarAutoJoinKick))
      Call PutIni(sSection, "RejoinTimer", CStr(mvarRejoinTimer))
      Call PutIni(sSection, "TestAlive", CStr(mvarTestAlive))
      Call PutIni(sSection, "AliveTimer", CStr(mvarAliveTime))
      Call PutIni(sSection, "NotifyJoins", CStr(mvarNotify_Joins))
      Call PutIni(sSection, "NotifyLeaves", CStr(mvarNotify_Leaves))
      Call PutIni(sSection, "NotifyAways", CStr(mvarNotify_Aways))
      Call PutIni(sSection, "PlaySounds", CStr(mvarPlaySounds))
      Call PutIni(sSection, "NoFormatting", CStr(mvarNoFormatting))
      Call PutIni(sSection, "ChatDisplaySize", CStr(mvarChat_DisplaySize))
      Call PutIni(sSection, "ChatFont", CStr(mvarChat_Font))
      Call PutIni(sSection, "ChatStyleBold", CStr(mvarChat_StyleBold))
      Call PutIni(sSection, "ChatStyleItalic", CStr(mvarChat_StyleItalic))
      Call PutIni(sSection, "ChatColour", CStr(mvarChat_Colour))
      Call PutIni(sSection, "WhisperWindow", CStr(mvarWhisper_Window))
      Call PutIni(sSection, "WhisperNotify", CStr(mvarWhisper_Notify))
      Call PutIni(sSection, "WhisperFont", CStr(mvarWhisper_Font))
      Call PutIni(sSection, "WhisperStyleBold", CStr(mvarWhisper_StyleBold))
      Call PutIni(sSection, "WhisperStyleItalic", CStr(mvarWhisper_StyleItalic))
      Call PutIni(sSection, "WhisperPrivMessage", CStr(mvarWhisper_PrivMessage))
      Call PutIni(sSection, "WhisperWhisper", CStr(mvarWhisper_Whisper))
      Call PutIni(sSection, "WhisperResponse", CStr(mvarWhisper_Response))
      Call PutIni(sSection, "WhisperColour", CStr(mvarWhisper_Colour))
      Call PutIni(sSection, "WhisperMessage", CStr(mvarWhisper_Message))
      Call PutIni(sSection, "NoWhispers", CStr(mvarWhisper_NoWhispers))
      Call PutIni(sSection, "ServerIP", CStr(mvarServerIP))
      Call PutIni(sSection, "ChatOCXVersion", CStr(mvarChatOCXVersion))
      Call PutIni(sSection, "ChatCLSID", CStr(mvarChatCLASSID))
      Call PutIni(sSection, "ToolTips", CStr(mvarToolTips))
      Call PutIni(sSection, "ShowOCXs", CStr(mvarShowOCXs))
      Call PutIni(sSection, "ShowMOTD", CStr(mvarShowMOTD))
      Call PutIni(sSection, "UseChatX", CStr(mvarChat_ChatX))
End Sub

Public Sub LoadPrefs()
Dim iTmp As Integer
Dim sTmp As String
Dim btmp As Boolean
Dim sSection As String

      sSection = "General"

      btmp = CBool(fGetIni(sSection, "AutoJoin", "True")):            mvarAutoJoin = btmp
      btmp = CBool(fGetIni(sSection, "ShowTrace", "False")):          mvarNotifyTrace = btmp
      btmp = CBool(fGetIni(sSection, "TryJoin", "False")):            mvarTryJoin = btmp
      sTmp = CStr(fGetIni(sSection, "LocalTime", "My Time")):         mvarLocalTime = sTmp
      btmp = CBool(fGetIni(sSection, "MaskLocalTime", "False")):      mvarMaskLocalTime = btmp
      btmp = CBool(fGetIni(sSection, "AutoJoinKick", "False")):       mvarAutoJoinKick = btmp
      iTmp = CInt(fGetIni(sSection, "RejoinTimer", "10")):            mvarRejoinTimer = iTmp
      btmp = CBool(fGetIni(sSection, "TestAlive", "False")):          mvarTestAlive = btmp
      iTmp = CInt(fGetIni(sSection, "AliveTimer", "1")):              mvarAliveTime = iTmp
      btmp = CBool(fGetIni(sSection, "NotifyJoins", "True")):         mvarNotify_Joins = btmp
      btmp = CBool(fGetIni(sSection, "NotifyLeaves", "True")):        mvarNotify_Leaves = btmp
      btmp = CBool(fGetIni(sSection, "NotifyAways", "True")):         mvarNotify_Aways = btmp
      btmp = CBool(fGetIni(sSection, "PlaySounds", "True")):          mvarPlaySounds = btmp
      btmp = CBool(fGetIni(sSection, "NoFormatting", "False")):       mvarNoFormatting = btmp
      iTmp = CInt(fGetIni(sSection, "ChatDisplaySize", "1")):         mvarChat_DisplaySize = iTmp
      sTmp = CStr(fGetIni(sSection, "ChatFont", "Arial")):            mvarChat_Font = sTmp
      btmp = CBool(fGetIni(sSection, "ChatStyleBold", "False")):      mvarChat_StyleBold = btmp
      btmp = CBool(fGetIni(sSection, "ChatStyleItalic", "False")):    mvarChat_StyleItalic = btmp
      iTmp = CInt(fGetIni(sSection, "ChatColour", "1")):              mvarChat_Colour = iTmp
      btmp = CBool(fGetIni(sSection, "WhisperWindow", "True")):       mvarWhisper_Window = btmp
      btmp = CBool(fGetIni(sSection, "WhisperNotify", "True")):       mvarWhisper_Notify = btmp
      sTmp = CStr(fGetIni(sSection, "WhisperFont", "Arial")):         mvarWhisper_Font = sTmp
      btmp = CBool(fGetIni(sSection, "WhisperStyleBold", "False")):   mvarWhisper_StyleBold = btmp
      btmp = CBool(fGetIni(sSection, "WhisperStyleItalic", "False")): mvarWhisper_StyleItalic = btmp
      btmp = CBool(fGetIni(sSection, "WhisperPrivMessage", "False")): mvarWhisper_PrivMessage = btmp
      btmp = CBool(fGetIni(sSection, "WhisperWhisper", "True")):      mvarWhisper_Whisper = btmp
      sTmp = CStr(fGetIni(sSection, "WhisperResponse", "Not Here")):  mvarWhisper_Response = sTmp
      iTmp = CInt(fGetIni(sSection, "WhisperColour", "1")):           mvarWhisper_Colour = iTmp
      btmp = CBool(fGetIni(sSection, "WhisperMessage", "False")):     mvarWhisper_Message = btmp
      btmp = CBool(fGetIni(sSection, "NoWhispers", "False")):         mvarWhisper_NoWhispers = btmp
      sTmp = CStr(fGetIni(sSection, "ServerIP", "207.68.167.253")):   mvarServerIP = sTmp
      sTmp = CStr(fGetIni(sSection, "ChatOCXVersion", "8,0,211,1802")):          mvarChatOCXVersion = sTmp
      sTmp = CStr(fGetIni(sSection, "ChatCLSID", "F58E1CEF-A068-4C15-BA5E-587CAF3EE8C6")):    mvarChatCLASSID = sTmp
      btmp = CBool(fGetIni(sSection, "ToolTips", "True")):          mvarToolTips = btmp
      btmp = CBool(fGetIni(sSection, "ShowOCXs", "True")):          mvarShowOCXs = btmp
      btmp = CBool(fGetIni(sSection, "ShowMOTD", "True")):          mvarShowMOTD = btmp
      btmp = CBool(fGetIni(sSection, "UseChatX", "True")):          mvarChat_ChatX = btmp
End Sub
Public Property Let ShowMOTD(ByVal vData As Boolean)
      mvarShowMOTD = vData
End Property
Public Property Get ShowMOTD() As Boolean
      ShowMOTD = mvarShowMOTD
End Property

Public Property Let ShowOCXs(ByVal vData As Boolean)
      mvarShowOCXs = vData
End Property
Public Property Get ShowOCXs() As Boolean
      ShowOCXs = mvarShowOCXs
End Property

Public Property Let ShowToolTips(ByVal vData As Boolean)
      mvarToolTips = vData
End Property
Public Property Get ShowToolTips() As Boolean
      ShowToolTips = mvarToolTips
End Property
Public Property Let ChatCLASSID(ByVal vData As String)
    mvarChatCLASSID = vData
End Property
Public Property Get ChatCLASSID() As String
    ChatCLASSID = mvarChatCLASSID
End Property
Public Property Let ChatOCXVersion(ByVal vData As String)
    mvarChatOCXVersion = vData
End Property
Public Property Get ChatOCXVersion() As String
    ChatOCXVersion = mvarChatOCXVersion
End Property
Public Property Let ServerIP(ByVal vData As String)
    mvarServerIP = vData
End Property
Public Property Get ServerIP() As String
    ServerIP = mvarServerIP
End Property
Public Property Let Away_RoomMessage(ByVal vData As Boolean)
      mvarAway_RoomMessage = vData
End Property
Public Property Get Away_RoomMessage() As Boolean
      Away_RoomMessage = mvarAway_RoomMessage
End Property
Public Property Let Away_PrivateMessage(ByVal vData As Boolean)
      mvarAway_PrivateMessage = vData
End Property
Public Property Get Away_PrivateMessage() As Boolean
      Away_PrivateMessage = mvarAway_PrivateMessage
End Property
Public Property Let Welcome_Whisper(ByVal vData As Boolean)
      mvarWelcome_Whisper = vData
End Property
Public Property Get Welcome_Whisper() As Boolean
      Welcome_Whisper = mvarWelcome_Whisper
End Property
Public Property Let Welcome_PrivateMessage(ByVal vData As Boolean)
      mvarWelcome_PrivateMessage = vData
End Property
Public Property Get Welcome_PrivateMessage() As Boolean
      Welcome_PrivateMessage = mvarWelcome_PrivateMessage
End Property
Public Property Let Welcome_RoomMessage(ByVal vData As Boolean)
      mvarWelcome_RoomMessage = vData
End Property
Public Property Get Welcome_RoomMessage() As Boolean
      Welcome_RoomMessage = mvarWelcome_RoomMessage
End Property
Public Property Let Away_Colour(ByVal vData As eCols)
      mvarAway_Colour = vData
End Property
Public Property Get Away_Colour() As eCols
      Away_Colour = mvarAway_Colour
End Property
Public Property Let Away_FontItalic(ByVal vData As Boolean)
      mvarAway_FontItalic = vData
End Property
Public Property Get Away_FontItalic() As Boolean
      Away_FontItalic = mvarAway_FontItalic
End Property
Public Property Let Away_FontBold(ByVal vData As Boolean)
      mvarAway_FontBold = vData
End Property
Public Property Get Away_FontBold() As Boolean
      Away_FontBold = mvarAway_FontBold
End Property
Public Property Let Away_FontName(ByVal vData As String)
      mvarAway_Font = vData
End Property
Public Property Get Away_FontName() As String
      Away_FontName = mvarAway_Font
End Property
Public Property Let Away_Message(ByVal vData As String)
      mvarAway_Message = vData
End Property
Public Property Get Away_Message() As String
      Away_Message = mvarAway_Message
End Property
Public Property Let Away_Active(ByVal vData As Boolean)
      mvarAway_Active = vData
End Property
Public Property Get Away_Active() As Boolean
      Away_Active = mvarAway_Active
End Property
Public Property Let Welcome_Colour(ByVal vData As eCols)
      mvarWelcome_Colour = vData
End Property
Public Property Get Welcome_Colour() As eCols
      Welcome_Colour = mvarWelcome_Colour
End Property
Public Property Let MessageMessage(ByVal vData As String)
      mvarMessage_Message = vData
End Property
Public Property Get MessageMessage() As String
      MessageMessage = mvarMessage_Message
End Property
Public Property Let Welcome_Font(ByVal vData As String)
      mvarWelcome_Font = vData
End Property
Public Property Get Welcome_Font() As String
      Welcome_Font = mvarWelcome_Font
End Property
Public Property Let Welcome_FontBold(ByVal vData As Boolean)
      mvarWelcome_FontBold = vData
End Property
Public Property Get Welcome_FontBold() As Boolean
      Welcome_FontBold = mvarWelcome_FontBold
End Property
Public Property Let Welcome_FontItalic(ByVal vData As Boolean)
      mvarWelcome_FontItalic = vData
End Property
Public Property Get Welcome_FontItalic() As Boolean
      Welcome_FontItalic = mvarWelcome_FontItalic
End Property
Public Property Let Welcome_Active(ByVal vData As Boolean)
      mvarWelcome_Active = vData
End Property
Public Property Get Welcome_Active() As Boolean
      Welcome_Active = mvarWelcome_Active
End Property
Public Property Let Last_HexRoom(ByVal vData As Boolean)
      mvarHexRoom = vData
End Property
Public Property Get Last_HexRoom() As Boolean
      Last_HexRoom = mvarHexRoom
End Property
Public Property Let Last_LastNick(ByVal vData As String)
      mvarLastNick = vData
End Property
Public Property Get Last_LastNick() As String
      Last_LastNick = mvarLastNick
End Property
Public Property Let Last_LastRoom(ByVal vData As String)
      mvarLastRoom = vData
End Property
Public Property Get Last_LastRoom() As String
      Last_LastRoom = mvarLastRoom
End Property
Public Property Let AutoJoin(ByVal vData As Boolean)
      mvarAutoJoin = vData
End Property
Public Property Get AutoJoin() As Boolean
      AutoJoin = mvarAutoJoin
End Property
Public Property Let ShowTrace(ByVal vData As Boolean)
      mvarNotifyTrace = vData
End Property
Public Property Get ShowTrace() As Boolean
      ShowTrace = mvarNotifyTrace
End Property
Public Property Let RejoinTimer(ByVal vData As Integer)
      mvarRejoinTimer = vData
End Property
Public Property Get RejoinTimer() As Integer
      RejoinTimer = mvarRejoinTimer
End Property
Public Property Let TryJoin(ByVal vData As Boolean)
      mvarTryJoin = vData
End Property
Public Property Get TryJoin() As Boolean
      TryJoin = mvarTryJoin
End Property
Public Property Let AliveTime(ByVal vData As Integer)
      mvarAliveTime = vData
End Property
Public Property Get AliveTime() As Integer
      AliveTime = mvarAliveTime
End Property
Public Property Let TestAlive(ByVal vData As Boolean)
      mvarTestAlive = vData
End Property
Public Property Get TestAlive() As Boolean
      TestAlive = mvarTestAlive
End Property
Public Property Let Whisper_Whisper(ByVal vData As Boolean)
      mvarWhisper_Whisper = vData
End Property
Public Property Get Whisper_Whisper() As Boolean
      Whisper_Whisper = mvarWhisper_Whisper
End Property
Public Property Let Whisper_PrivMessage(ByVal vData As Boolean)
      mvarWhisper_PrivMessage = vData
End Property
Public Property Get Whisper_PrivMessage() As Boolean
      Whisper_PrivMessage = mvarWhisper_PrivMessage
End Property
Public Property Let Whisper_Notify(ByVal vData As Boolean)
      mvarWhisper_Notify = vData
End Property
Public Property Get Whisper_Notify() As Boolean
      Whisper_Notify = mvarWhisper_Notify
End Property
Public Property Let Whisper_Colour(ByVal vData As eCols)
      mvarWhisper_Colour = vData
End Property
Public Property Get Whisper_Colour() As eCols
      Whisper_Colour = mvarWhisper_Colour
End Property
Public Property Let Whisper_StyleItalic(ByVal vData As Boolean)
      mvarWhisper_StyleItalic = vData
End Property
Public Property Get Whisper_StyleItalic() As Boolean
      Whisper_StyleItalic = mvarWhisper_StyleItalic
End Property
Public Property Let Whisper_StyleBold(ByVal vData As Boolean)
      mvarWhisper_StyleBold = vData
End Property
Public Property Get Whisper_StyleBold() As Boolean
      Whisper_StyleBold = mvarWhisper_StyleBold
End Property
Public Property Let Whisper_Font(ByVal vData As String)
      mvarWhisper_Font = vData
End Property
Public Property Get Whisper_Font() As String
      Whisper_Font = mvarWhisper_Font
End Property
Public Property Let Whisper_Message(ByVal vData As Boolean)
      mvarWhisper_Message = vData
End Property
Public Property Get Whisper_Message() As Boolean
      Whisper_Message = mvarWhisper_Message
End Property
Public Property Let Whisper_Notice(ByVal vData As Boolean)
      mvarWhisper_Notice = vData
End Property
Public Property Get Whisper_Notice() As Boolean
      Whisper_Notice = mvarWhisper_Notice
End Property
Public Property Let LocalTime(ByVal vData As String)
      mvarLocalTime = vData
End Property
Public Property Get LocalTime() As String
      LocalTime = Replace(mvarLocalTime, vbCrLf, "")
End Property
Public Property Let MaskLocalTime(ByVal vData As Boolean)
      mvarMaskLocalTime = vData
End Property
Public Property Get MaskLocalTime() As Boolean
      MaskLocalTime = mvarMaskLocalTime
End Property
Public Property Let AutoJoinKick(ByVal vData As Boolean)
      mvarAutoJoinKick = vData
End Property
Public Property Get AutoJoinKick() As Boolean
      AutoJoinKick = mvarAutoJoinKick
End Property
Public Property Let Chat_DisplaySize(ByVal vData As eChatSize)
      mvarChat_DisplaySize = vData
End Property
Public Property Get Chat_DisplaySize() As eChatSize
      Chat_DisplaySize = mvarChat_DisplaySize
End Property
Public Property Let Chat_Colour(ByVal vData As eCols)
      mvarChat_Colour = vData
End Property
Public Property Get Chat_Colour() As eCols
      Chat_Colour = mvarChat_Colour
End Property
Public Property Let Chat_Font(ByVal vData As String)
      mvarChat_Font = vData
End Property
Public Property Get Chat_Font() As String
      Chat_Font = mvarChat_Font
End Property
Public Property Let Chat_StyleBold(ByVal vData As Boolean)
      mvarChat_StyleBold = vData
End Property
Public Property Get Chat_StyleBold() As Boolean
      Chat_StyleBold = mvarChat_StyleBold
End Property
Public Property Let Chat_StyleItalic(ByVal vData As Boolean)
      mvarChat_StyleItalic = vData
End Property
Public Property Get Chat_StyleItalic() As Boolean
      Chat_StyleItalic = mvarChat_StyleItalic
End Property
Public Property Let NoFormatting(ByVal vData As Boolean)
      mvarNoFormatting = vData
End Property
Public Property Get NoFormatting() As Boolean
      NoFormatting = mvarNoFormatting
End Property
Public Property Let Whisper_Window(ByVal vData As Boolean)
      mvarWhisper_Window = vData
End Property
Public Property Get Whisper_Window() As Boolean
      Whisper_Window = mvarWhisper_Window
End Property
Public Property Let Whisper_Response(ByVal vData As String)
      mvarWhisper_Response = vData
End Property
Public Property Get Whisper_Response() As String
      Whisper_Response = Replace(mvarWhisper_Response, vbCrLf, "")
End Property
Public Property Let NoWhispers(ByVal vData As Boolean)
      mvarNoWhispers = vData
End Property
Public Property Get NoWhispers() As Boolean
      NoWhispers = mvarNoWhispers
End Property
Public Property Let Notify_Joins(ByVal vData As Boolean)
      mvarNotify_Joins = vData
End Property
Public Property Get Notify_Joins() As Boolean
      Notify_Joins = mvarNotify_Joins
End Property
Public Property Let Notify_Leaves(ByVal vData As Boolean)
      mvarNotify_Leaves = vData
End Property
Public Property Get Notify_Leaves() As Boolean
      Notify_Leaves = mvarNotify_Leaves
End Property
Public Property Let Notify_Aways(ByVal vData As Boolean)
      mvarNotify_Aways = vData
End Property
Public Property Get Notify_Aways() As Boolean
      Notify_Aways = mvarNotify_Aways
End Property
Public Property Let PlaySounds(ByVal vData As Boolean)
      mvarPlaySounds = vData
End Property
Public Property Get PlaySounds() As Boolean
      PlaySounds = mvarPlaySounds
End Property
Public Property Let Whisper_NoWhispers(ByVal vData As Boolean)
      mvarWhisper_NoWhispers = vData
End Property
Public Property Get Whisper_NoWhispers() As Boolean
      Whisper_NoWhispers = mvarWhisper_NoWhispers
End Property

