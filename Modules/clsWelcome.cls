VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Private Enum eCols
' CBlack = 0
' CWhite = 15
' CMaroon = 4
' CGreen = 2
' CNavy = 1
' COlive = 6
' CPurple = 5
' CTeal = 3
' CSilver = 7
' CGray = 8
' CRed = 12
' CLime = 10
' CBlue = 9
' CYellow = 14
' CFuschia = 13
' CAqua = 11
' End Enum


Private mvarWelcomeJoins_Active As Boolean
Private mvarWelcomeJoins_Message As String
Private mvarWelcomeJoins_Colour As eCols
Private mvarWelcomeJoins_FontName As String
Private mvarWelcomeJoins_FontBold As Boolean
Private mvarWelcomeJoins_FontItalic As Boolean
Private mvarWelcomeJoins_ViaWhisper As Boolean
Private mvarWelcomeJoins_ViaMainScreen As Boolean
Private mvarWelcomeBack_Active As Boolean
Private mvarWelcomeBack_Mesage As String
Private mvarWelcomeBack_Colour As eCols
Private mvarWelcomeBack_FontName As String
Private mvarWelcomeBack_FontBold As Boolean
Private mvarWelcomeBack_FontItalic As Boolean
Private mvarWelcomeBack_ViaOwnScreen As Boolean
Private mvarWelcomeJoins_ViaOwnScreen As Boolean
Private mvarWelcomeBack_ViaMainScreen As Boolean
Public Sub SavePrefs()

End Sub

Public Sub LoadPrefs()
Dim iTmp As Integer
Dim sTmp As String
Dim btmp As Boolean
Dim sSection As String

      sSection = "Welcome"
      mvarWelcomeJoins_Active = CBool(fGetIni(sSection, "Welcome_Active", "False"))
      mvarWelcomeJoins_Message = CStr(fGetIni(sSection, "Welcome_Message", "Welcome %n"))
      mvarWelcomeJoins_Colour = CInt(fGetIni(sSection, "Welcome_Colour", "1"))
      mvarWelcomeJoins_FontName = CStr(fGetIni(sSection, "Welcome_FontName", "Arial"))
      mvarWelcomeJoins_FontBold = CBool(fGetIni(sSection, "Welcome_FontBold", "False"))
      mvarWelcomeJoins_FontItalic = CBool(fGetIni(sSection, "Welcome_FontItalic", "False"))
      mvarWelcomeJoins_ViaWhisper = CBool(fGetIni(sSection, "Welcome_ViaWhisper", "True"))
      mvarWelcomeJoins_ViaMainScreen = CBool(fGetIni(sSection, "Welcome_ViaMainScreen", "False"))
      mvarWelcomeJoins_ViaOwnScreen = CBool(fGetIni(sSection, "Welcome_ViaOwnScreen", "False"))
      
      mvarWelcomeBack_Active = CBool(fGetIni(sSection, "Away_Active", "False"))
      mvarWelcomeBack_Mesage = CStr(fGetIni(sSection, "Away_Message", "Welcome Back %n"))
      mvarWelcomeBack_Colour = CInt(fGetIni(sSection, "Away_Colour", "1"))
      mvarWelcomeBack_FontName = CStr(fGetIni(sSection, "Away_FontName", "Arial"))
      mvarWelcomeBack_FontBold = CBool(fGetIni(sSection, "Away_FontBold", "False"))
      mvarWelcomeBack_FontItalic = CBool(fGetIni(sSection, "Away_FontItalic", "False"))
      mvarWelcomeBack_ViaOwnScreen = CBool(fGetIni(sSection, "Away_ViaOwnScreen", "False"))
      mvarWelcomeBack_ViaMainScreen = CBool(fGetIni(sSection, "Away_ViaMainScreen", "True"))

End Sub


Public Property Let WelcomeBack_ViaMainScreen(ByVal vData As Boolean)
      mvarWelcomeBack_ViaMainScreen = vData
End Property
Public Property Get WelcomeBack_ViaMainScreen() As Boolean
      WelcomeBack_ViaMainScreen = mvarWelcomeBack_ViaMainScreen
End Property
Public Property Let WelcomeJoins_ViaOwnScreen(ByVal vData As Boolean)
      mvarWelcomeJoins_ViaOwnScreen = vData
End Property
Public Property Get WelcomeJoins_ViaOwnScreen() As Boolean
      WelcomeJoins_ViaOwnScreen = mvarWelcomeJoins_ViaOwnScreen
End Property
Public Property Let WelcomeBack_ViaOwnScreen(ByVal vData As Boolean)
      mvarWelcomeBack_ViaOwnScreen = vData
End Property
Public Property Get WelcomeBack_ViaOwnScreen() As Boolean
      WelcomeBack_ViaOwnScreen = mvarWelcomeBack_ViaOwnScreen
End Property
Public Property Let WelcomeBack_FontItalic(ByVal vData As Boolean)
      mvarWelcomeBack_FontItalic = vData
End Property
Public Property Get WelcomeBack_FontItalic() As Boolean
      WelcomeBack_FontItalic = mvarWelcomeBack_FontItalic
End Property
Public Property Let WelcomeBack_FontBold(ByVal vData As Boolean)
      mvarWelcomeBack_FontBold = vData
End Property
Public Property Get WelcomeBack_FontBold() As Boolean
      WelcomeBack_FontBold = mvarWelcomeBack_FontBold
End Property
Public Property Let WelcomeBack_FontName(ByVal vData As String)
      mvarWelcomeBack_FontName = vData
End Property
Public Property Get WelcomeBack_FontName() As String
      WelcomeBack_FontName = mvarWelcomeBack_FontName
End Property
Public Property Let WelcomeBack_Colour(ByVal vData As eCols)
      mvarWelcomeBack_Colour = vData
End Property
Public Property Get WelcomeBack_Colour() As eCols
      WelcomeBack_Colour = mvarWelcomeBack_Colour
End Property
Public Property Let WelcomeBack_Mesage(ByVal vData As String)
      mvarWelcomeBack_Mesage = vData
End Property
Public Property Get WelcomeBack_Mesage() As String
      WelcomeBack_Mesage = mvarWelcomeBack_Mesage
End Property
Public Property Let WelcomeBack_Active(ByVal vData As Boolean)
      mvarWelcomeBack_Active = vData
End Property
Public Property Get WelcomeBack_Active() As Boolean
      WelcomeBack_Active = mvarWelcomeBack_Active
End Property
Public Property Let WelcomeJoins_ViaMainScreen(ByVal vData As Boolean)
      mvarWelcomeJoins_ViaMainScreen = vData
End Property
Public Property Get WelcomeJoins_ViaMainScreen() As Boolean
      WelcomeJoins_ViaMainScreen = mvarWelcomeJoins_ViaMainScreen
End Property
Public Property Let WelcomeJoins_ViaWhisper(ByVal vData As Boolean)
      mvarWelcomeJoins_ViaWhisper = vData
End Property
Public Property Get WelcomeJoins_ViaWhisper() As Boolean
      WelcomeJoins_ViaWhisper = mvarWelcomeJoins_ViaWhisper
End Property
Public Property Let WelcomeJoins_FontItalic(ByVal vData As Boolean)
      mvarWelcomeJoins_FontItalic = vData
End Property
Public Property Get WelcomeJoins_FontItalic() As Boolean
      WelcomeJoins_FontItalic = mvarWelcomeJoins_FontItalic
End Property
Public Property Let WelcomeJoins_FontBold(ByVal vData As Boolean)
      mvarWelcomeJoins_FontBold = vData
End Property
Public Property Get WelcomeJoins_FontBold() As Boolean
      WelcomeJoins_FontBold = mvarWelcomeJoins_FontBold
End Property
Public Property Let WelcomeJoins_FontName(ByVal vData As String)
      mvarWelcomeJoins_FontName = vData
End Property
Public Property Get WelcomeJoins_FontName() As String
      WelcomeJoins_FontName = mvarWelcomeJoins_FontName
End Property
Public Property Let WelcomeJoins_Colour(ByVal vData As eCols)
      mvarWelcomeJoins_Colour = vData
End Property
Public Property Get WelcomeJoins_Colour() As eCols
      WelcomeJoins_Colour = mvarWelcomeJoins_Colour
End Property
Public Property Let WelcomeJoins_Message(ByVal vData As String)
      mvarWelcomeJoins_Message = vData
End Property
Public Property Get WelcomeJoins_Message() As String
      WelcomeJoins_Message = mvarWelcomeJoins_Message
End Property
Public Property Let WelcomeJoins_Active(ByVal vData As Boolean)
      mvarWelcomeJoins_Active = vData
End Property
Public Property Get WelcomeJoins_Active() As Boolean
      WelcomeJoins_Active = mvarWelcomeJoins_Active
End Property
