VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' **********************
' Private variables to store the settings made:
Private mvarPath As String
Private mvarKey As String
Private mvarSection As String
Private mvarDefault As String
Private mvarLastReturnCode As Long

' Declares for cIniFile:

' Profile String functions:
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' local variable(s) to hold property value(s)
Private mvarNullChar As String 'local copy

Public Property Let NullChar(ByVal vData As String)
' Used to signify the Key data is null
        mvarNullChar = vData
End Property

Public Property Get NullChar() As String
        NullChar = mvarNullChar
End Property

Property Get LastReturnCode() As Long
' Did the last call succeed?
' 0 if not!
        LastReturnCode = mvarLastReturnCode
End Property

Property Let Default(sDefault As String)
' What to return if something goes wrong:
        mvarDefault = sDefault
End Property

Property Get Default() As String
' What to return if something goes wrong:
        Default = mvarDefault
End Property

Property Let Path(sPath As String)
' The filename of the INI file:
        mvarPath = sPath
End Property

Property Get Path() As String
' The filename of the INI file:
        Path = mvarPath
End Property

Property Let Key(sKey As String)
' The KEY= bit to look for
        mvarKey = sKey
End Property

Property Get Key() As String
' The KEY= bit to look for
        Key = mvarKey
End Property

Property Let Section(sSection As String)
' The [SECTION] bit to look for
        mvarSection = sSection
End Property

Property Get Section() As String
' The [SECTION] bit to look for
        Section = mvarSection
End Property

Property Get Value() As String
' Get the value of the current Key within Section of Path
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer
        sBuf = Space$(255)
        iSize = Len(sBuf)
        iRetCode = GetPrivateProfileString(mvarSection, mvarKey, mvarDefault, sBuf, iSize, mvarPath)

        If (iSize > 0) Then
           Value = Left$(sBuf, iRetCode)
           If Value = mvarNullChar Then
              Value = ""
           End If
        Else
           Value = ""
        End If
End Property

Property Let Value(sValue As String)
' Set the value of the current Key within Section of Path
Dim iPos As Integer
' Strip chr$(0):
        iPos = InStr(sValue, Chr$(0))
        Do While iPos <> 0
           sValue = Left$(sValue, (iPos - 1)) & Mid$(sValue, (iPos + 1))
           iPos = InStr(sValue, Chr$(0))
        Loop
        If sValue = "" Or sValue = " " Then
           sValue = mvarNullChar
        End If

        mvarLastReturnCode = WritePrivateProfileString(mvarSection, mvarKey, sValue, mvarPath)
End Property

Public Sub DeleteValue()
        ' Delete the value at Key within Section of Path
        mvarLastReturnCode = WritePrivateProfileString(mvarSection, mvarKey, 0&, mvarPath)
End Sub

Public Sub DeleteSection()
        ' Delete the Section in Path
        mvarLastReturnCode = WritePrivateProfileString(mvarSection, 0&, 0&, mvarPath)
End Sub

Property Get INISection() As String
' Return all the keys and values within the current
' section, separated by chr$(0):
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer

        sBuf = Space$(255)
        iSize = Len(sBuf)
        iRetCode = GetPrivateProfileString(mvarSection, 0&, mvarDefault, sBuf, iSize, mvarPath)

        If (iSize > 0) Then
           INISection = Left$(sBuf, iRetCode)
        Else
           INISection = ""
        End If
End Property

Property Let INISection(sSection As String)
' Set one or more the keys within the current section.
' Keys and Values should be separated by chr$(0):
        mvarLastReturnCode = WritePrivateProfileString(mvarSection, 0&, sSection, mvarPath)
End Property

Private Sub Class_Initialize()
mvarNullChar = "_"
End Sub
