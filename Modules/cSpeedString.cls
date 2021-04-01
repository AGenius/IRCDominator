VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpeedString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------------------------------------------------------
' SpeedString Class
' Written by Robin Schuil
' E-mail: robin@ykoon.nl
'
' For non-commercial usage only! Contact the author for information about commercial
' licenses.
' ---------------------------------------------------------------------------------

' API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private byteArray() As Byte     ' Store string in an ANSI byte array
Private hiByte As Long
Private hiBound As Long

Private Sub Class_Initialize()
      hiByte = 0
      hiBound = 1024              ' Default buffer size = 1024 bytes (1 Kb)
      ReDim byteArray(hiBound)
End Sub

Public Sub Append(ByRef StringData As String, Optional Length As Long)
Dim DataLength As Long
      If Length > 0 Then
         DataLength = Length
      Else
         DataLength = Len(StringData)
      End If
      If DataLength + hiByte > hiBound Then
         hiBound = hiBound + 1024
         ReDim Preserve byteArray(hiBound)
      End If
      CopyMemory ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
      hiByte = hiByte + DataLength
End Sub

Public Property Get Data() As String
Dim StringData As String
      StringData = Space(hiByte)
      CopyMemory ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
      Data = StringData
End Property

Public Sub Reset()
      Call Class_Initialize
End Sub
