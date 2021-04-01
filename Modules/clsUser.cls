VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eHostType
    Owner = 1
    Host = 2
    Guest = 0
    Sysop = 3
    Spectator = 4
End Enum
Private mvarDisplayName As String
Private mvarRealName As String
Private mvarStatus As Integer
Private mvarCapsCounter As Integer
Private mvarScrollCounter As Integer
Private mvarScrollTime As String
Private mvarGapeKeeperID As String
Private mvarSigninTime As String
Private mvarAway As Boolean
Private mvarHostType As eHostType
Private mvarLastSentence As String
Public Property Let LastSentence(ByVal vData As String)
    mvarLastSentence = vData
End Property
Public Property Get LastSentence() As String
    LastSentence = mvarLastSentence
End Property
Public Property Let HostType(ByVal vData As eHostType)
    mvarHostType = vData
End Property
Public Property Get HostType() As eHostType
     HostType = mvarHostType
End Property
Public Property Let Away(ByVal vData As Boolean)
    mvarAway = vData
End Property
Public Property Get Away() As Boolean
    Away = mvarAway
End Property
Public Property Let SigninTime(ByVal vData As String)
    mvarSigninTime = vData
End Property
Public Property Get SigninTime() As String
    SigninTime = mvarSigninTime
End Property
Public Property Let GapeKeeperID(ByVal vData As String)
    mvarGapeKeeperID = vData
End Property
Public Property Get GapeKeeperID() As String
    GapeKeeperID = mvarGapeKeeperID
End Property
Public Property Let ScrollTime(ByVal vData As String)
    mvarScrollTime = vData
End Property
Public Property Get ScrollTime() As String
    ScrollTime = mvarScrollTime
End Property
Public Property Let ScrollCounter(ByVal vData As Integer)
    mvarScrollCounter = vData
End Property
Public Property Get ScrollCounter() As Integer
    ScrollCounter = mvarScrollCounter
End Property
Public Property Let CapsCounter(ByVal vData As Integer)
    mvarCapsCounter = vData
End Property
Public Property Get CapsCounter() As Integer
    CapsCounter = mvarCapsCounter
End Property
Public Property Let Status(ByVal vData As Integer)
    mvarStatus = vData
End Property
Public Property Get Status() As Integer
    Status = mvarStatus
End Property
Public Property Let RealName(ByVal vData As String)
  mvarRealName = vData
End Property
Public Property Get RealName() As String
    RealName = mvarRealName
End Property
Public Property Let DisplayName(ByVal vData As String)
    mvarDisplayName = vData
End Property
Public Property Get DisplayName() As String
    DisplayName = mvarDisplayName
End Property
