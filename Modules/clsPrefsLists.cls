VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrefsLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum eListType
   lOwnerList = 0
   lHostList = 1
End Enum
Dim sNames(1) As String

Private mvarList_Owners_Active As Boolean
Private mvarList_Hosts_Active As Boolean
Private mvarList_Owner_Names As Collection
Private mvarList_Hosts_Name As Collection
Public Function IsInList(ByVal sNickName As String, ByVal lListType As eListType) As Boolean
Dim sNick As String
                        
      On Error Resume Next
      
      Select Case lListType
            
         Case eListType.lOwnerList
            If mvarList_Owner_Names.Count > 0 Then
               sNick = mvarList_Owner_Names.Item(sNickName)
               If sNick <> "" Then
                  IsInList = True
               End If
            End If
         Case eListType.lHostList
            If mvarList_Hosts_Name.Count > 0 Then
               sNick = mvarList_Hosts_Name.Item(sNickName)
               If sNick <> "" Then
                  IsInList = True
               End If
            End If
      End Select
End Function

Private Sub Class_Initialize()
      sNames(0) = "Owners.dat"
      sNames(1) = "Hosts.dat"
End Sub
Public Sub LoadPrefs()
Dim sSection As String

      sSection = "AutoHosting"
      
      mvarList_Owners_Active = CBool(fGetIni(sSection, "Owners_Active", "False"))
      mvarList_Hosts_Active = CBool(fGetIni(sSection, "Hosts_Active", "False"))
      Set mvarList_Owner_Names = LoadListToCollection(sNames(0))
      Set mvarList_Hosts_Name = LoadListToCollection(sNames(1))
End Sub

Public Sub SavePrefs()
Dim sSection As String

      sSection = "AutoHosting"
      
      Call PutIni(sSection, "Owners_Active", CStr(mvarList_Owners_Active))
      Call PutIni(sSection, "Hosts_Active", CStr(mvarList_Hosts_Active))
   
      Call SaveListFromCollection(mvarList_Owner_Names, sNames(0))
      Call SaveListFromCollection(mvarList_Hosts_Name, sNames(1))
End Sub
Public Sub SaveControl(ByRef cListControl As ListBox, ByVal lListType As eListType)
Dim i As Integer

      On Error Resume Next
      Select Case lListType
         Case eListType.lOwnerList
            For i = 1 To mvarList_Owner_Names.Count
               mvarList_Owner_Names.Remove (1)
            Next
            For i = 0 To cListControl.ListCount - 1
               mvarList_Owner_Names.Add cListControl.List(i), cListControl.List(i)
            Next
         Case eListType.lHostList
            For i = 1 To mvarList_Hosts_Name.Count
               mvarList_Hosts_Name.Remove (1)
            Next
            For i = 0 To cListControl.ListCount - 1
               mvarList_Hosts_Name.Add cListControl.List(i), cListControl.List(i)
            Next
      End Select
End Sub
Public Sub FillControl(ByRef cListControl As ListBox, ByVal lListType As eListType)
Dim i As Integer

      On Error GoTo ExitMe:
      cListControl.Clear
      Select Case lListType
         Case eListType.lOwnerList
            For i = 1 To mvarList_Owner_Names.Count
               cListControl.AddItem mvarList_Owner_Names.Item(i)
            Next
         Case eListType.lHostList
            For i = 1 To mvarList_Hosts_Name.Count
               cListControl.AddItem mvarList_Hosts_Name.Item(i)
            Next
      End Select
ExitMe:
End Sub
Public Sub RemoveFromList(ByVal sWord As String, ByVal lListType As eListType)

      On Error Resume Next
      
      Select Case lListType
         Case eListType.lOwnerList
            mvarList_Owner_Names.Remove sWord
         Case eListType.lHostList
            mvarList_Hosts_Name.Remove sWord
      End Select
End Sub
Public Sub ClearList(ByVal lListType As eListType)
      Select Case lListType
         Case eListType.lOwnerList
            Set mvarList_Owner_Names = Nothing
         Case eListType.lHostList
            Set mvarList_Hosts_Name = Nothing
      End Select
End Sub
Public Sub AddToList(ByVal sWord As String, ByVal lListType As eListType)

      On Error Resume Next
      
      Select Case lListType
         Case eListType.lOwnerList
            mvarList_Owner_Names.Add sWord, sWord
         Case eListType.lHostList
            mvarList_Hosts_Name.Add sWord, sWord
      End Select
End Sub
Public Property Let List_Hosts_Name(ByVal vData As String)
      Set mvarList_Hosts_Name = ConvertToCollection(vData)
End Property
Public Property Get List_Hosts_Name() As String
      List_Hosts_Name = ConvertFromCollection(mvarList_Hosts_Name)
End Property
Public Property Let List_Owner_Names(ByVal vData As String)
      Set mvarList_Owner_Names = ConvertToCollection(vData)
End Property
Public Property Get List_Owner_Names() As String
      List_Owner_Names = ConvertFromCollection(mvarList_Owner_Names)
End Property
Public Property Let List_Hosts_Active(ByVal vData As Boolean)
      mvarList_Hosts_Active = vData
End Property
Public Property Get List_Hosts_Active() As Boolean
      List_Hosts_Active = mvarList_Hosts_Active
End Property
Public Property Let List_Owners_Active(ByVal vData As Boolean)
      mvarList_Owners_Active = vData
End Property
Public Property Get List_Owners_Active() As Boolean
      List_Owners_Active = mvarList_Owners_Active
End Property
