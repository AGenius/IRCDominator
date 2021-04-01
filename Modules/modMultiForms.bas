Attribute VB_Name = "modMultiForms"
Option Explicit
Public frmWhisperForm() As New frmWhisper

Public Sub SetupFormArrays()
        ReDim frmWhisperForm(0)
End Sub
Public Sub ShowWhisperWindow(sNickName As String, bMinimized As Boolean)
Dim iArrayCount As Integer
Dim sCaption As String

    If sNickName <> "" Then
        If MultiFormSetup(frmWhisperForm, iArrayCount, sNickName) = False Then
            Load frmWhisperForm(iArrayCount)
            sCaption = Replace(frmWhisperForm(iArrayCount).Caption, "%n", ConvertFromUTF(TestNick(sNickName, False)))
            frmWhisperForm(iArrayCount).WindowIndex = iArrayCount
            frmWhisperForm(iArrayCount).Nickname = TestNick(sNickName, False)
            frmWhisperForm(iArrayCount).Caption = sCaption
            If bMinimized Then
               frmWhisperForm(iArrayCount).WindowState = vbMinimized
            Else
               frmWhisperForm(iArrayCount).WindowState = vbNormal
            End If
            frmWhisperForm(iArrayCount).Show
        End If
        DoEvents
    End If
End Sub
Public Function MultiFormSetup( _
        ByRef fThisForm, _
        ByRef iArrayCount As Integer, _
        ByRef sTitle As String) As Boolean
    
Dim bFound As Boolean

        iArrayCount = NextFormArray(fThisForm, sTitle, bFound)
        If bFound Then
            If fThisForm(iArrayCount).Visible = False Then
                fThisForm(iArrayCount).Show
                fThisForm(iArrayCount).ZOrder
            End If
            If fThisForm(iArrayCount).WindowState = vbMinimized Then
                fThisForm(iArrayCount).WindowState = vbNormal
                fThisForm(iArrayCount).ZOrder
            End If
            MultiFormSetup = True
        End If
      
End Function

Public Function NextFormArray( _
        ByRef fThisForm, _
        ByRef sTitle As String, _
        ByRef bFound As Boolean)

Dim i As Long
Dim iNewCount As Integer

        On Error Resume Next
      
        bFound = False
        iNewCount = UBound(fThisForm)    ' Find the maximum elements for the form array
        If sTitle <> "" Then
            For i = 1 To iNewCount
                Err.Clear
                If Locate(fThisForm(i).Caption, "[ " & sTitle & " ]") Then
                    If Err = 0 Then
                        
                        ' Form Already Loaded
                        
                        bFound = True
                        NextFormArray = i
                        Exit Function
                    End If
                End If
            Next
        End If
        For i = 1 To iNewCount
            If fThisForm(i).Name = "" Then
                
                ' This will generate an error if the form array element is no longer valid
                
                If Err = 91 And i < iNewCount Then
                    
                    ' Reuse this element
                    
                    NextFormArray = i
                    Exit Function
                End If
            End If
        Next
        i = iNewCount + 1
        ReDim Preserve fThisForm(i)
        NextFormArray = i
End Function

Public Sub CloseAllWindows(fThisForm)
Dim i As Integer
Dim fTemp As Form

        On Error Resume Next
        For i = 1 To UBound(fThisForm)
            Set fTemp = fThisForm(i)
            fTemp.Hide
            Unload fTemp
        Next
End Sub

Public Function FindWhisperWindow(fThisForm, sNickName As String) As Integer
Dim i As Integer
Dim fTemp As frmWhisper

        On Error Resume Next
        FindWhisperWindow = -1
        For i = 1 To UBound(fThisForm)
            Err.Clear
            Set fTemp = fThisForm(i)
            If UCase$(TestNick(sNickName, False)) = UCase$(fTemp.Nickname) Then
                If Err = 0 Then
                    FindWhisperWindow = i
                    Set fTemp = Nothing
                    Exit For
                End If
            End If
        Next
        Set fTemp = Nothing
End Function

Public Sub CloseDisplayWindow( _
        ByVal sNickName As String, _
        Optional bAuto As Boolean = False)

Dim iFrmIndex As Integer

        On Error GoTo Hell

        iFrmIndex = FindWhisperWindow(frmWhisperForm, sNickName)
        If iFrmIndex > -1 And frmWhisperForm(iFrmIndex).Tag = bAuto Then
            frmWhisperForm(iFrmIndex).Hide
            Unload frmWhisperForm(iFrmIndex)
            Set frmWhisperForm(iFrmIndex) = Nothing
        End If
        
Hell:
End Sub
