Attribute VB_Name = "Demo"
Option Explicit

Private slc As StateLossCallback

Sub Demo1()
    Static coll As New Collection
    Dim i As Long
    '
    On Error Resume Next
    For i = 1 To 5
        coll.Item "Test" & i
        If Err.Number <> 0 Then
            With New StateLossCallback
                coll.Add .Self, "Test" & i
                .Self.InitByMacroName "Test" & i
            End With
            Err.Clear
        End If
    Next i
    On Error GoTo 0
End Sub
Sub Demo2()
    Static stateTracker As StateLossCallback
    If stateTracker Is Nothing Then
        Set stateTracker = New StateLossCallback
        stateTracker.InitByMacroName "Test6", "Arg1", 2
    End If
End Sub
Sub Demo3()
    If slc Is Nothing Then
        Set slc = New StateLossCallback
        slc.InitByAddress AddressOf TestX, "Testing"
    End If
End Sub

Private Sub Test1()
    Debug.Print "State was lost. Clean proc 1"
End Sub
Private Sub Test2()
    Debug.Print "State was lost. Clean proc 2"
    End
End Sub
Private Sub Test3()
    Debug.Print "State was lost. Clean proc 3"
    Stop
End Sub
Private Sub Test6(ByVal s As String, ByVal x As Long)
    Debug.Print "State was lost. Clean proc 6", "Args: " & s & ", " & x
    Err.Raise 5
End Sub
Private Sub TestX(ByVal instancePtr As LongPtr, ByVal argText As String)
    Debug.Print "State was lost. Clean proc X", argText
    Err.Raise 5
End Sub
