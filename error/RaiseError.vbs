On Error Resume Next

Function HelloError(flag)
    On Error GoTo 0
    If flag Then
        Err.Raise
    End If
End Function

HelloError true
If Err.Number <> 0 Then
    WScript.Echo Err.Number & ":" & Err.Source & ":" & Err.Description
End If
