Function Max(vArray)
    Dim vArrayLen
    Dim vR

    If Not IsArray(vArray) Then
        Err.Raise
    End If

    vR =
    vArrayLen = UBound(vArray)
    If vArrayLen < 0 Then
        Max =
        Exit Function
    End If
End Function
