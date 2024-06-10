Function CompareVersion(v1, v2)
    Dim v1a
    Dim v2a
    Dim minCount
    Dim i

    v1a = Split(v1, ".")
    v2a = Split(v2, ".")

    v1aLen = UBound(v1a)
    v2aLen = UBound(v2a)

    For i = 0 To v1a
    Next
End Function
