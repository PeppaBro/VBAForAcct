Public Function Str2Num(ByVal Str As String)
    Dim i%, s$, c$
    For i = 1 To Len(Str)
        c = Mid(Str, i, 1)
        If (c Like "#") Or (Asc(c) = 46) Then s = s & c
    Next
    Str2Num = Val(s)
End Function
