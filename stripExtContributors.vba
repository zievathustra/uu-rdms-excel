Public Function stripExtContributors(sOld As String, sSearchFor As String) As String

Dim arrToStrip As Variant
Dim arrItem As String

Dim i As Integer

arrToStrip = Split(sOld, ",")
For i = 0 To UBound(arrToStrip)
    arrItem = arrToStrip(i)
    If InStr(arrItem, sSearchFor) >= 1 Then
        arrToStrip(i) = Replace(arrItem, arrItem, "")
    Else
        If i = 0 Then
            arrToStrip(i) = arrItem
        Else
            arrToStrip(i) = Trim(arrItem)
        End If
    End If
Next i

stripExtContributors = Join(arrToStrip, ", ")

End Function
