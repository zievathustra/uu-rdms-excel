Public Function getNumberFromStringInArray(sArrayofStrings As String, sSearchFor As String, sDelimiter1 As String, sDelimiter2 As String) As Integer

Dim arrToProcess As Variant
Dim arrToProcess2 As Variant
Dim arrItem As String
Dim arrItem2 As String

Dim i, j As Integer
j = 2

If sArrayofStrings <> "" Then
    arrToProcess = Split(sArrayofStrings, sDelimiter1)
    For i = 0 To UBound(arrToProcess)
        arrItem = arrToProcess(i)
        If InStr(arrItem, sSearchFor) >= 1 Then
            arrToProcess2 = Split(arrItem, sDelimiter2)
            arrItem2 = arrToProcess2(j)
            Exit For
        Else
            arrItem2 = "0"
        End If
    Next i
Else
    arrItem2 = "0"
End If

getNumberFromStringInArray = Int(arrItem2)

End Function
