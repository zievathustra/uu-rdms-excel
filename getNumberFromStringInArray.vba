Public Function getNumberFromStringInArray(sArrayofStrings As String, sSearchFor As String) As Integer

Dim arrToProcess As Variant
Dim arrToProcess2 as variant
Dim arrItem As String
dim arrItem2 as String

Dim i,j As Integer
j=2

arrToStrip = Split(sArrayofStrings, ",")
For i = 0 To UBound(arrToProcess)
    arrItem = arrToProcess(i)
    If InStr(arrItem, sSearchFor) >= 1 Then
        arrToProcess2 = Split(arrItem, ",")
        arrItem2 = arrToProcess2(j)
    Else
        arrItem2 = "0"
Next i

getNumberFromStringInArray = Int(arrItem2)

End Function
