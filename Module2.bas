Attribute VB_Name = "Module2"
Function GetFirstLetters(rng As Range) As String
'Update 20140325
    Dim arr
    Dim i As Long
    arr = VBA.Split(rng, " ")
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            GetFirstLetters = GetFirstLetters & Left(arr(i), 1)
        Next i
    Else
        GetFirstLetters = Left(arr, 1)
    End If
End Function
