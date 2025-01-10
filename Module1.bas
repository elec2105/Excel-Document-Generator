Attribute VB_Name = "Module1"
Function Translit(Txt As String) As String
 
    Dim Rus As Variant
    Rus = Array("à", "á", "â", "ã", "ä", "å", "¸", "æ", "ç", "è", "é", "ê", _
    "ë", "ì", "í", "î", "ï", "ð", "ñ", "ò", "ó", "ô", "õ", "ö", "÷", "ø", _
    "ù", "ú", "û", "ü", "ý", "þ", "ÿ", "À", "Á", "Â", "Ã", "Ä", "Å", _
    "¨", "Æ", "Ç", "È", "É", "Ê", "Ë", "Ì", "Í", "Î", "Ï", "Ð", _
    "Ñ", "Ò", "Ó", "Ô", "Õ", "Ö", "×", "Ø", "Ù", "Ú", "Û", "Ü", "Ý", "Þ", "ß")
 
    Dim Eng As Variant
    Eng = Array("a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "j", _
    "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", _
    "sh", "sch", "''", "y", "'", "e", "yu", "ya", "A", "B", "V", "G", "D", _
    "E", "JO", "ZH", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", _
    "S", "T", "U", "F", "KH", "TS", "CH", "SH", "SCH", "''", "Y", "'", "E", "YU", "YA")
     
    For i = 1 To Len(Txt)
        ñ = Mid(Txt, i, 1)
     
        flag = 0
        For j = 0 To 65
            If Rus(j) = ñ Then
                outchr = Eng(j)
                flag = 1
                Exit For
            End If
        Next j
        If flag Then outstr = outstr & outchr Else outstr = outstr & ñ
    Next i
     
    Translit = outstr
     
End Function

