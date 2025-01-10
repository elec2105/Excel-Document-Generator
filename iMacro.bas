Attribute VB_Name = "iMacro"
''''''''''''''''''''''''''''''''''''''
'Разработка макроса: @elec2105
''''''''''''''''''''''''''''''''''''''



Option Explicit
Option Private Module

Public AppWord As Word.Application, iWord As Word.Document



Sub CreateDoc()
Dim MyArray(), BasePath As String, iFolder As String, iTemplate As String
Dim tmpArray, tmpSTR As String, iRow As Long, iColl As Long, i As Long, j As Long, q As Long
Dim iExcel As Object
Dim ID As String, TextToReplace As String, Text As String
Dim Mark As String, MaxLen As Long

Application.ScreenUpdating = 0
On Error GoTo iEnd

iFolder = Range("FILE_WORD").Value: If Right(iFolder, 1) <> "\" Then iFolder = iFolder & "\"
iTemplate = Range("FILE_TEMPLATE").Value: If Right(iTemplate, 1) = ";" Then iTemplate = Left(iTemplate, Len(iTemplate) - 1)
BasePath = ThisWorkbook.Path & "\Result\": ' Call FolderCreateDel(BasePath)

With Sheets("data")
    iRow = .UsedRange.Row + .UsedRange.Rows.Count - 1: iColl = .UsedRange.Column + .UsedRange.Columns.Count - 1
    MyArray = .Range(.Cells(1, 1), .Cells(iRow, iColl)).Value
End With

'создаем скрытый объект Word
Set AppWord = CreateObject("Word.Application"): AppWord.Visible = False

'перебираем массив
For i = 2 To iRow
    If MyArray(i, 1) = "ok" Then
    
        'перебираем указанные word-шаблоны
        tmpArray = Split(MyArray(i, 3), ";")
        For q = 0 To UBound(tmpArray)
            tmpSTR = iFolder & tmpArray(q) & ".docx"
            If Len(Dir(tmpSTR)) > 0 Then
                Set iWord = AppWord.Documents.Open(tmpSTR, ReadOnly:=True)
                'делаем замену переменных
                For j = 4 To iColl
                    Call ExportWord(MyArray(1, j), MyArray(i, j))
                Next j
                
                iWord.SaveAs Filename:=BasePath & MyArray(i, 2) & " - " & tmpArray(q) & ".docx", FileFormat:=wdFormatXMLDocument
                iWord.Close False: Set iWord = Nothing
            End If
        'перебираем указанные excel-шаблоны
            tmpSTR = iFolder & tmpArray(q) & ".xlsx"
                If Len(Dir(tmpSTR)) > 0 Then
                    MaxLen = 200
                    ' Choose a character for Mark that is not in your data,
                    '  and is not a special char: ~?*
                    Mark = "^"
                    Set iExcel = Workbooks.Open(tmpSTR)
                    'делаем замену переменных
                     For j = 4 To iColl
                        'iExcel.Sheets(1).Cells.Replace MyArray(1, j), MyArray(i, j)
                        'Call ReplaceText(MyArray(1, j), MyArray(i, j))
                            ID = MyArray(1, j)
                            TextToReplace = MyArray(i, j)
                            If ID <> vbNullString Then
                                 Do
                                  Text = Left$(TextToReplace, MaxLen) & Mark
                                 ' Terminate the loop when all of TextToReplace has been processed
                                  If Text = Mark Then Text = vbNullString
                                     TextToReplace = Mid$(TextToReplace, MaxLen + 1)
                                     iExcel.Sheets(1).Cells.Replace _
                                     What:=ID, _
                                     Replacement:=Text, _
                                     LookAt:=xlPart, _
                                     SearchOrder:=xlByRows, _
                                     MatchCase:=False, _
                                     SearchFormat:=False, _
                                     ReplaceFormat:=False
                                     ID = Mark
                                     Loop Until Text = vbNullString
                                 End If
                    Next j
                    
                    iExcel.SaveAs Filename:=BasePath & MyArray(i, 2) & " - " & tmpArray(q) & ".xlsx" '".docx" ', FileFormat:=wdFormatXMLDocument
                    iExcel.Close False: Set iExcel = Nothing
                End If
        Next q
        'Erase tmpArray
    End If
Next i

AppWord.Quit: Set AppWord = Nothing
'Erase MyArray: BasePath = "": iFolder = "": iTemplate = ""

Application.ScreenUpdating = 1
MsgBox "Файлы сформированы.", vbInformation

Exit Sub

iEnd:
    AppWord.Quit: Set AppWord = Nothing
    'Erase MyArray: BasePath = "": iFolder = "": iTemplate = ""
    Application.ScreenUpdating = 1
    MsgBox "При обработке данных возникла ошибка.", vbCritical
End Sub

