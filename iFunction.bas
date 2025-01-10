Attribute VB_Name = "iFunction"
''''''''''''''''''''''''''''''''''''''
'Разработка макроса: @elec2105
''''''''''''''''''''''''''''''''''''''



Option Explicit
Option Private Module


Function ExportWord(ByVal iName As String, ByVal iVal As String) As Boolean
    Dim i As Long
    
    'осуществляем замену текста в основной документе
metka1:
    With iWord.Content.Find                'With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        '.Replacement.Style = ActiveDocument.Styles("Заголовок 2 Знак")  'Возвращает или устанавливает стиль объекта.
        .Text = iName
             If Len(iVal) > 255 Then 'поскольку Find.ReplaceText не может принимать строку больше 255 символов, ' пришлось в "цикле" подставлять строку по кусочкам, каждый раз добавляя в нее iName, ' чтобы в дальнейшем не потерять место, куда вставляем "хвостик" длинной строки
            .Replacement.Text = Left(iVal, 255 - Len(iName)) & iName
            iVal = Right(iVal, Len(iVal) - (255 - Len(iName)))
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=2
            GoTo metka1
             Else
            .Replacement.Text = iVal 'текст на который меняем
             End If
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=2
    End With
    'осуществляем замену текста в колонтитулах
    For i = 6 To 11
        On Error Resume Next
        With iWord.StoryRanges.Item(i).Find
            .ClearFormatting
            .Replacement.ClearFormatting
            '.Replacement.Style = ActiveDocument.Styles("Заголовок 2 Знак")  'Возвращает или устанавливает стиль объекта.
            .Text = iName
            .Replacement.Text = iVal
            .Forward = True                     'True, если задано направление поиска "Вперёд". False - в противном случае ("Назад").
            .Wrap = 1                           'Возвращает или устанавливает константу перечисления WdFindWrap. wdFindContinue = 1
            .Format = False                     'True если в операцию поиска включено форматирование, False - в противном случае.
            .MatchCase = True                   'True, если в процессе поиска следует различать регистр символов.
            .MatchWholeWord = True              'True, если в процессе поиска следует искать заданный текст как отдельное слово, а не как часть другого слова.
            .MatchAllWordForms = False          'True, если требуется найти все словоформы для заданного слова.
            .MatchSoundsLike = False            'True, если требуется найти слова похожие по звучанию на заданный текст.
            .MatchWildcards = False             'True, если в процессе поиска используются регулярные выражения.
            
            'MatchByte                          'True, если в процессе поиска следует различать символы полной и половинной ширины.
            'ParagraphFormat                    'Возвращает или устанавливает объект ParagraphFormat.
            'Found                              'True, если в результате выполнения поиска было найдено соответствие.
            'Font                               'Возвращает или устанавливает объект Font задающий форматирование шрифта.
            
            '.Execute Replace:=wdReplaceAll
            .Execute
            
            If .Found Then  'проверяем, найдена ли Закладка в документе Word
                ExportWord = True           'закладка найдена
                .Execute Replace:=2         'wdReplaceAll = 2
            Else
                ExportWord = False          'закладка НЕ найдена
            End If
       End With
       If Err.Number <> 0 Then Err.Clear
    Next i
End Function


Sub FolderCreateDel(ByVal iPath As String)
    Dim BasePath As String
    On Error Resume Next
    Kill iPath & "*.docx"
    MkDir (iPath)
End Sub
