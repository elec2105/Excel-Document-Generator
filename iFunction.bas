Attribute VB_Name = "iFunction"
''''''''''''''''''''''''''''''''''''''
'���������� �������: @elec2105
''''''''''''''''''''''''''''''''''''''



Option Explicit
Option Private Module


Function ExportWord(ByVal iName As String, ByVal iVal As String) As Boolean
    Dim i As Long
    
    '������������ ������ ������ � �������� ���������
metka1:
    With iWord.Content.Find                'With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        '.Replacement.Style = ActiveDocument.Styles("��������� 2 ����")  '���������� ��� ������������� ����� �������.
        .Text = iName
             If Len(iVal) > 255 Then '��������� Find.ReplaceText �� ����� ��������� ������ ������ 255 ��������, ' �������� � "�����" ����������� ������ �� ��������, ������ ��� �������� � ��� iName, ' ����� � ���������� �� �������� �����, ���� ��������� "�������" ������� ������
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
            .Replacement.Text = iVal '����� �� ������� ������
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
    '������������ ������ ������ � ������������
    For i = 6 To 11
        On Error Resume Next
        With iWord.StoryRanges.Item(i).Find
            .ClearFormatting
            .Replacement.ClearFormatting
            '.Replacement.Style = ActiveDocument.Styles("��������� 2 ����")  '���������� ��� ������������� ����� �������.
            .Text = iName
            .Replacement.Text = iVal
            .Forward = True                     'True, ���� ������ ����������� ������ "�����". False - � ��������� ������ ("�����").
            .Wrap = 1                           '���������� ��� ������������� ��������� ������������ WdFindWrap. wdFindContinue = 1
            .Format = False                     'True ���� � �������� ������ �������� ��������������, False - � ��������� ������.
            .MatchCase = True                   'True, ���� � �������� ������ ������� ��������� ������� ��������.
            .MatchWholeWord = True              'True, ���� � �������� ������ ������� ������ �������� ����� ��� ��������� �����, � �� ��� ����� ������� �����.
            .MatchAllWordForms = False          'True, ���� ��������� ����� ��� ���������� ��� ��������� �����.
            .MatchSoundsLike = False            'True, ���� ��������� ����� ����� ������� �� �������� �� �������� �����.
            .MatchWildcards = False             'True, ���� � �������� ������ ������������ ���������� ���������.
            
            'MatchByte                          'True, ���� � �������� ������ ������� ��������� ������� ������ � ���������� ������.
            'ParagraphFormat                    '���������� ��� ������������� ������ ParagraphFormat.
            'Found                              'True, ���� � ���������� ���������� ������ ���� ������� ������������.
            'Font                               '���������� ��� ������������� ������ Font �������� �������������� ������.
            
            '.Execute Replace:=wdReplaceAll
            .Execute
            
            If .Found Then  '���������, ������� �� �������� � ��������� Word
                ExportWord = True           '�������� �������
                .Execute Replace:=2         'wdReplaceAll = 2
            Else
                ExportWord = False          '�������� �� �������
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
