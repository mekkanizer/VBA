Sub eyecandy()
        '������� ����
   Dim SelectedItem
        '������� ������� �����
   Dim wb As Workbook
       
        '�������� ������ ������ ������
   With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "�������� ����� �������"    '������� � ���� �������
       '���� �� ��������� � ����� ��� ���������� �������� ����, ����� ��������
       .InitialFileName = ThisWorkbook.Path & Application.PathSeparator & "*.csv"
        .AllowMultiSelect = True    '����� ���������� ������ ��������
       If .Show = False Then Exit Sub
 
        Application.ScreenUpdating = False
        For Each SelectedItem In .SelectedItems    '������� ������ � �����
           '��������� �����
           Workbooks.OpenText _
                Filename:=SelectedItem, _
                Origin:=xlWindows, _
                StartRow:=1, _
                DataType:=xlDelimited, _
                TextQualifier:=xlTextQualifierNone, _
                ConsecutiveDelimiter:=False, _
                Semicolon:=True, _
                Local:=True
            Set wb = ActiveWorkbook
            With wb.Worksheets(1)
                '�������� ������ ���������� ����������
               If wb.Name = "������� - ������������� �����.csv" Then _
                    .Rows(2).Delete
                If wb.Name = "������� - � ������.csv" Then _
                    .Cells(2, 4).Delete Shift:=xlShiftToLeft
                '��������� ������ ��������
               .Columns.AutoFit
                With .UsedRange
                '�������� �����
               .Borders.LineStyle = xlContinuous
                '�������� ��������� ��������
               .Rows(1).Font.Bold = True
                '�������� ����� � ������ ����
               .Rows(1).Borders.Weight = xlThick
				'��������� ������-������� (��� �� ��������? ������ ���!)
			   .Replace _
                    What:=",", Replacement:=".", _
					SearchOrder:=xlByColumns, MatchCase:=True
			   '������ �������� ������ ����� ���������
               .Replace _
                    What:=Chr(160), Replacement:="", _
                    SearchOrder:=xlByColumns, MatchCase:=True
				'������ ��� �� ����� ����
			   .Replace _
                    What:=".", Replacement:=",", _
                    SearchOrder:=xlByColumns, MatchCase:=True
                End With
            End With
            wb.SaveAs Filename:=Replace(wb.FullName, ".csv", ".xls"), FileFormat:=56
            wb.Close SaveChanges:=False
        Next SelectedItem
        Application.ScreenUpdating = True
    End With
End Sub

