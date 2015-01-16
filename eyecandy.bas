Sub eyecandy()
        'текущий файл
   Dim SelectedItem
        'текущая рабочая книга
   Dim wb As Workbook
   Dim falsespace As Range
       
        'вызываем диалог выбора файлов
   With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Выберите файлы отчетов"    'надпись в окне диалога
       'путь по умолчанию к папке где расположен исходный файл, можно изменить
       .InitialFileName = ThisWorkbook.Path & Application.PathSeparator & "*.csv"
        .AllowMultiSelect = True    'выбор нескольких файлов разрешён
       If .Show = False Then Exit Sub
 
        Application.ScreenUpdating = False
        For Each SelectedItem In .SelectedItems    'перебор файлов в папке
           'открываем книгу
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
                'починить косяки делфийских криворучек
               If wb.Name = "Местная - неоплачиваемый лимит.csv" Then _
                    .Rows(2).Delete
               If wb.Name = "Местная - к оплате.csv" Then _
                    .Cells(2, 4).Delete Shift:=xlShiftToLeft
                'поправить ширину столбцов
               .Columns.AutoFit
                With .UsedRange
                    'запилить рамки
                    .Borders.LineStyle = xlContinuous
                    'выделить заголовки столбцов
                    .Rows(1).Font.Bold = True
                    'выделить рамки в первом ряду
                    .Rows(1).Borders.Weight = xlThick
                End With
                'I fucking won this war, fuck you excel and ibexpert
                If Range("F1").Value = "ITOG" Then
                    With .Range("F:G")
                        .NumberFormat = "0.00"
                        .Replace What:=Chr(160), Replacement:=""
                    End With
                    For Each curr In .Range("F:G")
                        If (IsNumeric(curr.Value)) And (curr.Value > 999999) Then
                            curr.Value = curr.Value / 1000
                        End If
                    Next curr
                End If
            End With
            wb.SaveAs Filename:=Replace(wb.FullName, ".csv", ".xls"), FileFormat:=56
            wb.Close SaveChanges:=False
        Next SelectedItem
        Application.ScreenUpdating = True
    End With
End Sub

