Option Explicit
 
Declare Function MakeSureDirectoryPathExists Lib "Imagehlp.dll" (ByVal strPath As String) As Long
'проверяет наличие папки с указанным путем и создает, если ее нет
'возвращает 0, если папку создать не удалось и не-0, если ОК
 
 
Sub tt()
    'текущий лист, текущий файл
    Dim sh As Object, SelectedItem
    'массив, элемент массива
    Dim a(), el
    'минимальный и максимальный номер в словаре
    Dim mi As Long, ma As Long
    'итератор кассиров
    Dim c As Integer
    'для считывания даты
    Dim d As String
    'путь к файлу
    Dim tfilepath As String
    'словарь, текущая рабочая книга
    Dim dic As Object, wb As Object
	'костыль, не дающий создать пустой словарь
	Dim hack As boolean
 
    'вызываем диалог выбора файлов
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Выберите файлы отчетов"    'надпись в окне диалога
        'путь по умолчанию к папке где расположен исходный файл, можно изменить
        .InitialFileName = ThisWorkbook.Path & Application.PathSeparator & "*.xls"
        .AllowMultiSelect = True    'выбор нескольких файлов разрешён
        If .Show = False Then Exit Sub
 
        Application.ScreenUpdating = False
        For Each SelectedItem In .SelectedItems    'перебор файлов в папке
            mi = 1000000000#: ma = 0
            Set dic = CreateObject("scripting.dictionary")
            Set wb = Workbooks.Open(SelectedItem)            'открываем книгу
            'операции с открытой книгой
            c = 0
            d = Mid(Cells(7, 2), 1, 10)
			hack = False
            tfilepath = wb.Path & "\import\"
            For Each sh In wb.Worksheets
                If sh.UsedRange.Columns.Count > 3 Then
                    a = sh.UsedRange.Columns(4).Value
                    For Each el In a
                        If el = "Номер" Then
							'пропускаем первое вхождение поля "номер"
                            If hack = False Then
								hack = True
							Else
								If mi <> 1000000000# Then
									c = c + 1
									vivod dic, mi, ma, d, c, tfilepath
									mi = 1000000000#: ma = 0
									Set dic = CreateObject("scripting.dictionary")
								End If
							End If
						End If
                        'пропускаем пустые строки
                        If IsNumeric(el) And el <> 0 Then
                            dic.Item(Val(el)) = 0&
                            If mi > el Then mi = el
                            If ma < el Then ma = el
                        End If
                    Next
                End If
            Next
            wb.Close 0
            c = c + 1
            vivod dic, mi, ma, d, c, tfilepath
        Next SelectedItem
    End With
 
    Application.ScreenUpdating = True
 
End Sub
 
 
Private Sub vivod(sl, mi, ma, d, c, tfilepath)
    Dim outsh As Object
    'для копирования в имя файла части даты
    Dim day As String
    Set outsh = Workbooks.Add(1).Sheets(1)
    Dim i&, ii&, flagS As Boolean, flagF As Boolean
 
    ReDim a(1 To (ma - mi + 3) / 2 + 1, 1 To 3)
    ii = 1: flagS = True: flagF = True
 
    For i = mi To ma + 1
        If sl.exists(i) Then
            If flagS Then
                a(ii, 1) = i: flagS = False: flagF = True
            End If
        Else
            If flagF Then
                a(ii, 2) = i - 1: a(ii, 3) = a(ii, 2) - a(ii, 1) + 1
                flagS = True: flagF = False: ii = ii + 1
            End If
        End If
    Next
 
    outsh.Cells(2, 1).Resize(ii - 1) = "Билет театральный рулонный (1 бил.=1 руб.)"
    outsh.Cells(2, 2).Resize(ii - 1) = "ТЕ"
    outsh.Cells(2, 3).Resize(ii - 1, 3) = a
    outsh.Cells(1, 1).Resize(1, 5) = Array("БСО", "Серия БСО", "Начальный номер", "Конечный номер", "Количество")
 
    If MakeSureDirectoryPathExists(tfilepath) = 0 Then _
       MsgBox "Не удалось создать путь": Exit Sub
 
    day = Mid(d, 1, 2) + Mid(d, 4, 2) + Mid(d, 9, 2) & "(" & c & ")"
 
    outsh.Parent.SaveAs Filename:=tfilepath & day & ".xls", FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
 
    outsh.Parent.Close 0
 
End Sub



