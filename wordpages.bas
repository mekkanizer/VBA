Option Explicit

Declare Function MakeSureDirectoryPathExists Lib "Imagehlp.dll" (ByVal strPath As String) As Long

Public Sub wordpages()
    Dim path As String
    Dim np, n1, n0, i As Integer
    Dim r As Range
    Dim buh, prefix, stamps_img, syear, smonth As String
    Dim orig_doc As Document
    Dim main_browser As Browser
        
        Dim delims, iCtr As Integer
        Dim buhkod As Boolean
        Dim pr As Paragraph
   
    'текущий файл
    Dim selecteditem
        'текущая рабочая книга
    Dim doc As Document
        
        'вызываем диалог выбора файлов
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Выберите файлы"    'надпись в окне диалога
        'путь по умолчанию к папке где расположен исходный файл, можно изменить
        .InitialFileName = ThisDocument.path & Application.PathSeparator & "*.rtf"
        .AllowMultiSelect = True    'выбор нескольких файлов разрешён
        If .Show = False Then Exit Sub
 
        Application.ScreenUpdating = False
        For Each selecteditem In .SelectedItems    'перебор файлов в папке
                'открываем книгу
        Documents.Open (selecteditem)
        Set orig_doc = ActiveDocument
        
        delims = Array(" =", "= ")

        For iCtr = 0 To UBound(delims)
            With ActiveDocument.Range.Find
                .Text = delims(iCtr)
                .Replacement.Text = delims(iCtr)
                .Forward = True
                .Execute Replace:=wdReplaceAll
            End With
        Next
        
        path = ActiveDocument.path & "\"
        prefix = Strings.Replace(Strings.Replace(orig_doc.Name, ".doc", ""), ".rtf", "", , , vbTextCompare)
        syear = Year(Now)
        smonth = Month(Now) - 1
        If smonth > 0 And smonth < 10 Then
                smonth = "0" & smonth
        End If

                Set main_browser = Application.Browser
                main_browser.Target = wdBrowsePage
                Application.ActiveDocument.Repaginate
                np = ActiveDocument.BuiltInDocumentProperties("Number of Pages")
                buh = "1"
                If prefix = "mtt" Then
                        stamps_img = "c:\tmp\p_mtt.png"
                Else
                        stamps_img = "c:\tmp\p_tn.png"
                End If
                                
                For i = 1 To np
                        'select and copy the text to the clipboard
                        orig_doc.Bookmarks("\page").Range.Copy
                                
                        ' open new document to paste the content of the clipboard into.
                        Documents.Add
                        If prefix = "invoiceVoip" Or prefix = "invoiceMTT" Or prefix = "invoiceTelenet" Then
                                ActiveDocument.PageSetup.Orientation = wdOrientLandscape
                        End If
                        Selection.PageSetup.LeftMargin = CentimetersToPoints(1.1)

                ' kolontituli
                If Not (prefix = "kvit" Or prefix = "kvitmtt") Then
                        If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
                                ActiveWindow.Panes(2).Close
                        End If
                        If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
                        ActivePane.View.Type = wdOutlineView Then
                                ActiveWindow.ActivePane.View.Type = wdPrintView
                        End If
                        If prefix = "telenet" Or prefix = "mtt" Or prefix = "voip" Then
                                ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
                                Selection.InlineShapes.AddPicture FileName:="c:\tmp\tn.jpg", linktofile:= _
                                False, savewithdocument:=True

                        If Selection.HeaderFooter.IsHeader = True Then
                                ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                        Else
                                ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
                        End If
                        Selection.InlineShapes.AddPicture FileName:=stamps_img, linktofile _
                        :=False, savewithdocument:=True
                        End If
                End If
                ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
                Selection.Paste
                ' removes the break that is copied at the end of the page, if any.
                Selection.TypeBackspace
                                
                If prefix = "kvit" Or prefix = "kvitmtt" Then
                    buhkod = False
                    For Each pr In Documents.Item(1).Paragraphs
                        If (pr.Range.Information(wdWithInTable) = False) _
                        And (Len(pr.Range.Text) > 2) Then
                            If buhkod = True Then
                                buhkod = False
                                buh = Right(Left(pr.Range.Text, Len(pr.Range.Text) - 1), _
                                Len(Left(pr.Range.Text, Len(pr.Range.Text) - 1)) - _
                                (InStrRev(Left(pr.Range.Text, Len(pr.Range.Text) - 1), " ") - 1)) & "_00"
                            Else
                                buhkod = True
                            End If
                        End If
                    Next pr
                End If
                                
                If prefix = "invoiceVoip" Or prefix = "invoiceMTT" Or prefix = "invoiceTelenet" _
                Or prefix = "actVoip" Or prefix = "actMTT" Or prefix = "actTelenet" Then
                                ' buhkod searching
                                  Set r = Documents.Item(1).Range
                                  r.Find.Execute findtext:="b=^#", Forward:=True
                                  n0 = r.End - 1
                                  Set r = Documents.Item(1).Range(n0)
                                  r.Find.Execute findtext:="^p", Forward:=True
                                  n1 = r.Start
                                        
                                  If (n0 < n1) Then
                                        buh = Documents.Item(1).Range(n0, n1).Text & "_"
                                  Else
                                        buh = "----_"
                                  End If
                                ' konvert number searching
                                  Set r = Documents.Item(1).Range
                                  r.Find.Execute findtext:="konv=^#", Forward:=True
                                  n0 = r.End - 1
                                  Set r = Documents.Item(1).Range(n0)
                                  r.Find.Execute findtext:="^w", Forward:=True
                                  n1 = r.Start
                                        
                                  If (n0 < n1) Then
                                        buh = buh & Documents.Item(1).Range(n0, n1).Text
                                  Else
                                        buh = buh & "----"
                                  End If
                Else
                                  If Not (prefix = "kvit" Or prefix = "kvitmtt") Then
                                                  ' buhkod searching
                                          Set r = Documents.Item(1).Range
                                          r.Find.Execute findtext:="B=^#", Forward:=True
                                          n0 = r.End - 1
                                          Set r = Documents.Item(1).Range
                                          r.Find.Execute findtext:="^#^p", Forward:=True
                                          n1 = r.Start + 1
                                                
                                          If (n0 < n1) Then
                                                buh = Documents.Item(1).Range(n0, n1).Text & "_"
                                          Else
                                                buh = buh & "----_"
                                          End If
                                        ' konvert number searching
                                          Set r = Documents.Item(1).Range
                                          r.Find.Execute findtext:="KONV=^#", Forward:=True
                                          n0 = r.End - 1
                                          Set r = Documents.Item(1).Range
                                          r.Find.Execute findtext:="^#^w", Forward:=True
                                          n1 = r.Start + 1
                                                
                                          If (n0 < n1) Then
                                                buh = buh & Documents.Item(1).Range(n0, n1).Text
                                          Else
                                                buh = buh & "----"
                                          End If
                                  End If

                  End If
                                  
                                  If MakeSureDirectoryPathExists(path & "for_each\") = 0 Then _
                                        MsgBox "Не удалось создать путь": Exit Sub
                  
                                  ChangeFileOpenDirectory path & "for_each"
                                  
                  If MakeSureDirectoryPathExists(path & "pdf\") = 0 Then _
                                        MsgBox "Не удалось создать путь": Exit Sub
                  ActiveDocument.ExportAsFixedFormat outputfilename:=path & "pdf" & "\" & syear & "_" & smonth & "_" & buh & "_" & prefix & ".pdf", exportformat:=wdExportFormatPDF
                  ActiveDocument.Close False

                  main_browser.Target = wdBrowsePage
                  main_browser.Next
           Next i
   ActiveDocument.Close savechanges:=wdDoNotSaveChanges
   Next selecteditem
   End With
End Sub
