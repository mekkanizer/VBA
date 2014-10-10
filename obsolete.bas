    'Обработку квитанций и счетов по образцу обработки актов 

    'Обработку счетов фактуры (инвойсов) в основную процедуру 

    'Пакетную обработку (прием кучи) 

    'То что Тимофей упомянул вк (=)
Public Sub WordPages()
	Dim path As String
	Dim np, n1, n0 As Integer
	Dim r As Range
	Dim buh, prefix, stamps_img As String
	Dim orig_doc As Document
	Dim main_browser As Browser
     
	Set main_browser = Application.Browser
	main_browser.Target = wdBrowsePage
	Set orig_doc = ActiveDocument
	np = ActiveDocument.BuiltInDocumentProperties("Number of Pages")
	buh = "1"
	path = orig_doc.path & "\"
	prefix = Strings.Replace(Strings.Replace(orig_doc.Name, ".doc", ""), _
		".rtf", "", , , vbTextCompare)
	sYear = Year(Now)
	sMonth = Month(Now) - 1
	'   sYear = 2012
	'  sMonth = 12
	If sMonth > 0 And sMonth < 10 Then
		sMonth = "0" & sMonth
	End If
	If prefix = "mtt" Then
		stamps_img = "C:\tmp\p_mtt.png"
	Else
		stamps_img = "C:\tmp\p_tn.png"
	End If

	For i = 1 To np
	  
	  'Select and copy the text to the clipboard
	  orig_doc.Bookmarks("\page").Range.Copy
			
	  ' Open new document to paste the content of the clipboard into.
	  Documents.Add
	  If prefix = "invoiceVoip" Or prefix = "invoiceMTT" Or prefix = "invoiceTelenet" Then
	   ActiveDocument.PageSetup.Orientation = wdOrientLandscape
	  End If
	  Selection.PageSetup.LeftMargin = CentimetersToPoints(1.1)

	  ' Kolontituli
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
			  Selection.InlineShapes.AddPicture FileName:="C:\tmp\tn.JPG", LinkToFile:= _
				False, SaveWithDocument:=True
				
			  If Selection.HeaderFooter.IsHeader = True Then
				ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
			  Else
				ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
			  End If
			  Selection.InlineShapes.AddPicture FileName:=stamps_img, LinkToFile _
				:=False, SaveWithDocument:=True
		 End If
	End If
	  ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
	  Selection.Paste
	  ' Removes the break that is copied at the end of the page, if any.
	  Selection.TypeBackspace

	  
	If prefix = "kvit" Or prefix = "kvitmtt" Then
		Set r = Documents.Item(1).Range
		r.Find.Execute FindText:="Aoo. eia: ^#", Forward:=True
		n0 = r.End - 1
		Set r = Documents.Item(1).Range
		r.Find.Execute FindText:="^#^p", Forward:=True
		n1 = r.Start + 1
		  
		If (n0 < n1) Then
		  buh = Documents.Item(1).Range(n0, n1).Text & "_" & "00"
		Else
		  buh = "----_00"
		End If
	End If
	If prefix = "invoiceVoip" Or prefix = "invoiceMTT" Or prefix = "invoiceTelenet" _
	Or prefix = "actVoip" Or prefix = "actMTT" Or prefix = "actTelenet" Then
			' Buhkod searching
			  Set r = Documents.Item(1).Range
			  r.Find.Execute FindText:="B=^#", Forward:=True
			  n0 = r.End - 1
			  Set r = Documents.Item(1).Range(n0)
			  r.Find.Execute FindText:="^p", Forward:=True
			  n1 = r.Start
				
			  If (n0 < n1) Then
				buh = Documents.Item(1).Range(n0, n1).Text & "_"
			  Else
				buh = "----_"
			  End If
			' Konvert number searching
			  Set r = Documents.Item(1).Range
			  r.Find.Execute FindText:="KONV=^#", Forward:=True
			  n0 = r.End - 1
			  Set r = Documents.Item(1).Range(n0)
			  r.Find.Execute FindText:="^w", Forward:=True
			  n1 = r.Start
				
			  If (n0 < n1) Then
				buh = buh & Documents.Item(1).Range(n0, n1).Text
			  Else
				buh = buh & "----"
			  End If
			  
			  
			  
	Else
			  If Not (prefix = "kvit" Or prefix = "kvitmtt") Then
					  ' Buhkod searching
				  Set r = Documents.Item(1).Range
				  r.Find.Execute FindText:="B=^#", Forward:=True
				  n0 = r.End - 1
				  Set r = Documents.Item(1).Range
				  r.Find.Execute FindText:="^#^p", Forward:=True
				  n1 = r.Start + 1
					
				  If (n0 < n1) Then
					buh = Documents.Item(1).Range(n0, n1).Text & "_"
				  Else
					buh = buh & "----_"
				  End If
				' Konvert number searching
				  Set r = Documents.Item(1).Range
				  r.Find.Execute FindText:="KONV=^#", Forward:=True
				  n0 = r.End - 1
				  Set r = Documents.Item(1).Range
				  r.Find.Execute FindText:="^#^w", Forward:=True
				  n1 = r.Start + 1
					
				  If (n0 < n1) Then
					buh = buh & Documents.Item(1).Range(n0, n1).Text
				  Else
					buh = buh & "----"
				  End If
			  End If

	  End If

	  ChangeFileOpenDirectory path & "for_each"

	  DocNum = DocNum + 1
	'    If prefix = "invoice" Then
	  ActiveDocument.ExportAsFixedFormat outputfilename:=path & "pdf" & "\" & sYear & "_" & sMonth & "_" & buh & "_" & prefix & ".pdf", exportformat:=wdExportFormatPDF
	  ActiveDocument.Close False
	'   Else
	  
	'  ActiveDocument.SaveAs FileName:=prefix & "_" & buh & ".doc"
	'   ActiveDocument.Close
	'   End If

	  ' Move the selection to the next page  in the document
	  main_browser.Target = wdBrowsePage
	  main_browser.Next
	Next i
	ActiveDocument.Close savechanges:=wdDoNotSaveChange
	End Sub

Sub WordPages_invoice()

	Dim path As String
	Dim np, n1, n0 As Integer
	Dim r As Range
	Dim buh, prefix, stamps_img As String
	Dim orig_doc As Document
	Dim main_browser As Browser
	 
	Set main_browser = Application.Browser
	main_browser.Target = wdBrowsePage
	Set orig_doc = ActiveDocument
	np = ActiveDocument.BuiltInDocumentProperties("Number of Pages")
	buh = "1"
	path = orig_doc.path & "\"
	prefix = Strings.Replace(Strings.Replace(orig_doc.Name, ".doc", ""), ".rtf", "", , , vbTextCompare)
	'   path = ActiveDocument.path
	If prefix = "mtt" Then
		stamps_img = "C:\tmp\p_mtt.png"
	Else
		stamps_img = "C:\tmp\p_tn.png"
	End If

	For i = 1 To np
	  
	  'Select and copy the text to the clipboard
	  orig_doc.Bookmarks("\page").Range.Copy
			
	  ' Open new document to paste the content of the clipboard into.
	  Documents.Add
	  ActiveDocument.PageSetup.Orientation = wdOrientLandscape
	  Selection.PageSetup.LeftMargin = CentimetersToPoints(1.1)

	  ' Kolontituli
	If Not (prefix = "kvit" Or prefix = "kvitmtt") Then
	  If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
		ActiveWindow.Panes(2).Close
	  End If
	  If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
		ActivePane.View.Type = wdOutlineView Then
		ActiveWindow.ActivePane.View.Type = wdPrintView
	  End If
	'     ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
	'    Selection.InlineShapes.AddPicture FileName:="C:\tmp\tn.JPG", LinkToFile:= _
		False, SaveWithDocument:=True
		
	'    If Selection.HeaderFooter.IsHeader = True Then
	'      ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
	'    Else
	'      ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
	'    End If
	'    Selection.InlineShapes.AddPicture FileName:=stamps_img, LinkToFile _
		:=False, SaveWithDocument:=True
	End If
	  ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
	  Selection.Paste
	  ' Removes the break that is copied at the end of the page, if any.
	  Selection.TypeBackspace

	  
	If prefix = "kvit" Or prefix = "kvitmtt" Then
	  Set r = Documents.Item(1).Range
	  r.Find.Execute FindText:="B=^#", Forward:=True
	  n0 = r.End - 1
	  Set r = Documents.Item(1).Range
	  r.Find.Execute FindText:="^#^p", Forward:=True
	  n1 = r.Start + 1
		
	  If (n0 < n1) Then
		buh = "00" & "_" & Documents.Item(1).Range(n0, n1).Text
	  Else
		buh = "00_----"
	  End If
	Else
	' Konvert number searching
	  Set r = Documents.Item(1).Range
	  r.Find.Execute FindText:="KONV=^#", Forward:=True
	  n0 = r.End - 1
	  Set r = Documents.Item(1).Range(n0)
	  r.Find.Execute FindText:="^w", Forward:=True
	  n1 = r.Start
		
	  If (n0 < n1) Then
		buh = Documents.Item(1).Range(n0, n1).Text & "_"
	  Else
		buh = "----_"
	  End If
	  
	  
	  ' Buhkod searching
	  Set r = Documents.Item(1).Range
	  r.Find.Execute FindText:="B=^#", Forward:=True
	  n0 = r.End - 1
	  Set r = Documents.Item(1).Range(n0)
	  r.Find.Execute FindText:="^p", Forward:=True
	  n1 = r.Start
		
	  If (n0 < n1) Then
		buh = buh & Documents.Item(1).Range(n0, n1).Text
	  Else
		buh = buh & "----"
	  End If
	End If
	  ChangeFileOpenDirectory path & "for_each"

	  DocNum = DocNum + 1
	'   ActiveDocument.SaveAs FileName:=prefix & "_" & buh & ".doc"
	'  ActiveDocument.Close
	  '
	 
	  ActiveDocument.ExportAsFixedFormat outputfilename:=path & "for_each" & "\" & prefix & "_" & buh & ".pdf", exportformat:=wdExportFormatPDF
	  
	'    ActiveDocument.ExportAsFixedFormat outputfilename:=prefix & "_" & buh & ".pdf", _
	'    exportformat:=wdExportFormatPDF, openafterexport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
	'    Range:=wdExportAllDocument, FROM:=1, To:=1, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
	'    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
	'    BitmapMissingFonts:=True, UseISO19005_1:=False

	ActiveDocument.Close False


	  '

	  ' Move the selection to the next page  in the document
	  main_browser.Target = wdBrowsePage
	  main_browser.Next
	Next i
	ActiveDocument.Close savechanges:=wdDoNotSaveChange
	End Sub


