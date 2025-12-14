Sub Makro_Prio_Arkusz2()

    'Definition of variables
    Dim wb1 As Workbook, wb2 As Workbook, wb3 As Workbook, wbPRIO As Workbook
    Dim wsSrc1 As Worksheet, wsSrc2 As Worksheet, wsSrc3 As Worksheet
    Dim wsDest As Worksheet, ws3 As Worksheet
    Dim wsArkusz1 As Worksheet
    Dim lastRowSrc1 As Long, lastRowSrc2 As Long, lastRowSrc3 As Long, lastRowDest As Long
    Dim wb As Workbook
    
    'Added initial message
    MsgBox "Make sure you have the following files open: 'ydrzewo 1 ...', 'ydrzewo 2 ...', 'ydrzewo 3 ...' and '... PRIO ...'", vbInformation
    
    'Looking for required workbooks
    For Each wb In application.Workbooks
        If InStr(1, wb.Name, "ydrzewo 1", vbTextCompare) > 0 Then Set wb1 = wb
        If InStr(1, wb.Name, "ydrzewo 2", vbTextCompare) > 0 Then Set wb2 = wb
        If InStr(1, wb.Name, "ydrzewo 3", vbTextCompare) > 0 Then Set wb3 = wb
        If InStr(1, wb.Name, "PRIO", vbTextCompare) > 0 Then Set wbPRIO = wb
    Next wb
    
    'Control if all required workbooks are found
    If wb1 Is Nothing Or wb2 Is Nothing Or wb3 Is Nothing Or wbPRIO Is Nothing Then
        MsgBox "Brakuje otwartych plików! Otwórz wszystkie wymagane pliki.", vbCritical
        Exit Sub
    End If
    
    'Setting source worksheets
    Dim idx As Long
    
    Set wsSrc1 = wb1.Sheets(1)
    Set wsSrc2 = wb2.Sheets(1)
    Set wsSrc3 = wb3.Sheets(1)
    
    Set wsArkusz1 = wbPRIO.Sheets("Arkusz1")
    idx = wsArkusz1.Index
    
    If idx < wbPRIO.Sheets.Count Then
        'if Arkusz1 is not the last sheet
        Set wsDest = wbPRIO.Sheets(idx + 1)
    Else
        'If Arkusz1 is the last sheet, create a new one
        Set wsDest = wbPRIO.Sheets.Add(After:=wsArkusz1)
        wsDest.Name = "Arkusz" & wbPRIO.Sheets.Count ' unikalna nazwa
    End If
    
    'Clearing destination worksheet
    wsDest.Cells.Clear
    
    'Copying data from ydrzewo 1
    lastRowSrc1 = wsSrc1.Cells(wsSrc1.Rows.Count, "B").End(xlUp).Row
    wsSrc1.Range("B6:K" & lastRowSrc1).Copy wsDest.Range("A1")
    
    'Copying data from ydrzewo 2
    lastRowSrc2 = wsSrc2.Cells(wsSrc2.Rows.Count, "B").End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
    wsSrc2.Range("B6:K" & lastRowSrc2).Copy wsDest.Range("A" & lastRowDest)
    
    'Copying data from ydrzewo 3
    lastRowSrc3 = wsSrc3.Cells(wsSrc3.Rows.Count, "B").End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
    wsSrc3.Range("B6:K" & lastRowSrc3).Copy wsDest.Range("A" & lastRowDest)
    
    'Adding VLOOKUP formula in column K in destination sheet
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    wsDest.Range("K1:K" & lastRowDest).FormulaLocal = _
        "=WYSZUKAJ.PIONOWO(A1;Arkusz1!B:C;2;0)"
    
    'Copying columns J and K to Arkusz3
    Set ws3 = wbPRIO.Sheets.Add(After:=wsDest)
    wsDest.Range("J1:K" & lastRowDest).Copy
    ws3.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    'Adding headers to ws3
    ws3.Range("A1").Value = "a"
    ws3.Range("B1").Value = "b"

    ws3.Range("A1:B1").AutoFilter
    
    'Sorting ws3 by column B
    ws3.Range("A:B").Sort Key1:=ws3.Range("B2"), Order1:=xlAscending, Header:=xlYes
    
    'Deleting duplicates in ws3 based on column A
    ws3.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
    
    'Final cleanup and message
    application.CutCopyMode = False
    MsgBox "Done! All operations completed successfully in the Prio file", vbInformation

End Sub
