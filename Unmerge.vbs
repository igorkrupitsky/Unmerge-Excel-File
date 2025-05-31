Set fso = CreateObject("Scripting.FileSystemObject")

Dim sFilePath1
If WScript.Arguments.Count = 1 then
	sFilePath1 = WScript.Arguments(0)
Else
    MsgBox("Please drag an excel file.")	    
    Wscript.Quit
End If

If fso.FileExists(sFilePath1) = False  Then
	MsgBox "File 1 is missing: " & sFilePath1
    Wscript.Quit
End If

Dim oExcel: Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True
oExcel.DisplayAlerts = false
Set oWorkBook1 = oExcel.Workbooks.Open(sFilePath1)

For Each oSheet in oWorkBook1.Worksheets
    oSheet.Activate

    iColCount = GetLastCol(oSheet)
    iRowsCount = GetLastRowWithData(oSheet)

    For iRow = 1 to iRowsCount
        For iCol = 1 to iColCount
            Set oRange = oSheet.Cells(iRow, iCol)
            If oRange.MergeCells Then
                If iRow > 1 And oRange.MergeArea.Count > 1 And oRange.MergeArea.Columns.Count = 1 And oRange.MergeArea.Rows.Count > 1 Then
                    sValue = oRange.value
                    iRowCount = oRange.MergeArea.Rows.Count
                    oRange.MergeArea.UnMerge

                    For i = 2 to iRowCount
                        Set oCell = oSheet.Cells(iRow + (i-1), iCol)

                        If oCell.Value = "" Then
                            oCell.Value = sValue
                        End If                            
                    Next                    

                End If
            End If
        Next
    Next
Next

MsgBox "Done"

Function GetLastRowWithData(oSheet)
    Dim iMaxRow: iMaxRow = oSheet.UsedRange.Rows.Count
    If iMaxRow > 500 Then
        iMaxRow = oSheet.Cells.Find("*", oSheet.Cells(1, 1),  -4163, , 1, 2).Row
    End If

    Dim iRow, iCol
    For iRow = iMaxRow to 1 Step -1
         For iCol = 1 to oSheet.UsedRange.Columns.Count
            If Trim(oSheet.Cells(iRow, iCol).Value) <> "" Then
                GetLastRowWithData = iRow
                Exit Function
            End If
         Next
    Next
    GetLastRowWithData = 1
End Function

Function GetLastCol(st)
    on error resume next
    GetLastCol = st.Cells.Find("*", st.Cells(1, 1), , 2, 2, 2, False).Column
    If Err.number <> 0 Then
        GetLastCol = 0
    End If
End Function

Function SheetExists(oWorkBook, sName)
    on error resume next
    Dim oSheet: Set oSheet = oWorkBook.Worksheets(sName) 
    If Err.number = 0 Then
        SheetExists = True
    Else
        SheetExists = False
        Err.Clear
    End If
End Function
