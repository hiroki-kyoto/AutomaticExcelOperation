Sub ExportDailyReport()
    Dim fnValid As String
    Dim fnInvalid As String
    Dim fnDailyReport As String
    
    Dim numValidCol As Integer
    Dim numValidRow As Integer
    Dim numInvalidCol As Integer
    Dim numInvalidRow As Integer
    
    Dim wbDailyReport As String
    Dim wbOutput As String
    Dim wbInvalid As String
    Dim wbValid As String
    
    Dim shInvalid As String
    Dim shValid As String
    Dim shOutput As String
    
    Dim perspValid As String
    Dim perspInvalid As String
    
    Dim cols_to_move(1 To 4) As String
    
    fnOutput = "output.xlsx"
    fnValid = SelectWorkbook("有效表工作簿")
    If fnValid = "" Then
        Exit Sub
    End If
    fnInvalid = SelectWorkbook("无效表工作簿")
    If fnInvalid = "" Then
        Exit Sub
    End If
    fnDailyReport = SelectWorkbook("渠道日报表工作簿")
    If fnDailyReport = "" Then
        Exit Sub
    End If
    
    ''' Save daily report workbook as another copy to be modified
    Workbooks.Open fnDailyReport
    wbDailyReport = ActiveWorkbook.Name
    Workbooks(wbDailyReport).SaveAs ThisWorkbook.Path & "/" & fnOutput
    wbOutput = ActiveWorkbook.Name
    shOutput = ActiveWorkbook.Sheets(1).Name
    
    ''' step#1 Create Perspectives for Valid Workbook and the Invalid
    Workbooks.Open fnValid
    wbValid = ActiveWorkbook.Name
    shValid = ActiveWorkbook.Sheets(1).Name
    numValidCol = CountColumns(wbValid, shValid)
    numValidRow = CountRows(wbValid, shValid)
    
    perspValid = MakePerspective(wbValid, _
        shValid, 1, 2, numValidCol, numValidRow, _
        "最终原因代码", "渠道", "申请书编号")
    ' Copy perspective sheets to new sheets without equation and write-protect
    perspValid = CopySheet(wbValid, perspValid)
    'Workbooks(wbValid).SaveAs ThisWorkbook.Path & "/val_mod.xlsx"
    
    Workbooks.Open fnInvalid
    wbInvalid = ActiveWorkbook.Name
    shInvalid = ActiveWorkbook.Sheets(1).Name
    numInvalidCol = CountColumns(wbInvalid, shInvalid)
    numInvalidRow = CountRows(wbInvalid, shInvalid)
    
    perspInvalid = MakePerspective(wbInvalid, _
        shInvalid, 1, 2, numInvalidCol, numInvalidRow, _
        "最终原因代码", "渠道", "申请书编号")
    ' Copy perspective sheets to new sheets without equation and write-protect
    perspInvalid = CopySheet(wbInvalid, perspInvalid)
    'Workbooks(wbInvalid).SaveAs ThisWorkbook.Path & "/inval_mod.xlsx"
    
    ' step#2-5 copy extra columns from sheet2 to sheet1
    cols_to_move(1) = "D210"
    cols_to_move(2) = "D617"
    cols_to_move(3) = "D710"
    cols_to_move(4) = "D307"
    Call CopyExtraColumnsFromSheet2ToSheet1(wbValid, perspValid, _
        wbInvalid, perspInvalid, cols_to_move)
    
    'step#6 copy extra rows from sheet1 to output sheet
    Call CopyExtraRowsFromSheet2ToSheet1(wbOutput, shOutput, wbValid, perspValid)
    
    
    'Workbooks(wbOutput).Close False
    'Workbooks(wbValid).Close False
    'Workbooks(wbInvalid).Close False
    
    MsgBox "任务完成^.^"
    
End Sub

Function SelectWorkbook(str As String)
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "选择" & str, "*.xlsx"
        If .Show = -1 Then
            SelectWorkbook = .SelectedItems(1)
        Else
            SelectWorkbook = ""
        End If
    End With
End Function

''' wb : specified workbook name
''' id : sheet name to count columns
Function CountColumns(wb As String, id As String)
    CountColumns = Workbooks(wb).Sheets(id).UsedRange.Columns.Count
End Function

''' wb : specified workbook name
''' id : sheet name to count rows
Function CountRows(wb As String, id As String):
    CountRows = Workbooks(wb).Sheets(id).UsedRange.Rows.Count
End Function

''' wb : workbook name
''' id : sheet id to make a perspective
''' bColIdx : begin Column Index,
''' bRowIdx : begin Row Index
''' eColIdx : end Column Index
''' eRowIdx : end Row Index
''' pColField : perspective column fields
''' pRowField : perspective row fields
''' pDataField : perspective data fields
''' return : a name of sheet of perspective
Function MakePerspective(wb As String, id As String, _
    bColIdx As Integer, _
    bRowIdx As Integer, _
    eColIdx As Integer, _
    eRowIdx As Integer, _
    pColField As String, _
    pRowField As String, _
    pDataField As String _
) As String
    Dim tid As String
    Dim pTableName As String
    
    Workbooks(wb).Sheets.Add
    tid = Workbooks(wb).ActiveSheet.Name
    MakePerspective = tid
    pTableName = tid & "Perspective"
    Workbooks(wb).Activate
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:= _
        id & "!R" & bRowIdx & "C" & bColIdx & ":R" & eRowIdx & "C" & eColIdx, _
        Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:=tid & "!R3C1", _
        TableName:=pTableName, _
        DefaultVersion:=xlPivotTableVersion12
    
    ActiveWorkbook.Sheets(tid).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pTableName).PivotFields(pColField)
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pTableName).PivotFields(pRowField)
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pTableName).AddDataField ActiveSheet.PivotTables(pTableName _
        ).PivotFields(pDataField), "计数项:" & pDataField, xlCount
End Function

''' Copy a sheet and paste it current workbook
''' wb : workbook name
''' sh : sheet name
''' return : the new sheet pasted
Function CopySheet(wb As String, sh As String) As String
    Workbooks(wb).Activate
    ActiveWorkbook.Sheets(sh).Select
    ActiveWorkbook.Sheets(sh).Cells.Select
    Selection.Copy
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    CopySheet = ActiveWorkbook.ActiveSheet.Name
End Function

''' check if str exists in arr
''' Yes, then return its first index, no, return -1
''' arr must be an range array
''' arr : range array, index from 1, two dimensional.
Function IndexOf(str, arr) As Integer
    Dim i As Integer
    Dim n As Integer
    n = UBound(arr, 1) - LBound(arr, 1) + 1
    
    IndexOf = -1
    
    For i = 1 To n
        If str = arr(i, 1) Then
            IndexOf = i
            Exit For
        End If
    Next
End Function

''' Find search target in sheet with given column id
Function FindInRow(wb, sh, row_id, search_target)
    Dim flag As Boolean
    Dim while_flag As Boolean
    Dim cur_col_id As Integer
    Dim cur_row_id As Integer
    flag = False
    while_flag = True
    Workbooks(wb).Activate
    ActiveWorkbook.Sheets(sh).Select
    Cells.Find(What:=search_target, After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
    If ActiveCell.Row = row_id Then
        flag = True
    Else
        cur_col_id = ActiveCell.Column
        cur_row_id = ActiveCell.Row
        Do While while_flag
            Cells.FindNext(After:=ActiveCell).Activate
            If ActiveCell.Row = row_id Then
                flag = True
                while_flag = False
            Else
                ' check if enter search loop
                If ActiveCell.Column = cur_col_id And ActiveCell.Row = cur_row_id Then
                    while_flag = False
                End If
            End If
        Loop
    End If
    If flag Then
        FindInRow = ActiveCell.Column
    Else
        FindInRow = -1
    End If
End Function

''' wb : workbook
''' sh : sheet
''' col_id_before : index of column before insertion
Function InsertColumn(wb, sh, col_id_before)
    Workbooks(wb).Activate
    ActiveWorkbook.Sheets(sh).Select
    Columns(col_id_before).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    InsertColumn = col_id_before
End Function

Function DeleteColumn(wb, sh, col_id)
    Workbooks(wb).Activate
    ActiveWorkbook.Sheets(sh).Select
    Columns(col_id).Select
    Selection.Delete Shift:=xlToLeft
    DeleteColumn = True
End Function


''' Find the columns in sheet2 but not in sheet1 and fill them into sheet1.
Sub CopyExtraColumnsFromSheet2ToSheet1(wb1 As String, sh1 As String, wb2 As String, sh2 As String, cols)
    
    nRows_1 = CountRows(wb1, sh1) - 3
    nRows_2 = CountRows(wb2, sh2) - 3
    rowNames_1 = Workbooks(wb1).Sheets(sh1).Range("A5:A" & (nRows_1 + 5 - 1))
    rowNames_2 = Workbooks(wb2).Sheets(sh2).Range("A5:A" & (nRows_2 + 5 - 1))
    Dim rowIds_add() As Integer
    Dim row_map() As Integer 'map from rows of sheet2 to that of sheet1
    Dim i As Integer
    Dim counter As Integer
    Dim nCols As Integer
    ' Notice that range array is 2-dimension and index begins from 1
    ' However, array(1 to n) is 1-dimension and index begins from 0
    counter = 0
    ReDim rowIds_add(1 To nRows_2) As Integer 'index from 1
    ReDim row_map(1 To nRows_2) As Integer 'index from 1
    
    For i = 1 To nRows_2
        map_id = IndexOf(rowNames_2(i, 1), rowNames_1) 'index from 1
        If map_id = -1 Then
            counter = counter + 1
            rowIds_add(counter) = i 'index from 1
            row_map(i) = nRows_1 + counter ' index from 1
        Else
            row_map(i) = map_id
        End If
    Next
    
    ' copy the extra names from sheet2 to sheet1
    ' move the last summary line to new position
    Workbooks(wb1).Activate
    ActiveWorkbook.Sheets(sh1).Select
    last_row_id = nRows_1 + 5
    ActiveWorkbook.ActiveSheet.Rows("" & last_row_id & ":" & last_row_id).Select
    Selection.Cut
    last_row_id = nRows_1 + 4 + counter + 1
    ActiveWorkbook.ActiveSheet.Rows("" & last_row_id & ":" & last_row_id).Select
    ActiveSheet.Paste
    
    For i = 1 To counter
        Workbooks(wb1).Sheets(sh1).Range("A" & (nRows_1 + 4 + i)).Value _
            = rowNames_2(rowIds_add(i), 1)
    Next
    
    nCols = UBound(cols, 1) - LBound(cols, 1) + 1
    
    For i = 1 To nCols
        ' add this column to sheet1
        new_col_id = InsertColumn(wb1, sh1, CountColumns(wb1, sh1))
        ' find out column indexes of given column names
        ' ROW:4 COLUMN: B-...
        ' Cells(Row_id, Col_id)
        head_row_id = 4
        Workbooks(wb1).Sheets(sh1).Cells(head_row_id, new_col_id) = cols(i)
        col_id = FindInRow(wb2, sh2, head_row_id, cols(i))
        If col_id = -1 Then
            MsgBox "错误：在工作簿" & wb2 & "的表" & sh2 & "中没有找到字段：" & cols(i)
            Error 1
        End If
        
        For j = 1 To nRows_2
            Workbooks(wb1).Sheets(sh1).Cells(head_row_id + row_map(j), new_col_id) = _
                Workbooks(wb2).Sheets(sh2).Cells(head_row_id + j, col_id)
        Next
    Next
    
    ''' add equations for last row and last column in sheet1
    last_col_id = CountColumns(wb1, sh1)
    ' last column first
    For i = 5 To (last_row_id - 1)
        begin_address = Workbooks(wb1).Sheets(sh1).Cells(i, 2).Address
        end_address = Workbooks(wb1).Sheets(sh1).Cells(i, last_col_id - 1).Address
        Workbooks(wb1).Sheets(sh1).Cells(i, last_col_id) = _
            "=SUM(" & begin_address & ":" & end_address & ")"
    Next
    ' last row following
    For i = 2 To last_col_id
        begin_address = Workbooks(wb1).Sheets(sh1).Cells(5, i).Address
        end_address = Workbooks(wb1).Sheets(sh1).Cells(last_row_id - 1, i).Address
        Workbooks(wb1).Sheets(sh1).Cells(last_row_id, i) = _
            "=SUM(" & begin_address & ":" & end_address & ")"
    Next
    
    ''' Delete columns from sheet2
    For i = 1 To nCols
        head_row_id = 4
        col_id = FindInRow(wb2, sh2, head_row_id, cols(i))
        If col_id = -1 Then
            MsgBox "错误：在工作簿" & wb2 & "的表" & sh2 & "中没有找到字段：" & cols(i)
            Error 1
        End If
        Call DeleteColumn(wb2, sh2, col_id)
    Next
    
    ''' Recalculate the summary entry for sheet2
    ''' add equations for last row and last column in sheet2
    last_col_id = CountColumns(wb2, sh2)
    last_row_id = CountRows(wb2, sh2) + 2
    ' last column first
    For i = 5 To (last_row_id - 1)
        begin_address = Workbooks(wb2).Sheets(sh2).Cells(i, 2).Address
        end_address = Workbooks(wb2).Sheets(sh2).Cells(i, last_col_id - 1).Address
        Workbooks(wb2).Sheets(sh2).Cells(i, last_col_id) = _
            "=SUM(" & begin_address & ":" & end_address & ")"
    Next
    ' last row following
    For i = 2 To last_col_id
        begin_address = Workbooks(wb2).Sheets(sh2).Cells(5, i).Address
        end_address = Workbooks(wb2).Sheets(sh2).Cells(last_row_id - 1, i).Address
        Workbooks(wb2).Sheets(sh2).Cells(last_row_id, i) = _
            "=SUM(" & begin_address & ":" & end_address & ")"
    Next
    ''' over
End Sub

''' Get the last unblank index of first row
Function GetEndRowIndex(wb As String, sh As String, row_beg As Integer, col_id As Integer)
    Dim i As Integer
    i = row_beg
    Do While Workbooks(wb).Sheets(sh).Cells(i, col_id) <> ""
        i = i + 1
    Loop
    GetEndRowIndex = i - 1
End Function

Function GetEndColIndex(wb As String, sh As String, col_beg As Integer, row_id As Integer)
    Dim i As Integer
    i = col_beg
    Do While Workbooks(wb).Sheets(sh).Cells(row_id, i) <> ""
        i = i + 1
    Loop
    GetEndColIndex = i - 1
End Function

''' Insert a row at given row index
''' formats of columns at that row will inherit the row above it.
Sub InsertRowAt(wb As String, sh As String, row_id As Integer)
    Workbooks(wb).Activate
    ActiveWorkbook.Sheets(sh).Select
    ActiveSheet.Rows(row_id).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows(row_id - 1).Select
    Selection.Copy
    Rows(row_id).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Function Max(a As Integer, b As Integer)
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

''' copy new sources from valid sheet to output sheet
''' wb1 : the output workbook
''' sh1 : the output sheet
''' wb2 : the valid workbook
''' sh2 : the valid sheet
Sub CopyExtraRowsFromSheet2ToSheet1(wb1 As String, sh1 As String, wb2 As String, sh2 As String)
    
    Dim rowIds_add() As Integer
    Dim row_map() As Integer 'map from rows of sheet2 to that of sheet1
    Dim i As Integer
    Dim counter As Integer
    Dim nCols As Integer
    Dim last_col_id As Integer
    Dim sum_value As Long
    Dim a As Double
    Dim b As Double
    
    Workbooks(wb2).Activate
    Sheets(sh2).Select
    nRows_1 = GetEndRowIndex(wb1, sh1, 1, 2) - 2 'B3-B$end
    nRows_2 = CountRows(wb2, sh2) - 3 'A5-A$end-1
    rowNames_1 = Workbooks(wb1).Sheets(sh1).Range("B3:B" & (2 + nRows_1))
    rowNames_2 = Workbooks(wb2).Sheets(sh2).Range("A5:A" & (nRows_2 + 5 - 1))
    
    counter = 0
    ReDim rowIds_add(1 To nRows_2) As Integer 'index from 1
    ReDim row_map(1 To nRows_2) As Integer 'index from 1
    
    For i = 1 To nRows_2
        map_id = IndexOf(rowNames_2(i, 1), rowNames_1) 'index from 1
        If map_id = -1 Then
            counter = counter + 1
            rowIds_add(counter) = i 'index from 1
            row_map(i) = nRows_1 + counter ' index from 1
        Else
            row_map(i) = map_id
        End If
    Next
    
    For i = 1 To counter
        'insert a row
        Call InsertRowAt(wb1, sh1, nRows_1 + 2 + i)
        Workbooks(wb1).Sheets(sh1).Range("B" & (nRows_1 + 2 + i)).Value _
            = rowNames_2(rowIds_add(i), 1)
    Next
    
    'Fill in the column with last column from sheet2
    'A column is inserted to sheet1
    new_col_id = InsertColumn(wb1, sh1, 4) 'insert a column after to [ÐÂÔöÓÐÐ§ÉêÇëÁ¿]
    'Fill in cells with default value [0]
    Workbooks(wb1).Sheets(sh1).Cells(1, new_col_id).Value = "ÐÂÔöÓÐÐ§ÉêÇëÁ¿"
    For i = 3 To (nRows_1 + 2)
        Workbooks(wb1).Sheets(sh1).Cells(i, new_col_id).Value = 0
    Next
    
    last_col_id = GetEndColIndex(wb2, sh2, 1, 4)
    first_row_id = 5
    last_row_id = GetEndRowIndex(wb2, sh2, 5, last_col_id) - 1 'the last row is summary row, not included
    sum_value = 0
    For i = first_row_id To last_row_id
        Workbooks(wb1).Sheets(sh1).Cells(2 + row_map(i - 4), new_col_id).Value = _
            Workbooks(wb2).Sheets(sh2).Cells(i, last_col_id).Value
        sum_value = sum_value + Workbooks(wb2).Sheets(sh2).Cells(i, last_col_id).Value
    Next
    
    'Update the summary at the second row
    Workbooks(wb1).Sheets(sh1).Cells(2, new_col_id).Value = sum_value
    'Update increasing rate column
    nRows_1 = nRows_1 + counter
    For i = 2 To (nRows_1 + 2)
        a = Workbooks(wb1).Sheets(sh1).Cells(i, new_col_id - 1).Value
        b = Workbooks(wb1).Sheets(sh1).Cells(i, new_col_id).Value
        If a > 0 Then
            Workbooks(wb1).Sheets(sh1).Cells(i, new_col_id + 1).Value = (b - a) / a
        Else
            Workbooks(wb1).Sheets(sh1).Cells(i, new_col_id + 1).Value = "#DIV/0!"
        End If
    Next
    
End Sub


Sub Test()
    'MsgBox FindInRow(ThisWorkbook.Name, ActiveSheet.Name, 3, "33")
    'MsgBox Cells(1, 2)
    'MsgBox Range("B3").Row
    'Columns(3).Select
    'MsgBox Cells(1, 1).Address
    'MsgBox ActiveSheet.UsedRange.Columns.Count
    'MsgBox GetEndRowIndex(ActiveWorkbook.Name, ActiveSheet.Name, 1, 1)
    'MsgBox GetEndColIndex(ActiveWorkbook.Name, ActiveSheet.Name, 1, 1)
    MsgBox Max(0, -1)
End Sub
