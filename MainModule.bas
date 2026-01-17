Option Explicit

Sub CleanData(Control As IRibbonControl)
    ' Enhanced Data Cleaning: Removes blanks, duplicates, totals, harmonizes formats, handles protected cells
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.ProtectContents Then
        MsgBox "Sheet is protected. Unprotect it first.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Remove blank rows
    Dim i As Long
    For i = lastRow To 1 Step -1
        If Application.CountA(ws.Rows(i)) = 0 Then ws.Rows(i).Delete
    Next i
    
    ' Remove total rows (simple check: if cell in last column says "Total")
    For i = lastRow To 2 Step -1 ' Skip header
        If UCase(ws.Cells(i, lastCol).Value) = "TOTAL" Then ws.Rows(i).Delete
    Next i
    
    ' Remove duplicates (basic on first column)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' Harmonize formats (dates, numbers)
    Dim cell As Range
    For Each cell In ws.UsedRange
        If IsDate(cell.Text) Then cell.Value = CDate(cell.Text)
        If IsNumeric(cell.Text) Then cell.Value = CDbl(cell.Text)
    Next cell
    
    MsgBox "Data cleaned: Blanks, totals, and duplicates removed. Formats harmonized.", vbInformation
End Sub

Sub SummaryStats(Control As IRibbonControl)
    ' Basic Summary Statistics
    Dim ws As Worksheet, summarySheet As Worksheet
    Set ws = ActiveSheet
    Set summarySheet = GetOrCreateSheet("Basic Summary")
    
    summarySheet.Cells.Clear
    summarySheet.Cells(1, 1).Value = "Statistic"
    summarySheet.Cells(1, 2).Value = "Value"
    
    summarySheet.Cells(2, 1).Value = "Row Count"
    summarySheet.Cells(2, 2).Value = ws.UsedRange.Rows.Count - 1 ' Exclude header
    
    summarySheet.Cells(3, 1).Value = "Sum (First Numeric Col)"
    summarySheet.Cells(3, 2).Value = Application.Sum(ws.Columns(FindFirstNumericColumn(ws)))
    
    summarySheet.Cells(4, 1).Value = "Average (First Numeric Col)"
    summarySheet.Cells(4, 2).Value = Application.Average(ws.Columns(FindFirstNumericColumn(ws)))
    
    MsgBox "Basic summary generated in 'Basic Summary' sheet.", vbInformation
End Sub

Sub AdvancedSummary(Control As IRibbonControl)
    ' Advanced Summary: Min, Max, Stratification (simple bins)
    Dim ws As Worksheet, advSummarySheet As Worksheet
    Set ws = ActiveSheet
    Set advSummarySheet = GetOrCreateSheet("Advanced Summary")
    
    advSummarySheet.Cells.Clear
    advSummarySheet.Cells(1, 1).Value = "Advanced Statistic"
    advSummarySheet.Cells(1, 2).Value = "Value"
    
    Dim numCol As Integer
    numCol = FindFirstNumericColumn(ws)
    
    advSummarySheet.Cells(2, 1).Value = "Min Value"
    advSummarySheet.Cells(2, 2).Value = Application.Min(ws.Columns(numCol))
    
    advSummarySheet.Cells(3, 1).Value = "Max Value"
    advSummarySheet.Cells(3, 2).Value = Application.Max(ws.Columns(numCol))
    
    ' Simple stratification (count in bins, e.g., <100, 100-1000, >1000)
    advSummarySheet.Cells(4, 1).Value = "Count <100"
    advSummarySheet.Cells(4, 2).Value = Application.CountIf(ws.Columns(numCol), "<100")
    
    advSummarySheet.Cells(5, 1).Value = "Count 100-1000"
    advSummarySheet.Cells(5, 2).Value = Application.CountIfs(ws.Columns(numCol), ">=100", ws.Columns(numCol), "<=1000")
    
    advSummarySheet.Cells(6, 1).Value = "Count >1000"
    advSummarySheet.Cells(6, 2).Value = Application.CountIf(ws.Columns(numCol), ">1000")
    
    MsgBox "Advanced summary generated in 'Advanced Summary' sheet.", vbInformation
End Sub

Sub AdvancedDuplicates(Control As IRibbonControl)
    ' Advanced Duplicate Detection: Same-Same-Different on up to 3 fields
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Assume columns A, B, C for test (e.g., Invoice#, Amount, Date)
    Dim dupSheet As Worksheet
    Set dupSheet = GetOrCreateSheet("Duplicates")
    dupSheet.Cells.Clear
    
    Dim i As Long, j As Long, dupCount As Long
    dupCount = 0
    For i = 2 To lastRow ' Skip header
        For j = i + 1 To lastRow
            If ws.Cells(i, 1) = ws.Cells(j, 1) And ws.Cells(i, 2) = ws.Cells(j, 2) And ws.Cells(i, 3) <> ws.Cells(j, 3) Then
                dupCount = dupCount + 1
                dupSheet.Cells(dupCount + 1, 1).Value = "Duplicate Pair: Row " & i & " and " & j
            End If
        Next j
    Next i
    
    If dupCount > 0 Then
        MsgBox dupCount & " same-same-different duplicates found in 'Duplicates' sheet.", vbInformation
    Else
        MsgBox "No advanced duplicates found.", vbInformation
    End If
End Sub

Sub JoinAppendData(Control As IRibbonControl)
    ' Join/Append Datasets from Different Sheets
    ' Assume Sheet1 and Sheet2; append Sheet2 to Sheet1
    Dim srcWs As Worksheet, tgtWs As Worksheet
    If ActiveWorkbook.Sheets.Count < 2 Then
        MsgBox "Need at least two sheets for append.", vbExclamation
        Exit Sub
    End If
    Set tgtWs = ActiveWorkbook.Sheets(1) ' Target
    Set srcWs = ActiveWorkbook.Sheets(2) ' Source
    
    Dim tgtLastRow As Long
    tgtLastRow = tgtWs.Cells(tgtWs.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Append (simple copy; assume same structure)
    srcWs.UsedRange.Copy Destination:=tgtWs.Cells(tgtLastRow, 1)
    
    MsgBox "Data from Sheet2 appended to Sheet1.", vbInformation
End Sub

Sub GenerateReport(Control As IRibbonControl)
    ' Simple Report Generation: Audit Log
    Dim reportSheet As Worksheet
    Set reportSheet = GetOrCreateSheet("Audit Report")
    reportSheet.Cells.Clear
    
    reportSheet.Cells(1, 1).Value = "AuditXcel AI Report"
    reportSheet.Cells(2, 1).Value = "Date: " & Now
    reportSheet.Cells(3, 1).Value = "Active Sheet: " & ActiveSheet.Name
    reportSheet.Cells(4, 1).Value = "Row Count: " & ActiveSheet.UsedRange.Rows.Count
    reportSheet.Cells(5, 1).Value = "Actions Performed: Data Cleaned, Summaries Generated" ' Placeholder
    
    MsgBox "Report generated in 'Audit Report' sheet.", vbInformation
End Sub

Sub FraudDetection(Control As IRibbonControl)
    ' Enhanced Placeholder: Basic Gap Detection (e.g., missing sequence in Col A)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim gapSheet As Worksheet
    Set gapSheet = GetOrCreateSheet("Gaps")
    gapSheet.Cells.Clear
    
    Dim prevVal As Long, currVal As Long, gapCount As Long
    gapCount = 0
    prevVal = ws.Cells(2, 1).Value ' Assume numeric sequence starting row 2
    For i = 3 To lastRow
        currVal = ws.Cells(i, 1).Value
        If currVal <> prevVal + 1 Then
            gapCount = gapCount + 1
            gapSheet.Cells(gapCount + 1, 1).Value = "Gap between " & prevVal & " and " & currVal
        End If
        prevVal = currVal
    Next i
    
    If gapCount > 0 Then
        MsgBox gapCount & " gaps found in 'Gaps' sheet. (Basic fraud check)", vbInformation
    Else
        MsgBox "No gaps detected.", vbInformation
    End If
End Sub

' Helper Functions
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function FindFirstNumericColumn(ws As Worksheet) As Integer
    Dim col As Integer
    For col = 1 To ws.UsedRange.Columns.Count
        If IsNumeric(ws.Cells(2, col).Value) Then ' Check row 2
            FindFirstNumericColumn = col
            Exit Function
        End If
    Next col
    FindFirstNumericColumn = 1 ' Default
End Function
