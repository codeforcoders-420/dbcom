Sub CompareSheetsAndLogDifferences()
    ' Call the comparison function for each pair of sheets
    CompareSheets "RVU_CFG", "RVU_PROD", "RVU Differences"
    CompareSheets "GCPI_CFG", "GCPI_PROD", "GCPI Differences"
    CompareSheets "NAT_CFG", "NAT_PROD", "NAT Differences"
    CompareSheets "ZIP_CFG", "ZIP_PROD", "ZIP Differences"
End Sub

Sub CompareSheets(sheet1Name As String, sheet2Name As String, resultSheetName As String)
    Dim ws1 As Worksheet, ws2 As Worksheet, wsResult As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastCol1 As Long, lastCol2 As Long
    Dim i As Long, j As Long
    Dim diffFound As Boolean
    Dim resultRow As Long

    ' Set the worksheets
    Set ws1 = ThisWorkbook.Sheets(sheet1Name)
    Set ws2 = ThisWorkbook.Sheets(sheet2Name)
    
    ' Check if result sheet exists, if not, create it
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets(resultSheetName)
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsResult.Name = resultSheetName
    Else
        wsResult.Cells.Clear ' Clear previous results
    End If
    
    ' Get the last rows and columns
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column
    
    ' Ensure the data range matches
    If lastRow1 <> lastRow2 Or lastCol1 <> lastCol2 Then
        wsResult.Cells(1, 1).Value = "Data structure mismatch between " & sheet1Name & " and " & sheet2Name
        Exit Sub
    End If

    ' Copy headers to the result sheet
    For j = 1 To lastCol1
        wsResult.Cells(1, j).Value = ws1.Cells(1, j).Value
    Next j
    
    ' Initialize variables
    resultRow = 2
    diffFound = False
    
    ' Compare each row
    For i = 2 To lastRow1 ' Assuming row 1 contains headers
        For j = 1 To lastCol1
            If ws1.Cells(i, j).Value <> ws2.Cells(i, j).Value Then
                diffFound = True
                ' If a difference is found, copy the entire row from Sheet1
                For j = 1 To lastCol1
                    wsResult.Cells(resultRow, j).Value = ws1.Cells(i, j).Value
                Next j
                resultRow = resultRow + 1
                Exit For
            End If
        Next j
    Next i
    
    ' If no differences found, write the message
    If Not diffFound Then
        wsResult.Cells(1, 1).Value = "No Differences found, All Records Match"
    End If
End Sub
