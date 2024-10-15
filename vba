Sub CompareSheetsAndWriteDetailedDifferences()
    ' Call the comparison function for each pair of sheets
    CompareSheetsDetailed "RVU_CFG", "RVU_PROD", "RVU Differences"
    ' You can add more calls here for other sheet pairs if needed
End Sub

Sub CompareSheetsDetailed(sheet1Name As String, sheet2Name As String, resultSheetName As String)
    Dim ws1 As Worksheet, ws2 As Worksheet, wsResult As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
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
    
    ' Get the last rows for both sheets
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    
    ' Copy headers to the result sheet and add a column for "Source Sheet"
    wsResult.Cells(1, 1).Value = "Rawdata"
    wsResult.Cells(1, 2).Value = "Count"
    wsResult.Cells(1, 3).Value = "Difference Description"
    
    ' Initialize variables
    resultRow = 2
    diffFound = False
    
    ' Compare each row in Sheet1 against Sheet2
    For i = 2 To lastRow1 ' Assuming row 1 contains headers
        ' Check if the row in Sheet1 matches the corresponding row in Sheet2
        If i <= lastRow2 Then
            If ws1.Cells(i, 1).Value = ws2.Cells(i, 1).Value Then
                ' Data spans match, check the count
                If ws1.Cells(i, 2).Value <> ws2.Cells(i, 2).Value Then
                    diffFound = True
                    ' Log the count mismatch
                    wsResult.Cells(resultRow, 1).Value = ws1.Cells(i, 1).Value ' Rawdata
                    wsResult.Cells(resultRow, 2).Value = ws1.Cells(i, 2).Value ' Count from RVU_CFG
                    wsResult.Cells(resultRow, 3).Value = "Data Span Match. Count mismatch between " & sheet1Name & " & " & sheet2Name
                    resultRow = resultRow + 1
                End If
            Else
                ' Log the existence in RVU_CFG but not matching in RVU_PROD
                diffFound = True
                wsResult.Cells(resultRow, 1).Value = ws1.Cells(i, 1).Value
                wsResult.Cells(resultRow, 2).Value = ws1.Cells(i, 2).Value
                wsResult.Cells(resultRow, 3).Value = "Exist in " & sheet1Name & ". Missing in " & sheet2Name
                resultRow = resultRow + 1
            End If
        Else
            ' Log the extra row from Sheet1 if Sheet2 has fewer rows
            diffFound = True
            wsResult.Cells(resultRow, 1).Value = ws1.Cells(i, 1).Value
            wsResult.Cells(resultRow, 2).Value = ws1.Cells(i, 2).Value
            wsResult.Cells(resultRow, 3).Value = "Exist in " & sheet1Name & ". Missing in " & sheet2Name
            resultRow = resultRow + 1
        End If
    Next i

    ' Compare each row in Sheet2 against Sheet1 to find any unmatched rows from Sheet2
    For i = 2 To lastRow2 ' Assuming row 1 contains headers
        If i > lastRow1 Or (ws1.Cells(i, 1).Value <> ws2.Cells(i, 1).Value) Then
            diffFound = True
            ' Log the existence in RVU_PROD but not matching in RVU_CFG
            wsResult.Cells(resultRow, 1).Value = ws2.Cells(i, 1).Value
            wsResult.Cells(resultRow, 2).Value = ws2.Cells(i, 2).Value
            wsResult.Cells(resultRow, 3).Value = "Exist in " & sheet2Name & ". Missing in " & sheet1Name
            resultRow = resultRow + 1
        End If
    Next i

    ' If no differences found, write the message
    If Not diffFound Then
        wsResult.Cells(1, 1).Value = "No Differences found, All Records Match"
    End If
End Sub

