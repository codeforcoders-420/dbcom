Sub CompareSheets()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsResult As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, resultRow As Long
    Dim data1 As Range, data2 As Range
    Dim i As Long, j As Long
    Dim foundMatch As Boolean
    Dim countMismatch As Boolean
    Dim count1 As Long, count2 As Long
    Dim rawData1 As String, rawData2 As String
    Dim message As String
    
    ' Set your sheet references
    Set ws1 = ThisWorkbook.Sheets("RVU_CFG")
    Set ws2 = ThisWorkbook.Sheets("RVU_PROD")
    Set wsResult = ThisWorkbook.Sheets("RVU Differences")
    
    ' Clear previous results
    wsResult.Cells.Clear
    wsResult.Cells(1, 1).Value = "Data"
    wsResult.Cells(1, 2).Value = "Count"
    wsResult.Cells(1, 3).Value = "Source"
    resultRow = 2 ' Start from the second row
    
    ' Get the last row of data in each sheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in the first sheet
    For i = 2 To lastRow1
        rawData1 = ws1.Cells(i, 1).Value
        count1 = ws1.Cells(i, 2).Value
        foundMatch = False
        countMismatch = False
        
        ' Check if this rawData exists in the second sheet
        For j = 2 To lastRow2
            rawData2 = ws2.Cells(j, 1).Value
            count2 = ws2.Cells(j, 2).Value
            
            ' If a match is found
            If rawData1 = rawData2 Then
                foundMatch = True
                If count1 <> count2 Then
                    countMismatch = True
                End If
                Exit For ' No need to check further for this rawData
            End If
        Next j
        
        ' Log results based on the checks
        If foundMatch Then
            If countMismatch Then
                wsResult.Cells(resultRow, 1).Value = rawData1
                wsResult.Cells(resultRow, 2).Value = count1
                wsResult.Cells(resultRow, 3).Value = "Data Span Match. Count mismatch between RVU_CFG & RVU_PROD"
                resultRow = resultRow + 1
            End If
        Else
            wsResult.Cells(resultRow, 1).Value = rawData1
            wsResult.Cells(resultRow, 2).Value = count1
            wsResult.Cells(resultRow, 3).Value = "Exist in RVU_CFG. Missing in RVU_PROD"
            resultRow = resultRow + 1
        End If
    Next i
    
    ' Now check for any missing rows in RVU_CFG from RVU_PROD
    For j = 2 To lastRow2
        rawData2 = ws2.Cells(j, 1).Value
        count2 = ws2.Cells(j, 2).Value
        foundMatch = False
        
        For i = 2 To lastRow1
            rawData1 = ws1.Cells(i, 1).Value
            count1 = ws1.Cells(i, 2).Value
            
            ' Check for existence of both Rawdata and Count
            If rawData2 = rawData1 And count2 = count1 Then
                foundMatch = True
                Exit For
            End If
        Next i
        
        ' Log results based on the checks
        If Not foundMatch Then
            wsResult.Cells(resultRow, 1).Value = rawData2
            wsResult.Cells(resultRow, 2).Value = count2
            wsResult.Cells(resultRow, 3).Value = "Exist in RVU_PROD. Missing in RVU_CFG"
            resultRow = resultRow + 1
        End If
    Next j
    
    ' Check if any differences were found
    If resultRow = 2 Then
        wsResult.Cells(2, 1).Value = "No Differences found. All Records Match"
    End If
End Sub

