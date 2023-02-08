Sub stocks_3()
'to make the code loop all sheets at once
For Each ws In Worksheets
    
    'declaring all my variables
    Dim ticker As String
    Dim row As Long
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim totalvol As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yarly_change As Double
    Dim perc_change As Double
    Dim max As Double
    Dim min As Double
    Dim greatest As Double
        
        i = 2
        close_price = 0
        open_price = Cells(i, 3).Value
        totalvol = 0
        yearly_change = 0
        max = 0
        min = 0
        greatest = 0
        
        'finding the last row
        row = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'formatting columns as percent and currency
        ws.Range("K2:K" & row).NumberFormat = "0.00%"
        ws.Range("J2:J" & row).NumberFormat = "$ #,##0.00"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'creating all new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Columns("I:U").AutoFit

'looping through all the sheets of the workbook to extract the values
For i = 2 To row

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ticker = ws.Cells(i, 1).Value
        totalvol = totalvol + ws.Cells(i, 7).Value
        open_price = open_price
        close_price = ws.Cells(i, 6).Value
        yearly_change = close_price - open_price
        perc_change = yearly_change / open_price
        
        'printing all the information requested into the new columns
        ws.Range("J" & Summary_Table_Row).Value = yearly_change
            'conditional formatting for the interior color depending on the values
            If yearly_change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            Else
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            End If
            
        ws.Range("K" & Summary_Table_Row).Value = perc_change
            If perc_change > 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            Else
            ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            End If

        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("L" & Summary_Table_Row).Value = totalvol

        Summary_Table_Row = Summary_Table_Row + 1
        open_price = ws.Cells(i + 1, 3).Value

        totalvol = 0
    Else

        totalvol = totalvol + Cells(i, 7).Value
    End If
Next i
'fits the column size to the lenght of the new text that appears
ws.Columns("Q").AutoFit

'extracting greatest % increase and printing it
For i = 2 To row

    If ws.Cells(i, 11).Value > max Then
    max = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Else
    ws.Cells(2, 17).Value = max
    
    End If
    Next i
    'fits the column size to the lenght of the text
    ws.Columns("Q").AutoFit
    
'extracting greatest % decrease and printing it
For i = 2 To row

    If ws.Cells(i, 11).Value < min Then
    min = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Else
    ws.Cells(3, 17).Value = min
    
    End If
    Next i
'extracting greatest total volume and printing it
For i = 2 To row

    If ws.Cells(i, 12).Value > greatest Then
    greatest = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Else
    ws.Cells(4, 17).Value = greatest
    
    End If
    Next i
    ws.Columns("Q").AutoFit

Next ws
End Sub




