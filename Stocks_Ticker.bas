Attribute VB_Name = "Module1"
Sub Stocks_Ticker()
    Dim ws As Worksheet
    Dim ticker As String
    Dim Ticker_Total As Double
    Dim Summary_Table_Row As Integer
    Dim lastRow As Long
    Dim i As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set initial values for each worksheet
        Ticker_Total = 0
        Summary_Table_Row = 2
        
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all tickers
        For i = 2 To lastRow
            ' Check if we are still within the same ticker, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker name
                ticker = ws.Cells(i, 1).Value
                
                ' Add to the ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                
                ' Print the ticker in the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = ticker
                
                ' Print the ticker Amount to the Summary Table
                ws.Range("P" & Summary_Table_Row).Value = Ticker_Total
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the ticker Total
                Ticker_Total = 0
            Else
                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub



