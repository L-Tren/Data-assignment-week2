Attribute VB_Name = "Module3"
Sub CalculateAnnualChangePrice()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As Variant
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim annualChange As Double
    Dim tickerDict As Object
    Dim tickerRow As Long
    Dim rowcount As Integer
    Dim Summary_Table_Row As Integer

    ' Set the worksheet to the desired sheet
    For Each ws In ThisWorkbook.Worksheets
    
    'Insert answer starting in column 2
    Summary_Table_Row = 2

    ' set rowcount to zero
    rowcount = 0

    ' Find the last row of data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Create a dictionary to store processed tickers
    Set tickerDict = CreateObject("Scripting.Dictionary")

    ' Loop through the rows
    For i = 2 To lastRow
        ' Get the ticker code
        ticker = ws.Cells(i, "A").Value

        ' Check if the ticker code has already been processed
        If Not tickerDict.Exists(ticker) Then
            ' Get the opening price
            openingPrice = ws.Cells(i, "C").Value

            ' Add the ticker code to the dictionary
            tickerDict.Add ticker, i
        End If

        ' Check if the ticker code changes
        If ws.Cells(i + 1, "A").Value <> ticker Then
            ' Get the closing price
            closingPrice = ws.Cells(i, "F").Value

            'get opening price
            openingPrice = ws.Cells((i - rowcount), "C").Value

            ' Calculate the annual change
            annualChange = closingPrice - openingPrice

            ' Get the row of the ticker code
            ' tickerRow = tickerDict(ticker)

            ' Output the annual change in column N for the ticker row
            ws.Cells(Summary_Table_Row, "N").Value = annualChange

            ' Apply conditional formatting to highlight negative values as red and positive values as green
            If annualChange < 0 Then
                ws.Cells(Summary_Table_Row, "N").Interior.Color = RGB(255, 0, 0) ' Red
            ElseIf annualChange > 0 Then
                ws.Cells(Summary_Table_Row, "N").Interior.Color = RGB(0, 255, 0) ' Green
            End If

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1

            rowcount = 0
        End If

        rowcount = rowcount + 1
    Next i
    Next ws
End Sub
