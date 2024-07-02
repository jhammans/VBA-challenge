Attribute VB_Name = "Module1"
Sub StockChecker()
    'Declare variables to read spreadsheet
    Dim stock_ticker As String
    Dim stock_date As String
    Dim stock_open As Variant
    Dim stock_close As Variant
    Dim stock_vol As Double
    
    Dim SheetName As String
    Dim LastRow As Long
    Dim LastCol As Long
    
    'Declare variables to write quarterly summaries
    Dim Ticker As String
    Dim QuarterlyChange As Variant
    Dim PercentChange As Variant
    Dim TotalStockVolume As Double
    Dim QuarterlyOpen As Variant
    Dim QuarterlyClose As Variant
    Dim Quarter As String
    Dim SummaryRow As Integer
        
    'Loop through each worksheet in the workbook
    For Each ws In Worksheets
    
        'Get worksheet name, last row and last column - https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/select-a-range
        SheetName = ws.Name
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                
        'Write column headings for calculated/aggregated values
        ws.Cells(1, LastCol + 2).Value = "Ticker"
        ws.Cells(1, LastCol + 3).Value = "Quarterly Change"
        ws.Cells(1, LastCol + 4).Value = "Percent Change"
        ws.Cells(1, LastCol + 5).Value = "TotalStockVolume"
        
        Ticker = ""
        QuarterlyChange = 0
        PercentChange = 0
        TotalStockVolume = 0
        SummaryRow = 2
        
        'Loop through rows
        For i = 2 To LastRow
            stock_ticker = ws.Cells(i, 1).Value
            stock_date = ws.Cells(i, 2).Value
            stock_open = ws.Cells(i, 3).Value
            stock_close = ws.Cells(i, 6).Value
            stock_vol = ws.Cells(i, 7).Value
            
            ' Determine the quarter based on the month
            Select Case Mid(stock_date, 5, 2)
                Case "01", "02", "03"
                    Quarter = "Q1"
                Case "04", "05", "06"
                    Quarter = "Q2"
                Case "07", "08", "09"
                    Quarter = "Q3"
                Case "10", "11", "12"
                    Quarter = "Q4"
            End Select
            
            If Ticker = "" Then
                Ticker = stock_ticker
                QuarterlyOpen = stock_open
                TotalStockVolume = stock_vol
            ElseIf Ticker = stock_ticker Then
                TotalStockVolume = TotalStockVolume + stock_vol
                QuarterlyClose = stock_close
            Else
                'Write summary and reset summary variables
                ws.Cells(SummaryRow, LastCol + 2).Value = Ticker
                
                ws.Cells(SummaryRow, LastCol + 3).Value = QuarterlyClose - QuarterlyOpen
                ws.Cells(SummaryRow, LastCol + 3).NumberFormat = "$#,##0.00"
                
                ws.Cells(SummaryRow, LastCol + 4).Value = ((QuarterlyClose - QuarterlyOpen) / QuarterlyOpen)
                ws.Cells(SummaryRow, LastCol + 4).NumberFormat = "0.00%"
                
                If ws.Cells(SummaryRow, LastCol + 4).Value > 0 Then
                    ws.Cells(SummaryRow, LastCol + 4).Interior.Color = vbGreen
                ElseIf ws.Cells(SummaryRow, LastCol + 4).Value < 0 Then
                    ws.Cells(SummaryRow, LastCol + 4).Interior.Color = vbRed
                End If
                
                ws.Cells(SummaryRow, LastCol + 5).Value = TotalStockVolume
                ws.Cells(SummaryRow, LastCol + 5).NumberFormat = "#,##0"
                
                'Reset summary variables
                SummaryRow = SummaryRow + 1
                Ticker = ""
                QuarterlyChange = 0
                PercentChange = 0
                TotalStockVolume = 0
            End If
        Next i
        
        Dim VolumeRange As Range
        Dim ChangeRange As Range
        Dim MaxStockVolume As Variant
        Dim MinStockChange As Variant
        Dim MaxStockChange As Variant
        Dim MaxStockRow As Variant
        Dim MinChangeRow As Variant
        Dim MaxChangeRow As Variant
        
        
        ' Set the range to find the aggregated values
        Set VolumeRange = Range("L2:L" & ws.Cells(ws.Rows.Count, 12).End(xlUp).Row)
        Set ChangeRange = Range("K2:K" & ws.Cells(ws.Rows.Count, 11).End(xlUp).Row)
        ' Find the maximum value in the range using the aggregation functions
        MaxStockVolume = Application.WorksheetFunction.Max(VolumeRange)
        MinStockChange = Application.WorksheetFunction.Min(ChangeRange)
        MaxStockChange = Application.WorksheetFunction.Max(ChangeRange)
        ' Find the row from the results of the matching aggregation
        MaxStockRow = Application.Match(MaxStockVolume, ws.Columns(12), 0)
        MinChangeRow = Application.Match(MinStockChange, ws.Columns(11), 0)
        MaxChangeRow = Application.Match(MaxStockChange, ws.Columns(11), 0)
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("P2").Value = ws.Range("I" & MaxChangeRow)
        ws.Range("Q2").Value = MaxStockChange
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("P3").Value = ws.Range("I" & MinChangeRow)
        ws.Range("Q3").Value = MinStockChange
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("O4").Value = "Greatest total volume"
        ws.Range("P4").Value = ws.Range("I" & MaxStockRow)
        ws.Range("Q4").Value = MaxStockVolume
        ws.Range("Q4").NumberFormat = "#,##0"
        Exit For
    Next ws
End Sub
