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
    Dim PrevQuarter As String
    Dim SummaryRow As Integer
        
    'Loop through each worksheet in the workbook
    For Each ws In Worksheets
    
        'Get worksheet name, last row and last column - https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/select-a-range
        SheetName = ws.Name
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        Ticker = ""
        QuarterlyChange = 0
        PercentChange = 0
        TotalStockVolume = 0
        SummaryRow = 2
        PrevQuarter = ""
        
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
                'Write quarterly summary and reset quarterly variables
                ws.Cells(SummaryRow, 9).Value = Ticker
                
                ws.Cells(SummaryRow, 10).Value = QuarterlyClose - QuarterlyOpen
                ws.Cells(SummaryRow, 10).NumberFormat = "$#,##0.00"
                
                ws.Cells(SummaryRow, 11).Value = ((QuarterlyClose - QuarterlyOpen) / QuarterlyOpen)
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                If ws.Cells(SummaryRow, 11).Value > 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = vbGreen
                ElseIf ws.Cells(SummaryRow, 11).Value < 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = vbRed
                End If
                
                ws.Cells(SummaryRow, 12).Value = TotalStockVolume
                ws.Cells(SummaryRow, 12).NumberFormat = "#,##0"
                
                ws.Cells(SummaryRow, 14).Value = QuarterlyOpen
                ws.Cells(SummaryRow, 15).Value = QuarterlyClose
                
                'reset quarterly variables
                SummaryRow = SummaryRow + 1
                Ticker = ""
                QuarterlyChange = 0
                PercentChange = 0
                TotalStockVolume = 0
                PrevQuarter = Quarter
            End If
        Next i
    Next ws
End Sub
