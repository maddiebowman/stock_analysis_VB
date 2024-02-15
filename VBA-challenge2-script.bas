Attribute VB_Name = "Module1"
Sub Stock_Analysis():
 
'Loop through all the sheets.
    For Each ws In Worksheets
 
'create variables
        Dim Ticker As String
        Dim DtockVolume As Double
        DtockVolume = 0
        Dim OpenPrice As Double
        OpenPrice = ws.Cells(2, 3).Value
        Dim ClosePrice As Double
        Dim PercentChange As Double
        Dim YearlyChange As Double
 
        
        'ticker name in the summary
        Dim SummaryTickerRow As Integer
        SummaryTickerRow = 2

 
    ' Create Headers
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"

 
        ' Define last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    ' Loop through worksheet
        For I = 2 To LastRow
 
            ' Has ticker changed?
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            'Or i = 2 Then
              ' Reset variables for new ticker
              Ticker = ws.Cells(I, 1).Value
              DtockVolume = DtockVolume + ws.Cells(I, 7).Value
 
               
              'summary table
              ws.Range("I" & SummaryTickerRow).Value = Ticker
              ws.Range("L" & SummaryTickerRow).Value = DtockVolume
              ClosePrice = ws.Cells(I, 6).Value
              YearlyChange = (ClosePrice - OpenPrice) ' F256 - C2
              ws.Range("J" & SummaryTickerRow).Value = YearlyChange
 
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
 
              'summary table
              ws.Range("K" & SummaryTickerRow).Value = PercentChange
              ws.Range("K" & SummaryTickerRow).NumberFormat = "0.00%"
              'Reset row counter
              SummaryTickerRow = SummaryTickerRow + 1
 
              'Reset volume
              DtockVolume = 0
 
              'Reset the price
              OpenPrice = ws.Cells(I + 1, 3)
            Else
               'Add volume
              DtockVolume = DtockVolume + ws.Cells(I, 7).Value
            End If
        Next I
 
    'summary table
    Summary_Table_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For I = 2 To Summary_Table_LastRow
            If ws.Cells(I, 10).Value > 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(I, 10).Interior.ColorIndex = 3
            End If
        Next I
 
 
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
 
        For I = 2 To Summary_Table_LastRow
            If ws.Cells(I, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_LastRow)) Then
                ws.Cells(2, 16).Value = ws.Cells(I, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(I, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
 
            ElseIf ws.Cells(I, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_LastRow)) Then
                ws.Cells(3, 16).Value = ws.Cells(I, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(I, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(I, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_LastRow)) Then
                ws.Cells(4, 16).Value = ws.Cells(I, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(I, 12).Value
            End If
        Next I
    Next ws
End Sub
