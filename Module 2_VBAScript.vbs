Sub Pt1()

    Dim ws As Integer
    Dim ws_count As Integer
    ws_count = ActiveWorkbook.Worksheets.Count
    
    For ws = 1 To ws_count
    
    ThisWorkbook.Worksheets(ws).Activate
        
        'Insert new
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Dim Ticker As String
        Dim totalvolume As Double
        totalvolume = 0
        OpenPriceRow = 2
    
            
        'Summary Table
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range("K2:K" & LastRow).NumberFormatLocal = "0.00%"
        Range("J2:J" & LastRow).NumberFormat = "0.00"
                
        'loop through Ticker
        For i = 2 To LastRow
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Ticker = Cells(i, 1).Value
                OpenPrice = Cells(OpenPriceRow, 3).Value
                ClosePrice = Cells(i, 6).Value
                
                'Calculate yearly change, percentage change, and total volume
                yearly_change = Cells(i, 6).Value - OpenPrice
            
                percent_change = (Cells(i, 6).Value - OpenPrice) / OpenPrice
                
                totalvolume = totalvolume + Cells(i, 7).Value
        
                 
                'Print to summary table
                Range("I" & SummaryTableRow).Value = Ticker
                Range("J" & SummaryTableRow).Value = yearly_change
                Range("K" & SummaryTableRow).Value = percent_change
                Range("L" & SummaryTableRow).Value = totalvolume
                SummaryTableRow = SummaryTableRow + 1
        
                'Reset
                Ticker = Cells(i + 1, 1).Value
                OpenPriceRow = i + 1
                totalvolume = 0
            
            Else
                
                totalvolume = totalvolume + Cells(i, 7).Value
            
            End If
        
        Next i
    
    
        'Greatest %
        lastSummaryTableRow = Cells(Rows.Count, 10).End(xlUp).Row
        Dim greatVol As Double
        Dim greatpercentIn As Double
        Dim greatpercentDe As Double
        
        Dim tickerIn As String
        Dim tickerDe As String
        Dim tickerVol As String
        
        For j = 2 To lastSummaryTableRow
            
            If Cells(j, 11).Value > greatpercentIn Then
                greatpercentIn = Cells(j, 11).Value
                tickerIn = Cells(j, 9).Value
            End If
            
            If Cells(j, 11).Value < greatpercentDe Then
                greatpercentDe = Cells(j, 11).Value
                tickerDe = Cells(j, 9).Value
            End If
            
            If Cells(j, 12).Value > greatVol Then
                greatVol = Cells(j, 12).Value
                tickerVol = Cells(j, 9).Value
            End If
            
        Next j
        
        Range("P2").Value = tickerIn
        Range("P3").Value = tickerDe
        Range("P4").Value = tickerVol
        
        Range("Q2").Value = greatpercentIn
        Range("Q3").Value = greatpercentDe
        Range("Q2:Q" & 3).NumberFormatLocal = "0.00%"
        Range("Q4").Value = greatVol
            
    Next ws
            
    MsgBox ("Complete")

End Sub
