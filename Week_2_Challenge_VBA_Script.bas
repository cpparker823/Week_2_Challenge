Attribute VB_Name = "Module1"
Sub YearlyChange():

'Create Workbook Loop
Dim ws As Worksheet
For Each ws In Worksheets

    'Define variables and key values
    
        Dim Ticker As String
        Dim PrevTicker As String
        Dim NextTicker As String
        Dim YearlyChange As Double
        Dim TickerOpen As Double
        Dim TickerClose As Double
        Dim TotalStockVolume As Double
        Dim LastRow As Double
        Dim PercentChange As Double
        Dim GIncV As Double
        Dim GDecV As Double
        Dim GTotV As Double
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        GIncV = 0
        GDecV = 0
        GTotV = 0
    
    'Create analysis tables
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Range("O3") = "Greatest % Increase"
    ws.Range("O4") = "Greatest % Decrease"
    ws.Range("O5") = "Greatest Total Volume"
    ws.Range("P2") = "Ticker"
    ws.Range("Q2") = "Value"
  
'Create Worksheet Loop
Dim i As Single
For i = 2 To LastRow
   
   'Define next, current, and previous ticker
    NextTicker = ws.Cells(i + 1, 1).Value
    Ticker = ws.Cells(i, 1).Value
    PrevTicker = ws.Cells(i - 1, 1).Value
   
    'Add to the ticker total volume
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    'Open value will be found in first row of new ticker
    If PrevTicker <> Ticker Then
        TickerOpen = ws.Cells(i, 3).Value
        
    'Close value will be found in last row before new ticker symbol
    ElseIf NextTicker <> Ticker Then
        TickerClose = ws.Cells(i, 6).Value
    
    'Calculate the yearly change once we have the open and close values
        YearlyChange = TickerClose - TickerOpen
        
    'Print the ticker symbol in analysis table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
            
    'Print the Yearly Change to the analysis table
        ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        
        'Color code red if negative and green if positive
            If YearlyChange < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf YearlyChange >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
    'Print the volume to the analysis table
        ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
    
    'Calculate and print the percent change into the analysis table
       PercentChange = YearlyChange / TickerOpen
       
       ws.Range("K" & Summary_Table_Row).Value = PercentChange
       
    'Store the new value if it meets GREATEST total value or increase or decrease percentage
    'Print the new value and ticker symbol in the table
       'Increase
            If PercentChange > GIncV Then
            GIncV = PercentChange
            
            ws.Range("Q3").Value = GIncV
            ws.Range("P3").Value = Ticker
            
       'Decrease
            ElseIf PercentChange < GDecV Then
            GDecV = PercentChange
            
            ws.Range("Q4").Value = GDecV
            ws.Range("P4").Value = Ticker
            
            End If
       
       'Total Volume
            If TotalStockVolume > GTotV Then
            GTotV = TotalStockVolume
            
            ws.Range("Q5").Value = GTotV
            ws.Range("P5").Value = Ticker
            End If
       
    'Reset stock volume to 0
        TotalStockVolume = 0
       
    'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
  
    End If

Next i

    'Format cells to display percentages
    ws.Range("k2:k" & LastRow).NumberFormat = "0.00%"
    ws.Range("Q3,Q4").NumberFormat = "0.00%"

Next ws

End Sub

