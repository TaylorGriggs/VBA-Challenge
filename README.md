# VBA-Challenge

# Location of Repo: https://github.com/TaylorGriggs/VBA-Challenge

# VBA Code for Alhpabetic_Testing
Sub Alphabetical_Testing():
    
    'Establish variables used for
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim Ticker As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim totalSV As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Volume_Total As LongLong
    
    
    Ticker = 0
    Summary_Table_Row = 2
    
    
    'Fill in header information
    For Each ws In Worksheets
        Dim worksheetName As String
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("O1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Value"
     

    'Display ticker type in Column "Ticker"
    
    
    'Last Row Counter
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Greatest % Creation
    Dim GreatInc As Double
    Dim GreatTick As String
    Dim LeastInc As Double
    Dim LeastTick As String
    Dim GreatVol As LongLong
    
    GreatInc = 0
    LeastInc = 0
    GreatVol = 0
    
    'Loop through all ticker data
        For i = 2 To LastRow
        
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            'We want the initial Open Price before the loop
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
            Open_Price = ws.Cells(i, 3).Value
            
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            'Set the ticker name
            TickerName = ws.Cells(i, 1).Value
            
            'Print Ticker Name is Summary Table
            ws.Range("I" & Summary_Table_Row).Value = TickerName
            
            'Set Close Price Var to match given info
            Close_Price = ws.Cells(i, 6).Value
            
            'Calculate yearly change
            YearlyChange = ws.Cells(i, 10).Value
            YearlyChange = Close_Price - Open_Price
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
            'Conditional Formatting for green if positive and red if negative
            If YearlyChange > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
            
            'Percent change calculation
            Percent_Change = ws.Cells(i, 11).Value
            Percent_Change = (YearlyChange / Open_Price)
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
            
            'Calculate greatest increase in percentage
            If Percent_Change > GreatInc Then
                GreatInc = Percent_Change
                GreatTick = TickerName
            End If
            
            'Calculate least change in %
            If Percent_Change < LeastInc Then
                LeastInc = Percent_Change
                LeastTick = TickerName
            End If
            
            If Volume_Total > GreatVol Then
                GreatVol = Volume_Total
                GreatVolTick = TickerName
            End If
            
        'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            Volume_Total = 0
            End If
            
            
        Next i
        Summary_Table_Row = 2
    'Display Greatest % Increase Info
    ws.Range("N" & Summary_Table_Row).Value = "Greatest % Increase"
    ws.Range("O" & Summary_Table_Row).Value = GreatTick
    ws.Range("P" & Summary_Table_Row).Value = GreatInc
    'Display Least % Increase Info
    ws.Range("N" & Summary_Table_Row + 1).Value = "Least % Increase"
    ws.Range("O" & Summary_Table_Row + 1).Value = LeastTick
    ws.Range("P" & Summary_Table_Row + 1).Value = LeastInc
    'Display Greatest Total Volume Info
    ws.Range("N" & Summary_Table_Row + 2).Value = "Greatest Total Volume"
    ws.Range("O" & Summary_Table_Row + 2).Value = GreatVolTick
    ws.Range("P" & Summary_Table_Row + 2).Value = GreatVol
    Next ws


End Sub
