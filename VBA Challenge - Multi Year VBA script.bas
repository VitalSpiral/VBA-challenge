Attribute VB_Name = "Module1"
Sub Stock_Challenge()

    'Loop through all sheets
    
    For Each ws In Worksheets
    
        ' Setup column and results cells labels and formats for worksheet(year)
        
        ws.Range("I1") = "Ticker    "
        
        ws.Range("J1") = "Yearly Change"
        
        ws.Range("K1") = "Percent Change"
        
        ws.Range("L1") = "Total Stock Volume"
        
        ws.Range("O2") = "Greatest % Increase"
        
        ws.Range("Q2").Style = "Percent"
        
        ws.Range("Q2").NumberFormat = "##.##%"
        
        ws.Range("O3") = "Greatest % Decrease"
        
        ws.Range("Q3").Style = "Percent"
        
        ws.Range("Q3").NumberFormat = "##.##%"
        
        ws.Range("O4") = "Greatest Total Volume"
        
        ws.Range("Q4").NumberFormat = "###################"
        
        ws.Range("O1") = "                                      "
        
        ws.Range("P1") = "Ticker     "
        
        ws.Range("Q1") = "Value              "
        
        ws.Columns("A:Q").EntireColumn.AutoFit
        
        ws.Range("K2:K10000").Style = "Percent"
        
        ws.Range("K2:K10000").NumberFormat = "##.##%"
        
    
        'Declare variables for worksheet and initialize where needed.
        
        Dim StockTicker As String
    
        Dim StockVolumeTotal As Double
        
        StockVolumeTotal = 0
    
        Dim StockTableRow As Integer
        
        StockTableRow = 2
    
        Dim DayCount As Integer
    
        DayCount = 0
    
        Dim FirstClose As Double
        
        Dim LastClose As Double
        
        Dim YearlyClose As Double
        
        Dim PercentChange As Variant
    
        Dim GreatestIncreaseTicker As String
        
        Dim GreatestPercentIncrease As Variant
        
        Dim GreatestDecreaseTicker As String
        
        Dim GreatestPercentDecrease As Variant
        
        Dim GreatestTotalVolumeTicker As String
        
        Dim GreatestTotalVolume As Double
    
        'Determine the Last Row of worksheet
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop to locate individual stock day ranges and calculate Yearly Change, Percent Change and Total Stock Volume
        
        For i = 2 To lastRow
        
            'First If locates Stocker Ticker change, grabs Stocker Ticker and Last Close, and calculates Yearly Change
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                StockTicker = ws.Cells(i, 1).Value
            
                LastClose = ws.Cells(i, 6).Value
            
                YearlyChange = LastClose - FirstClose
            
                'This If catches the overflow case when First Close = 0 and returns blank Percent Change value.
                
                If FirstClose = 0 Then
            
                PercentChange = " "
                
                'Otherwise Percent Change calculated as normally done
                
                Else
            
                PercentChange = YearlyChange / FirstClose
            
                End If
                
                'Adding last closing day volume to Stock Volume Total
            
                StockVolumeTotal = StockVolumeTotal + ws.Cells(i, 7).Value
            
                'Assigning values to Stock Ticker, Stock Total Volume, Yearly Change and Percent Change
                
                ws.Range("I" & StockTableRow).Value = StockTicker
            
                ws.Range("L" & StockTableRow).Value = StockVolumeTotal
            
                ws.Range("J" & StockTableRow).Value = YearlyChange
            
                ws.Range("K" & StockTableRow).Value = PercentChange
            
                'This If formats Stock Yearly Change cells interior color red, green or white(if blank)
                                
                If YearlyChange < 0 Then
            
                    ws.Range("J" & StockTableRow).Interior.ColorIndex = 3
                
                ElseIf YearlyChange > 0 Then
            
                    ws.Range("J" & StockTableRow).Interior.ColorIndex = 4
                
                Else: ws.Range("J" & StockTableRow).Interior.ColorIndex = 0
            
                End If
                               
                'Increment the Stock Table row and reset Total Stock Volume and reset count of days in year
                
                StockTableRow = StockTableRow + 1
            
                StockVolumeTotal = 0
            
                DayCount = 0
            
            Else
        
                'If Stock Ticker is same add to Stock Volume Total and increment day count in year
                
                StockVolumeTotal = StockVolumeTotal + ws.Cells(i, 7).Value
            
                DayCount = DayCount + 1
            
                'This If catches the First Stock Close value for first day of year
                
                If DayCount = 1 Then
                
                FirstClose = ws.Cells(i, 6).Value
                
                End If
           
            End If
        
        Next i
        
        'Determine Last Row of new Stock Table
        
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        'Grab values from new Stock Table for first Stock Ticker to initialize values for Greatest Increse/Decrease and Largest Stock Volume
        
        GreatestPercentIncrease = ws.Cells(2, 11).Value
                  
        GreatestPercentDecrease = ws.Cells(2, 11).Value
         
        GreatestTotalVolume = ws.Cells(2, 12).Value
        
        For j = 2 To lastrow2
        
            'This If makes Percent Change "0" if set as " "
            
            If ws.Cells(j + 1, 11).Value = " " Then
            
                ws.Cells(j + 1, 11).Value = 0
                
            End If
        
            'This If replaces Ticker and Increase If Percent Change is greater
            
            If ws.Cells(j + 1, 11).Value >= GreatestPercentIncrease Then
                        
                GreatestIncreaseTicker = ws.Cells(j + 1, 9).Value
                GreatestPercentIncrease = ws.Cells(j + 1, 11).Value
                
            End If
        
            'This If replaces Ticker and Decrease If Percent Change is smaller
            
            If ws.Cells(j + 1, 11).Value <= GreatestPercentDecrease Then
        
                GreatestDecreaseTicker = ws.Cells(j + 1, 9).Value
                GreatestPercentDecrease = ws.Cells(j + 1, 11).Value
                     
            End If
         
            'This If replaces Ticker and Total Volume If Stock Total Volume is greater
            
            If ws.Cells(j + 1, 12).Value > GreatestTotalVolume Then
        
                GreatestTotalVolumeTicker = ws.Cells(j + 1, 9).Value
                GreatestTotalVolume = ws.Cells(j + 1, 12).Value
            
            End If
                
        Next j
        
            'Assigning values to cells for Greatest Increase/Decrease and Greatest Stock Volume Total
            
            ws.Range("P2") = GreatestIncreaseTicker
            ws.Range("Q2") = GreatestPercentIncrease
    
            ws.Range("P3") = GreatestDecreaseTicker
            ws.Range("Q3") = GreatestPercentDecrease
    
            ws.Range("P4") = GreatestTotalVolumeTicker
            ws.Range("Q4") = GreatestTotalVolume
       
   Next ws
   
End Sub

