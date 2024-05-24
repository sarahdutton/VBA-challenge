Sub Challenge2()
    
    For Each ws In Worksheets
    
        'Set variables
        
        Dim TickerName As String
        Dim TotalVolume As Double
        TotalVolume = 0

        Dim SummaryTickerRow As Integer
        SummaryTickerRow = 2
        
        Dim OpenPrice As Double
        OpenPrice = Cells(2, 3).Value
        Dim ClosePrice As Double
        
        Dim QtrChange As Double
        Dim PercentChange As Double
        
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVol As Double

        'Label the headers
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Find the Last Row
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop

        For i = 2 To LastRow

           
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              TickerName = Cells(i, 1).Value

              TotalVolume = TotalVolume + Cells(i, 7).Value

              'Print Ticker Name and Total Volume
              
              Range("I" & SummaryTickerRow).Value = TickerName
              Range("L" & SummaryTickerRow).Value = TotalVolume

              ClosePrice = Cells(i, 6).Value

              'Find Quarterly Change and Print
              
              QtrChange = (ClosePrice - OpenPrice)
              Range("J" & SummaryTickerRow).Value = QtrChange

                If (OpenPrice = 0) Then

                    PercentChange = 0

                Else
                    
                    PercentChange = QtrChange / OpenPrice
                
                End If

              'Print the Percent Change
              Range("K" & SummaryTickerRow).Value = PercentChange
              Range("K" & SummaryTickerRow).NumberFormat = "0.00%"
   
              'Reset Row Counter
              SummaryTickerRow = SummaryTickerRow + 1

              'Reset volume of trade to zero
              TotalVolume = 0

              'Reset the opening price
              OpenPrice = Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              TotalVolume = TotalVolume + Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting that will highlight positive change in green and negative change in red
    'First find the last row of the summary table

    LastRow_SummaryTable = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To LastRow_SummaryTable
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
            

        
        GreatestIncrease = ws.Range("Q2").Value
        GreatestDecrease = ws.Range("Q3").Value
        GreatestVol = ws.Range("Q4").Value
        
        
            
            'Greatest Increase
            If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = Cells(i, 9).Value
            
            Else
            
            GreatestIncrease = GreatestIncrease
            
            End If
            
            'Greatest Decrease
            If ws.Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
            Else
            
            GreatestDecrease = GreatestDecrease
            
            End If
    
            
            'Greatest Volume
            
            If ws.Cells(i, 12).Value > GreatestVol Then
            GreatestVol = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
            Else
            
            GreatestVol = GreatestVol
            
            End If
            
                        'Format
                        
        ws.Range("Q2").Value = Format(GreatestIncrease, "Percent")
        ws.Range("Q3").Value = Format(GreatestDecrease, "Percent")
        ws.Range("Q4").Value = Format(GreatestVol, "Scientific")
       
    Next i
    Next ws

End Sub

