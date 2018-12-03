Sub StockVolume()

'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

' Create a script that will loop through all the stocks and take the following info.
  ' Yearly change from what the stock opened the year at to what the closing price was.
  ' The percent change from the what it opened the year at to what it closed.
  ' The total Volume of the stock
  ' Ticker symbol
' You should also have conditional formatting that will highlight positive change in green and negative change in red.

' Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

' Make the appropriate adjustments to your script that will allow it to run on every worksheet just by running it once.


    ' ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' ' --------------------------------------------
        ' Understand the size of the data of each sheet
        ' --------------------------------------------
        ' Determine the Last Row Number
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Determine the Last Column Number.
        'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        'First Built the program with Cells function and LastColumn. But later found that Range function is much easier for reviewing when there is no nest loop.
        'And since we are not treating a random data set, the dataset's last Coulumn number is set to be 7 or "G". And the table locations are also predetermined.
        
        ' ' ----------------------------------------------------------------------------
        ' Create the Table of each stock's summary and the Table of the signature stocks
        ' ------------------------------------------------------------------------------
        ws.Range("I1:Q1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
        ws.Range("O2:O4").Value = Application.WorksheetFunction.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Volume"))
        
        ' ' --------------------------------------------
        ' Fill the Table of each stock's summary
        ' --------------------------------------------
        'The number of the stocks
        Dim NumOfStock As Long
        NumOfStock = 1
        'The row number of one stock's starting row and ending row
        Dim STRow, ENrow As Long
        STRow = 2
        'Stock Name
        Dim SotckName As String
        'StockName = ws.Range("A2").Value
                
        For i = 2 To LastRow
            'Use contidioning to find the stock end row. It can use name or dates. Here is using the dates.
            If ws.Range("B" & i).Value > ws.Range("B" & (i + 1)).Value Then
            'If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'The conditioning happens at the end row of this stock
                ENrow = i
                
                'Give this stock's name
                StockName = ws.Range("A" & i).Value
            
                'Fill the Table with Stock Name and add the number of the stock
                NumOfStock = NumOfStock + 1
                ws.Range("I" & NumOfStock).Value = StockName
                            
                'Define and Calculate the Yearly Change and Ratio of price change
                Dim StockInitial, StockClose, YearlyChange, ChangeRatio As Double
                
                StockInitial = ws.Range("C" & STRow).Value
                StockClose = ws.Range("F" & ENrow).Value
                
                    'For the stock started public at the mid of the year and document the value as zero before public
                    If StockInitial = 0 Then
                        For j = STRow To ENrow
                            If ws.Range("C" & j).Value <> 0 Then
                            StockInitial = ws.Range("C" & j).Value
                            Exit For
                            End If
                        Next j
                    End If
                
				'Calculate the Yearly Change and put into the table
                YearlyChange = StockClose - StockInitial
                ws.Range("J" & NumOfStock).Value = YearlyChange
                
                    'Highlights the positive and negative changes with green and red
                    If YearlyChange >= 0 Then
                    ws.Range("J" & NumOfStock).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & NumOfStock).Interior.ColorIndex = 3
                    End If
                
                'Calculate the Ratio of price change
                'For the possibility that an abnormal stock's value is 0 for the whole year.
                    If StockInitial = 0 Then
                    ws.Range("K" & NumOfStock).Value = ""
                    Else
                    ChangeRatio = YearlyChange / StockInitial
                    ws.Range("K" & NumOfStock).Value = ChangeRatio
                    ws.Range("K" & NumOfStock).NumberFormat = "0.00%"
                    End If
                
                
                
                'Define trading volume and calculate the sum
                Dim TRVolume As Double
                TRVolume = Application.Sum(ws.Range("G" & STRow, "G" & ENrow))
                ws.Range("L" & NumOfStock).Value = TRVolume
                
                'Start row change the next stock
                STRow = i + 1
       
            End If
        Next i
        
        ' ' --------------------------------------------------------------------------
        ' Table of the signature stocks of the year
        ' ----------------------------------------------------------------------------
                
        'Find the size of the table of the stock created above
        LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Define the value of the greatest increasing and decreasing stock, the volume value of the most traded stock, and give the initial to zero in case tha first stock is the goal
        Dim maxvalue, minvalue, MaxTotalVol As Double
        maxvalue = 0
        minvalue = 0
        MaxTotalVol = 0
        
        'Define the name of the corresponding stocks
        Dim maxStock, minStock, VolStock As String
                      
        'Find the values
        For i = 2 To LastRow2
            If ws.Range("K" & i).Value > maxvalue Then
            maxvalue = ws.Range("K" & i).Value
            maxStock = ws.Range("I" & i).Value
           
            ElseIf ws.Range("K" & i).Value < minvalue Then
            minvalue = ws.Range("K" & i).Value
            minStock = ws.Range("I" & i).Value
            
            End If
            
            If ws.Range("L" & i).Value > MaxTotalVol Then
            MaxTotalVol = ws.Range("L" & i).Value
            VolStock = ws.Range("I" & i).Value
                        
            End If
           
        Next i
        
		'Fill the Table of the siginature stocks
        ws.Range("P2").Value = maxStock
        ws.Range("Q2").Value = maxvalue
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3").Value = minStock
        ws.Range("Q3").Value = minvalue
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4").Value = VolStock
        ws.Range("Q4").Value = MaxTotalVol
            
        ws.Columns.AutoFit
    Next ws

End Sub
