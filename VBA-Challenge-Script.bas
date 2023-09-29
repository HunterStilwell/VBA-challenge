' Attribute VB_Name = "Module1"
Sub VBAChallenge():

' loop that goes through every sheet in the workbook
For Each ws In Worksheets
        ' set the values for the new columns that need to be created
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        ' create a variable to hold the last row
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' create variables to hold information
        Dim ticker As String
        Dim totalStock As Double
        Dim firstOpen As Double
        Dim yearChange As Double
        Dim perChange As Double
        Dim summary As Integer
    
        ' set the variables at 0 (2 for summary)
        totalStock = 0
        firstOpen = 0
        lastClose = 0
        yearChange = 0
        perChange = 0
        summary = 2
    
        ' create a loop to go through rows 2 to the last row
        For Row = 2 To lastRow
            ' first condition
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
        
                ' put ticker value in Column I
                ticker = ws.Cells(Row, 1).Value
                ws.Cells(summary, 9).Value = ticker
            
                ' finish adding and put total stock amount in Column L
                totalStock = totalStock + ws.Cells(Row, 7).Value
                ws.Cells(summary, 12).Value = totalStock
            
                ' record the last close amount and determine the yearly change
                lastClose = ws.Cells(Row, 6).Value
                yearChange = lastClose - firstOpen
                ws.Cells(summary, 10).Value = yearChange
            
                    ' conditional formatting for yearly change
                    If yearChange = 0 Then
                        ws.Cells(summary, 10).Interior.ColorIndex = 0
                    ElseIf yearChange > 0 Then
                        ws.Cells(summary, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summary, 10).Interior.ColorIndex = 3
                    End If
                    
                ' formulate percentage change
                perChange = (yearChange / firstOpen)
                ws.Cells(summary, 11).Value = perChange
                ws.Cells(summary, 11).NumberFormat = "0.00%"
                ' Code from AskBCS Learning Assistant: Range(<selected column>).NumberFormat = "0.000%"
                
                ' reset values of variables (add 1 onto the summary count)
                summary = summary + 1
                totalStock = 0
                firstOpen = 0
                yearChange = 0
                perChange = 0
            
            ' second condition
            ElseIf ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
                
                ' record first open value
                firstOpen = ws.Cells(Row, 3).Value
                
                'set initial total stock for the ticker
                totalStock = ws.Cells(Row, 7).Value
            
            ' final condition (if the ticker is the same in the next cell)
            Else
                ' add to the total stock amount
                totalStock = totalStock + ws.Cells(Row, 7).Value
            
            End If
        
        Next Row
    
        ' Determine greatest increase percentage
        Dim greatPerInc As Double
        greatPerInc = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        ws.Range("Q2").Value = greatPerInc
        ws.Range("Q2").NumberFormat = "0.00%"
        ' Code from AskBCS Learning Assistant: Range(<selected column>).NumberFormat = "0.000%"
        
        ' find the matching ticker
        Dim tickerPerInc As Integer
        tickerPerInc = WorksheetFunction.Match(greatPerInc, ws.Range("K1:K" & lastRow), 0)
        ws.Range("P2").Value = ws.Range("I" & tickerPerInc)
        
        ' Determine greatest decrease percentage
        Dim greatPerDec As Double
        greatPerDec = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
        ws.Range("Q3").Value = greatPerDec
        ws.Range("Q3").NumberFormat = "0.00%"
        ' Code from AskBCS Learning Assistant: Range(<selected column>).NumberFormat = "0.000%"
        
        ' find the matching ticker
        Dim tickerPerDec As Integer
        tickerPerDec = WorksheetFunction.Match(greatPerDec, ws.Range("K1:K" & lastRow), 0)
        ws.Range("P3").Value = ws.Range("I" & tickerPerDec)
        
        ' find the greatest total stock
        Dim maxTotalStock As Double
        maxTotalStock = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
        ws.Range("Q4").Value = maxTotalStock
        
        ' find the matching ticker
        Dim tickerMaxStock As Integer
        tickerMaxStock = WorksheetFunction.Match(maxTotalStock, ws.Range("L1:L" & lastRow), 0)
        ws.Range("P4").Value = ws.Range("I" & tickerMaxStock)
        
        ' formatting for new columns
        ws.Range("I1:Q1").EntireColumn.AutoFit
    
    Next ws
    
End Sub
