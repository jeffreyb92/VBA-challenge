Attribute VB_Name = "Module1"
Sub StockScript():

    For Each ws In Worksheets
        'Defining and assigning our variables
        Dim Opening As Double
        Opening = 0
    
        Dim Closing As Double
        Closing = 0
    
        Dim Difference As Double
        Difference = Closing - Opening

        Dim totalVolume As Double
    
        totalVolume = 0
    
        Dim counter As Integer
        counter = 2
    
        Dim Ticker As String
   
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Initiating our For loop to do the calculations
        For i = 2 To LastRow
        
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And ws.Cells(i, 3).Value > 0 Then
                Opening = ws.Cells(i, 3).Value
                Ticker = ws.Cells(i, 1).Value
            
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value And ws.Cells(i, 7).Value > 0 Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And ws.Cells(i, 7).Value > 0 Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                Closing = ws.Cells(i, 6).Value
            
                'Inserting our titles in to the table
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"
            
                'Inserting our Values if this is the last cell for this stock into the table
                ws.Range("I" & counter).Value = Ticker
                ws.Range("J" & counter).Value = (Closing - Opening) / Opening
                If ws.Range("J" & counter).Value > 0 Then
                    ws.Range("J" & counter).Interior.Color = vbGreen
                Else
                    ws.Range("J" & counter).Interior.Color = vbRed
                End If
                ws.Range("K" & counter).Value = Format((Closing - Opening) / Closing, "0.00%")
                ws.Range("L" & counter).Value = totalVolume
            
                counter = counter + 1
                totalVolume = 0
            End If
    
        Next i
        'Inserting ourt titles into the table
        ws.Range("P1").Value = "Ticker"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Value"
    
        'for loop to find max positive change
        Dim maxchange As Double
        maxchange = 0
        For j = 2 To LastRow
            If ws.Range("K" & j).Value > maxchange Then
                maxchange = ws.Cells(j, 11).Value
                Ticker = ws.Cells(j, 9).Value
            End If
        Next j

        ws.Range("P2").Value = Ticker
    
        ws.Range("Q2").Value = Format(maxchange, "0.00%")
 
        'for loop to find max negative value
        Dim maxneg As Double
        maxneg = 0
        For k = 2 To LastRow
        
            If ws.Range("K" & k).Value < maxneg Then
                maxneg = ws.Range("K" & k).Value
                Ticker = ws.Cells(k, 9).Value
            End If
        Next k

        ws.Range("P3").Value = Ticker
    
        ws.Range("Q3").Value = Format(maxneg, "0.00%")
        
        'for loop to find the max total volume
        Dim maxtotal As Double
        maxtotal = 0
        For q = 2 To LastRow

            If ws.Range("L" & q).Value > maxtotal Then
                maxtotal = ws.Cells(q, 12).Value
                Ticker = ws.Cells(q, 9).Value
            End If
        Next q

        ws.Range("P4").Value = Ticker
    
        ws.Range("Q4").Value = maxtotal
    'Move to the next workbook
    Next ws
    
End Sub
