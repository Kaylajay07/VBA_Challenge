Sub StockData()

    Dim ticker As String
    Dim tablerow As Integer
    Dim volume, openprice, closeprice As Double
    Dim maxinc, maxdec, maxvol As Double
      
    For Each ws In Worksheets
    ws.Activate
    
        tablerow = 2
        volume = 0
        openprice = Cells(2, "C")
        maxinc = 0
        maxdec = 0
        maxvol = 0
        
        Cells(1, "I") = "Ticker "
        Cells(1, "J") = "Yearly Change"
        Cells(1, "K") = "Percentage Change"
        Cells(1, "L") = "Stock Volume "
        Cells(2, "N") = "Greatest % Increase"
        Cells(3, "N") = "Greatest % Decrease"
        Cells(4, "N") = "Greatest Total Volume"
        Cells(1, "O") = "Ticker"
        Cells(1, "P") = "Value"
        
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
             volume = volume + Cells(i, "G")
             If Cells(i, "A") <> Cells(i + 1, "A") Then
                 Cells(tablerow, "I") = Cells(i, "A")
                 
                 closeprice = Cells(i, "F")
                 Cells(tablerow, "J") = closeprice - openprice
                 
                 If Cells(tablerow, "J") > 0 Then
                    Cells(tablerow, "J").Interior.ColorIndex = 4
                 Else
                    Cells(tablerow, "J").Interior.ColorIndex = 3
                 End If
                 
                 If openprice > 0 Then
                    Cells(tablerow, "K") = FormatPercent((closeprice - openprice) / openprice, 2)
                 Else
                    Cells(tablerow, "K") = 0
                 End If
                 
                 Cells(tablerow, "L") = volume
                 
                 If Cells(tablerow, "K") > maxinc Then
                    maxinc = Cells(tablerow, "K")
                    maxincticker = Cells(tablerow, "I")
                End If
                
                 If Cells(tablerow, "K") < maxdec Then
                    maxdec = Cells(tablerow, "K")
                    maxdecticker = Cells(tablerow, "I")
                End If
                
                If Cells(tablerow, "L") > maxvol Then
                    maxvol = Cells(tablerow, "L")
                    maxvolticker = Cells(tablerow, "I")
                
                End If
                 
                 tablerow = tablerow + 1
                 volume = 0
                 openprice = Cells(i + 1, "C")
             
            End If
            
            Cells(2, "O") = maxincticker
            Cells(2, "P") = FormatPercent(maxinc)
            
            Cells(3, "O") = maxdecticker
            Cells(3, "P") = FormatPercent(maxdec)
            
            Cells(4, "O") = maxvolticker
            Cells(4, "P") = maxvol
        Next i
    
    Next ws

End Sub

