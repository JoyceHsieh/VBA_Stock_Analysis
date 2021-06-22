Sub Ticker()

    For Each ws In Worksheets
    
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Chabge"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    
    Dim s As String
    Dim num As Integer
    Dim Op As Double
    Dim Clos As Double
    
    s = " "
    num = 1
    Op = 0
    Clos = 0



    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        s = ws.Cells(i, 1).Value
        num = num + 1
        ws.Cells(num, 10).Value = s
        Op = ws.Cells(i, 3).Value
                

        
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value And Op <> 0 Then
        Clos = ws.Cells(i, 6).Value
        x = Clos - Op
        ws.Cells(num, 11).Value = x
        y = (Clos - Op) / Op
        ws.Cells(num, 12).Value = y
        ws.Cells(num, 12).Style = "Percent"

        Clos = 0
        Op = 0
        
        End If
        
        If ws.Cells(num, 11).Value > 0 Then
        ws.Cells(num, 11).Interior.ColorIndex = 4
        
        Else
        ws.Cells(num, 11).Interior.ColorIndex = 3
        
        End If
        
    
                
    Next i
    
    Next ws
End Sub


Sub tsv()

    For Each ws In Worksheets

    Dim s As String
    Dim num As Integer
    Dim vol As Long

    s = " "
    num = 1
    vol = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


      For i = 2 To LastRow
    
        If s <> ws.Cells(i, 1).Value Then
        s = ws.Cells(i, 1).Value
        num = num + 1
        ws.Cells(num, 10).Value = s
        
        vol = ws.Cells(i, 7).Value
        ws.Cells(num, 13).Value = vol
        
        ElseIf s = ws.Cells(i, 1).Value Then
        ws.Cells(num, 13).Value = ws.Cells(i, 7).Value + ws.Cells(num, 13).Value


        End If
        
        Next i

        
    Next ws
    
End Sub


Sub Bonus()

  For Each ws In Worksheets
  
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"

        LastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
        
        Dim best_stock As String
        Dim best_value As Double
        best_value = ws.Cells(2, 12).Value
        
        Dim worst_stock As String
        Dim worst_value As Double
        worst_value = ws.Cells(2, 12).Value
        
        
        Dim most_vol_stock As String
        Dim most_vol_value As Double
        most_vol_value = ws.Cells(2, 13).Value
        
        
        For j = 2 To LastRow
            If ws.Cells(j, 12).Value > best_value Then
                best_value = ws.Cells(j, 12).Value
                best_stock = ws.Cells(j, 10).Value
            End If

            
            If ws.Cells(j, 12).Value < worst_value Then
                worst_value = ws.Cells(j, 12).Value
                worst_stock = ws.Cells(j, 10).Value
            End If

        
            If ws.Cells(j, 13).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 13).Value
                most_vol_stock = ws.Cells(j, 10).Value
            End If

        Next j
        
        
        ws.Cells(2, 17).Value = best_stock
        ws.Cells(2, 18).Value = best_value
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = worst_stock
        ws.Cells(3, 18).Value = worst_value
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = most_vol_stock
        ws.Cells(4, 18).Value = most_vol_value
        
        ws.Columns("J:M").EntireColumn.AutoFit
        ws.Columns("P:R").EntireColumn.AutoFit
        
        Next ws

End Sub




