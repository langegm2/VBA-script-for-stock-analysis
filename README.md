# VBA-script-for-stock-analysis
module two challenge
vba challenge code
Sub stock_analysis()
Dim total As Double
Dim rowindex As Long
Dim change As Double
Dim columnindex As Integer
Dim start As Long
Dim rowcount As Long
Dim percentchange As Double
Dim days As Integer
Dim dailychange As Single
Dim averagechange As Double
Dim ws As Worksheet

For Each ws In Worksheets
    columnindex = 0
    total = 0
    change = 0
    start = 2
    dailychange = 0
    
    ws.Range("i1").Value = "ticker"
    ws.Range("j1").Value = "yearly change"
    ws.Range("k1").Value = "percent change"
    ws.Range("L1").Value = "Total stock volume"
    ws.Range("p1").Value = "ticker"
    ws.Range("Q1").Value = "value"
    ws.Range("o2").Value = "Greatest % increase"
    ws.Range("o3").Value = "greatest % decrease"
    ws.Range("o4").Value = "Greatest total volume"
    
    
    rowcount = ws.Cells(Rows.Count, "a").End(xlUp).Row
    
    For rowindex = 2 To rowcount
        If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
    
        total = total + ws.Cells(rowindex, 7).Value
    
        If total = 0 Then
        ws.Range("i" & 2 + columnindex).Value = Cells(rowindex, 1).Value
        ws.Range("J" & 2 + columnindex).Value = 0
        ws.Range("k" & 2 + columnindex).Value = "%" & 0
        ws.Range("L" & 2 + columnindex).Value = 0
        Else
            If ws.Cells(start, 3) = 0 Then
                For find_value = start To rowindex
                    If ws.Cells(find_value, 3).Value <> 0 Then
                    start = find_value
                    Exit For
                End If
                    
           
            Next find_value
                End If
                change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
                percentchange = change / ws.Cells(start, 3)
                
                start = rowindex + 1
                ws.Range("i" & 2 + columnindex) = ws.Cells(rowindex, 1).Value
                ws.Range("j" & 2 + columnindex) = change
                ws.Range("j" & 2 + columnindex).NumberFormat = "0.00"
                ws.Range("k" & 2 + columnindex).Value = percentchange
                ws.Range("K" & 2 + columnindex).NumberFormat = "0.00"
                ws.Range("L" & 2 + columnindex).Value = total
                
                Select Case change
                    Case Is > 0
                    ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 4
                    Case Is < 0
                    ws.Range("j" & 2 + columnindex).Interior.ColorIndex = 3
                    Case Else
                    ws.Range("j" & 2 + columnindex).Interior.ColorIndex = 0
                End Select
                
                    
                    
                End If
                total = 0
                change = 0
                columnindex = columnindex + 1
                days = 0
                dailychange = 0
            Else
                total = total + ws.Cells(rowindex, 7).Value
                
            End If
                 
    
    
    Next rowindex
    
        ws.Range("q2") = "%" & worksheetfuntion.Max(ws.Range("k2:k" & rowcount)) * 100
        ws.Range("q3") = "%" & worksheetfuntion.Min(ws.Range("k2:k" & rowcount)) * 100
        ws.Range("q4") = worksheetfuntion.Max(ws.Range("L2:L" & rowcount))
        
        increase_number = worksheetfuntion.Match(worksheetfuntion.Max(ws.Range("k2:k" & rowcount)), ws.Range("K2:k" & rowcount), 0)
        decrease_number = worksheetfuntion.Match(worksheetfuntion.Min(ws.Range("k2:k" & rowcount)), ws.Range("K2:k" & rowcount), 0)
        volume_number = worksheetfuntion.Match(worksheetfuntion.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)
        
        ws.Range("p2") = ws.Cells(increase_number + 1, 9)
        ws.Range("p3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("p4") = ws.Cells(volume_number + 1, 9)
        
        
    
    
Next ws



End Sub

