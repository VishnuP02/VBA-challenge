Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        lastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long


        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        
       OpenPrice = Cells(2, 3).Value
       
        
        For i = 2 To lastRow
         
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
              
                Ticker = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker
               
                ClosePrice = Cells(i, 6).Value
               
                YearlyChange = ClosePrice - OpenPrice
                Cells(Row, 10).Value = YearlyChange
                'Cells(Row, 10).NumberFormat = "0.000000000"
              
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / OpenPrice
                    Cells(Row, 11).Value = PercentChange
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
               
                Row = Row + 1
                
                OpenPrice = Cells(i + 1, 3)
               
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
       
        YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
       
        For j = 2 To YCLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

        
        For Z = 2 To YCLastRow
            If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(Z, 9).Value
                Cells(2, 17).Value = Cells(Z, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(Z, 9).Value
                Cells(3, 17).Value = Cells(Z, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(Z, 9).Value
                Cells(4, 17).Value = Cells(Z, 12).Value
            End If
        Next Z
        
    Next WS
        
End Sub
