Attribute VB_Name = "Module1"
Sub WallStreetVBA()
    
    
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        
        Dim Ticker As String
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Dim table_row As Integer
        Dim column As Integer
        
        
        Total_Stock_Volume = 0
        table_row = 2

        
        'Find the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Create Summary Table Headings
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        
        'Set Initial Open Price
        Open_Price = Cells(2, 3).Value
         
         
        'Loop through all ticker types
        
        For i = 2 To LastRow
         
            'Check if next row contains same ticker type as current row. If the next row ticker type is different than the current row then...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'Grab Ticker name and Print in Summary Table
                Ticker = Cells(i, 1).Value
                ws.Range("I" & table_row).Value = Ticker
                
                'Set Close Price
                Close_Price = Cells(i, 6).Value
            
                'Calculate Yearly Change and Print in Summary Table
                Yearly_Change = Close_Price - Open_Price
                Cells(table_row, 10).Value = Yearly_Change
                
                'Calculate Percent Change and Print in Summary Table
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(table_row, 11).Value = Percent_Change
                    Cells(table_row, 11).NumberFormat = "0.00%"
                
                End If
                
                'Calculate Total Stock Volume and Print in Summary Table
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                Cells(table_row, 12).Value = Total_Stock_Volume
                
                'Add one to the summary table row
                table_row = table_row + 1
                
                'Reset Open Price
                Open_Price = Cells(i + 1, 3)
                
                'Reset Total Stock Volume
                Total_Stock_Volume = 0
            
            
            'If cells are the same ticker type...
            
            Else
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            End If
        
        Next i
        
        
        'Apply conditional formatting to Summary Table to highlight positive changes in green and negative changes in red
        
        ' Determine the Last Row in Summary Table
        
        SummaryTableLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'For Yearly Change, highlight postive changes in green and negative changes in red
            For j = 2 To SummaryTableLastRow
                
                If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                    Cells(j, 10).Interior.ColorIndex = 10
                
                ElseIf Cells(j, 10).Value < 0 Then
                    Cells(j, 10).Interior.ColorIndex = 3
                
                End If
            Next j
        
        
        ' Set Greatest % Increase, % Decrease, and Total Volume
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 16).Value = "Greatest % Decrease"
        Cells(4, 17).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        ' Look through each rows to find the greatest value and its associate ticker
        
        For k = 2 To SummaryTableLastRow
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & SummaryTableLastRow)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & SummaryTableLastRow)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & SummaryTableLastRow)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k
        
    Next ws
        
End Sub

