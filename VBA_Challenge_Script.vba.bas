Attribute VB_Name = "Module1"
Sub stock()
    
    
    Dim ws As Worksheet
    Dim ticker As String
    Dim volume As Double
    Dim summary_table_row As Integer
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim lastrow As Long
    Dim percent_change As Double
    
    
 For Each ws In Worksheets
     
    opening_price = ws.Cells(2, 3).Value
    volume = 0
    summary_table_row = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
                ws.Range("L" & summary_table_row).Value = ticker
                
                closing_price = ws.Cells(i, 6).Value
                
                yearly_change = closing_price - opening_price
                ws.Range("M" & summary_table_row).Value = yearly_change
                ws.Range("M" & summary_table_row).NumberFormat = "$#.00"
                
                     If ws.Range("M" & summary_table_row).Value > 0 Then
                        ws.Range("M" & summary_table_row).Interior.Color = vbGreen
                    ElseIf ws.Range("M" & summary_table_row).Value < 0 Then
                        ws.Range("M" & summary_table_row).Interior.Color = vbRed
                     
                    Else
                       ws.Range("M" & summary_table_row).Interior.Color = vbWhite
                    End If
            
            
                    If (opening_price = 0 And closing_price = 0) Then
                        percent_change = 0
                        
                        ElseIf (opening_price = 0 And closing_price <> 0) Then
                            percent_change = 1
                            
                Else
                
                percent_change = yearly_change / opening_price
                ws.Range("N" & summary_table_row).Value = percent_change
                ws.Range("N" & summary_table_row).NumberFormat = "0.00%"
                End If
                
                
                volume = volume + ws.Cells(i, 7).Value
                ws.Range("O" & summary_table_row).Value = volume
                
                volume = 0
                opening_price = ws.Cells(i + 1, 3).Value
                summary_table_row = summary_table_row + 1
                Else
                volume = volume + ws.Cells(i, 7).Value
                
             End If
                
            Next i
   
    
ws.Cells(1, 12).Value = "Ticker"
ws.Cells(1, 13).Value = "Yearly Change"
ws.Cells(1, 14).Value = "Percent Change"
ws.Cells(1, 15).Value = "Total Stock Volume"

Next ws

            
    
End Sub


