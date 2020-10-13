Attribute VB_Name = "Module1"
Sub YearChangePractice():

Dim open_price As Double

Dim close_price As Double

Dim year_change As Double

Dim percent_change As Double

Dim lastrow As Long

Dim str As Integer

Dim ticker As String

Dim vol As Double

For Each ws In Worksheets


    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    str = 2
    open_price = 0
    close_price = 0
    year_change = 0
    percent_change = 0
    vol = 0
    
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    

    For i = 2 To lastrow
    
        If open_price = 0 Then
          open_price = ws.Cells(i, 3).Value
       End If
    
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
           ticker = ws.Cells(i, 1).Value
           vol = vol + ws.Cells(i, 7).Value
           close_price = ws.Cells(i, 6).Value
           year_change = close_price - open_price
           
            If open_price = 0 Then
              percent_change = 0
            Else
               percent_change = year_change / open_price
               ws.Range("K" & str).Value = percent_change
               ws.Range("K" & str).Value = Format(percent_change, "Percent")
             End If
           
           ws.Range("I" & str).Value = ticker
           ws.Range("J" & str).Value = year_change
           
                If year_change > 0 Then
                   ws.Range("J" & str).Interior.ColorIndex = 4
                ElseIf year_change < 0 Then
                   ws.Range("J" & str).Interior.ColorIndex = 3
                Else
                   ws.Range("J" & str).Interior.ColorIndex = 6
                End If
               
           ws.Range("L" & str).Value = vol
           
           open_price = 0
           close_price = 0
           vol = 0
           
           str = str + 1
        
        Else
        
           vol = vol + ws.Cells(i, 7).Value
           
        End If
        

    Next i
    
    Next ws

End Sub
