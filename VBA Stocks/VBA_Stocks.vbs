'ticker symbol
Sub ticker()

For Each ws in Worksheets

Dim Ticker as String
Dim Summary_Table_Row as Integer
Summary_Table_Row = 2

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1").value= "Ticker"

    For i = 2 to Lastrow
        if ws.cells(i+1,1).value <> ws.cells(i,1).value then 
        Ticker = ws.cells(i,1).value 
        ws.range("I" & Summary_Table_Row).value = Ticker
        Summary_Table_Row = Summary_Table_Row +1
        end if
    next i
next ws 

end sub 




Sub total_stock_volume()

For Each ws In Worksheets

Dim ticker As String
Dim vol_total As Long
Dim volume_summary As Long
vol_total = 0
volume_summary = 2

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("L1").Value = "Total Stock Volume"

  For i = 2 To Lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.cells(i,1).value 
        vol_total = vol_total + ws.Cells(i, 7).Value
        ws.Range("L" & volume_summary).Value = vol_total
        volume_summary = volume_summary + 1
        vol_total = 0
        Else 
        vol_total = vol_total + ws.Cells(i, 7).Value
        End If
    Next i
Next ws

End Sub




Sub yearly_change()
For Each ws In Worksheets
Dim yearlychange As Double
Dim change_row As Long
Dim percentchange_row As Double
Dim open_stock As Double
Dim Close_stock As Double
Dim percentchange As Double
change_row = 2
percentchange_row = 2

'ticker rows
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"

open_stock = ws.Cells(2, 3).Value

  For i = 2 To Lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Close_stock = ws.Cells(i, 6).Value
    yearlychange = Close_stock - open_stock
     If Close_stock = 0 Then
        percentchange = 0
        Else: percentchange = Round((yearlychange / Close_stock), 2)
        End If
    ws.Range("J" & change_row).Value = yearlychange
        change_row = change_row + 1
    ws.Range("K" & percentchange_row).Value = "%" & percentchange
    percentchange_row = percentchange_row + 1
    open_stock = ws.Cells(i + 1, 3).Value
    End If
Next i
Next ws

End Sub


