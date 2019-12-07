sub take_ticker():
Sheets.Add(Before:=Sheets(1)).Name = "Combined_Ticker"

Dim ticker as String

Set combined_ticker = Worksheets("Combined_Ticker")
For Each ws in Worksheets
    lastrowTicker = ws.Cells(Rows.Count, "A").End(xlUp).Row -1

    lastrow = combined_ticker.Cells(Rows.Count, "A").End(xlUp).Row +1
    
    combined_ticker.Range("A" & lastrow & ":G" & ((lastrowTicker - 1) + lastrow)).Value = ws.Range("A2:G" & (lastrowTicker + 1)).Value

Next ws
End sub


Sub yearly_change():

End Sub


Sub percent_change():



End Sub



Sub total_volume_per_ticker():

End Sub