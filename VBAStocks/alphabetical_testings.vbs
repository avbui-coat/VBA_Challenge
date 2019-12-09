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

Sub insert_header():
 
Set combined_ticker = Worksheets("Combined_Ticker")
 
combined_ticker.Range("A1:G1") = Sheets(2).Range("A1:G1").Value

End Sub

Sub insert_columns():

Set combined_ticker = Worksheets("Combined_Ticker")

' Assign columns
combined_ticker.Range("I1").Value = "Ticker"
combined_ticker.Range("J1").Value = "Yearly Raw Change"
combined_ticker.Range("K1").Value = "Yearly Percent Change"
combined_ticker.Range("L1").Value = "Total Volume"

End Sub

Sub yearly_change():

lastrow = combined_ticker.Cells(Rows.Count, "A").End(xlUp).Row +1

for i = 2 to lastrow



next i








End Sub


Sub percent_change():
' Sheets.Add(Before:=Sheets(1)).Name = "Percent_Change"



End Sub



Sub total_volume_per_ticker():

End Sub