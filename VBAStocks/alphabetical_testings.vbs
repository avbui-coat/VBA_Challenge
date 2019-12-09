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

' Assign columns for summary side
combined_ticker.Range("I1").Value = "Ticker"
combined_ticker.Range("J1").Value = "Yearly Raw Change"
combined_ticker.Range("K1").Value = "Yearly Percent Change"
combined_ticker.Range("L1").Value = "Total Volume"

End Sub

Sub loop_for_summary():

' summarize volume total'

Dim totalVol As Double
Dim ticker As String
Dim summaryrow as Integer
Dim i as Integer

Set combined_ticker = Worksheets("Combined_Ticker")

lastrow = combined_ticker.Cells(Rows.Count, "A").End(xlUp).Row - 1

summaryrow = 2
totalVol = 0

For i = 2 To lastrow
    
    if (combined_ticker.Cells(i + 1, 1).Value <> Combined_ticker.Cells(i, 1).Value) Then
        'to calculate volume total'
        combined_ticker.Cells(summaryrow, 12).Value = totalVol + combined_ticker.Cells(i,7).Value
        combined_ticker.Cells(summaryrow, 9).Value = combined_ticker.Cells(i, 1).Value
       
        'to calculate percent change'





        'to index the summary side'
        summaryrow = summaryrow + 1
        totalVol = 0

    else
        totalVol = totalVol + combined_ticker.Cells(i, 7).Value
    end if


next i


End Sub


Sub percent_change():




End Sub



Sub total_volume_per_ticker():

End Sub

Sub color_index_for_change():



End Sub