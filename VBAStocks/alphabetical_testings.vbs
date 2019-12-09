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



Sub color_index_for_change():



End Sub

Sub loop_for_summary_each_sheet():

' summarize volume total'
'totalVol will me long so I had to use CLng() to work around it - VBA gave errors: Overflow
Dim totalVol As Long
Dim ticker As String
Dim summaryrow as Integer
Dim lastrow as Long

summaryrow = 2
totalVol = 0

For each ws in Worksheets

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

    For i = 2 To i = Clng(lastrow)
    
        if (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        'to calculate volume total'
            ws.Cells(summaryrow, 12).Value = CLng(totalVol) + ws.Cells(i,7).Value
            ws.Cells(summaryrow, 9).Value = ws.Cells(i, 1).Value
       
        'to calculate percent change'
           
        'to index the summary side'
            summaryrow = summaryrow + 1
            totalVol = 0

        else
            totalVol = CLng(totalVol) + ws.Cells(i, 7).Value
        end if


    next i
next ws

End Sub


Sub loop_for_combined_sheet():

Set combined_ticker = Worksheets("Combined_Ticker")
' summarize volume total'
'totalVol will me long so I had to use CLng() to work around it - VBA gave errors: Overflow
Dim totalVol As Long
Dim ticker As String
Dim summaryrow as Integer
Dim lastrow as Long

summaryrow = 2
totalVol = 0

lastrow = combined_ticker.Cells(Rows.Count, "A").End(xlUp).Row - 1

    For i = 2 To i = Clng(lastrow)
    
        if (combined_ticker.Cells(i + 1, 1).Value <> combined_ticker.Cells(i, 1).Value) Then
        'to calculate volume total'
            combined_ticker.Cells(summaryrow, 12).Value = CLng(totalVol) + combined_ticker.Cells(i,7).Value
            combined_ticker.Cells(summaryrow, 9).Value = combined_ticker.Cells(i, 1).Value
       
        'to calculate percent change'
           
        'to index the summary side'
            summaryrow = summaryrow + 1
            totalVol = 0

        else
            totalVol = CLng(totalVol) + combined_ticker.Cells(i, 7).Value
        end if


    next i


End Sub
