Sub insert_columns_to_eachSheet():

For each ws in Worksheets
' Assign columns for summary side
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Raw Change"
    ws.Range("K1").Value = "Yearly Percent Change"
    ws.Range("L1").Value = "Total Volume"
Next ws

End Sub


Sub Total_Volume_Summary_Loop():

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
       
        'to index the summary side'
            summaryrow = summaryrow + 1
            totalVol = 0

        else
            totalVol = CLng(totalVol) + ws.Cells(i, 7).Value
        end if


    next i
next ws

End Sub

Sub Yearly_Change_Summary_Loop():

Dim summaryrow as Integer
Dim raw_change as Long

raw_change = 0
summaryrow = 2

For each ws in Worksheets

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1


    For i = 2 To i = Clng(lastrow)
        

        if (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            lastdate = ws.Cells(Row.Count,"B").End(xlUp.Row)
     
            raw_change = ws.Cells(i, 3).Value - ws.Cells(i + lastdate, 5).Value
           
            ws.Cells(summaryrow, 12).Value = raw_change
            ws.Cells(summaryrow, 9).Value = ws.Cells(i, 1).Value
       
        'to index the summary side'
            summaryrow = summaryrow + 1
            raw_change = 0

        end if


    next i
next ws


summaryrow = 2

End Sub