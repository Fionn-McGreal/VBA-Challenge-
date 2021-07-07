
Sub VBAloops()
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 8).Value = "Ticker"
Cells(1, 9).Value = "Yearly Change"
Cells(1, 10).Value = "Percent Change"
Cells(1, 11).Value = "Total Stock Volume"


Start = 2

Summrow = 2

Totalvolume = 0

For i = 2 To lastrow

OpenValue = Cells(Start, 3).Value


Totalvolume = Totalvolume + Cells(i, 7).Value

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ClosedValue = Cells(i, 6).Value

Start = i + 1

YearlyCHange = ClosedValue - OpenValue

Cells(Summrow, 9) = YearlyCHange

If Cells(Summrow, 9).Value > 0 Then

Cells(Summrow, 9).Interior.ColorIndex = 4

ElseIf Cells(Summrow, 9).Value <= 0 Then

Cells(Summrow, 9).Interior.ColorIndex = 3

End If



Cells(Summrow, 8) = Cells(i, 1).Value

Cells(Summrow, 11) = Totalvolume

Totalvolume = 0



If OpenValue = 0 Then

Cells(Summrow, 10) = 0


Else

PercentChange = YearlyCHange / OpenValue


Cells(Summrow, 10) = PercentChange

Cells(Summrow, 10).NumberFormat = "0.00%"

End If


Summrow = Summrow + 1


End If

Next i

End Sub
