Attribute VB_Name = "Module1"
Sub stock_homework()

' Naming Variables

Dim total As Double
Dim ticker As String
Dim percentChange As Double
Dim yearlyOpen As Double
Dim yearlyClose As Double
Dim yearlyChange As Double
Dim lastrow As Long


total = 0


'Naming Titles

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total"

Dim table As Integer
table = 2


'last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    
For i = 2 To lastrow

If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

yearlyOpen = Cells(i, 3).Value

End If

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

yearlyClose = Cells(i, 6).Value

ticker = Cells(i, 1).Value

total = total + Cells(i, 7).Value

yearlyChange = yearlyClose - yearlyOpen

percentChange = yearlyChange / yearlyOpen

Range("I" & table).Value = ticker

Range("L" & table).Value = total

Range("J" & table).Value = yearlyChange

Range("K" & table).Value = perentChange

Range("K" & table).NumberFormat = "0.00%"


table = table + 1
total = 0
yearlyChange = 0
yearlyOpen = Cells(i + 1, 3).Value

Else

total = total + Cells(i, 7).Value
yearlyChange = yearlyClose - yearlyOpen
percentChange = yearlyChange / yearlyOpen

End If

Next i


For i = 2 To table

If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 10
Else
Cells(i, 10).Interior.ColorIndex = 3
End If

Next i


End Sub

