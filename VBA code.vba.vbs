Sub stock_data():
'Set variable
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Integer
Dim start As Long
Dim row_count As Long
Dim persent_change As Double
Dim days As Integer
Dim daily_change As Double
Dim average_change As Double
Dim last_ticker As String


'Set title
Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Set inital Values
j = 0
total = 0
change = 0
start = 2
last_ticker = Cells(start, 1).Value

greatest_increase = -1E+30 ' Set to a very low number
greatest_decrease = 1E+30 ' Set to a very high number
greatest_volume = -1E+30 ' Set to a very low number

'get the last row data
row_count = Cells(Rows.Count, "A").End(xlUp).Row


'Set for loop to go through all rows
For i = 2 To row_count

'if the ticker changes we want to print all data and then move on to all data
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'keep track of total volume
total = total + Cells(i, 7).Value

'Calulating the quarterlly change and persent change 2 equations
change = Cells(i, 6).Value - Cells(start, 3).Value ' Closing - Opening
If Cells(start, 3).Value <> 0 Then
    percent_change = (change / Cells(start, 3).Value) * 100
Else
    percent_change = 0
End If

 ' Update greatest values
    If percent_change > greatest_increase Then
        greatest_increase = percent_change
        ticker_greatest_increase = last_ticker
    End If

    If percent_change < greatest_decrease Then
        greatest_decrease = percent_change
        ticker_greatest_decrease = last_ticker
    End If

    If total > greatest_volume Then
        greatest_volume = total
        ticker_greatest_volume = last_ticker
    End If

'change start to next ticker

'Print our results
Range("I" & 2 + j).Value = last_ticker ' Ticker symbol
Range("J" & 2 + j).Value = change ' Quarterly Change
Range("K" & 2 + j).Value = percent_change / 100 ' Percent Change (Convert back to decimal format)
Range("L" & 2 + j).Value = total ' Total Stock Volume

' Format Percent Change with percentage sign
Range("K" & 2 + j).NumberFormat = "0.00%"

'add color to quarterly change column
If change > 0 Then
    Range("J" & 2 + j).Interior.Color = RGB(0, 255, 0) ' Green
ElseIf change < 0 Then
    Range("J" & 2 + j).Interior.Color = RGB(255, 0, 0) ' Red
Else
    Range("J" & 2 + j).Interior.ColorIndex = xlNone  'No color for zero change
End If

'reset variables for new ticker
total = 0
change = 0
j = j + 1
days = 0
start = i + 1 ' Move start to the next row for the new ticker
last_ticker = Cells(start, 1).Value ' Update last_ticker for the new ticker

        Else
            ' Accumulate total volume for the current ticker
            total = total + Cells(i, 7).Value
            
'end the for loop and go to the next i

    End If
Next i

'Find the min and max
    
    Range("P2").Value = ticker_greatest_increase
    Range("Q2").Value = greatest_increase / 100

    Range("P3").Value = ticker_greatest_decrease
    Range("Q3").Value = greatest_decrease / 100

    Range("P4").Value = ticker_greatest_volume
    Range("Q4").Value = greatest_volume

    ' Format results
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "#,##0"
    
End Sub
