Attribute VB_Name = "Module1"
Sub StocksData()

For Each ws In Worksheets

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'Declarations
Dim Ticker As String
Dim Volume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim LastRow As Long
Dim RowValue As Integer
Dim StockOpen As Double

'Find the last non-blank cell in column A(1)
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

RowValue = 2
Volume = 0

For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value
    YearlyChange = ws.Cells(i, 6).Value - StockOpen
    If StockOpen = 0 Then
    PercentChange = 0
    Else:
    PercentChange = YearlyChange / StockOpen
    End If
    
    
ws.Range("I" & RowValue).Value = Ticker
ws.Range("L" & RowValue).Value = Volume

Volume = 0

ws.Range("J" & RowValue).Value = YearlyChange
ws.Range("K" & RowValue).Value = PercentChange
ws.Range("K" & RowValue).Style = "Percent"
ws.Range("K" & RowValue).NumberFormat = "0.00%"

RowValue = RowValue + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
StockOpen = ws.Cells(i, 3).Value
Volume = Volume + ws.Cells(i, 7).Value


Else: Volume = Volume + ws.Cells(i, 7).Value

End If

    Next i

For i = 2 To LastRow

If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4

ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
End If
    Next i
    
'Challenges
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

'Greatest%Increase Calculations
For a = 2 To LastRow


    If ws.Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    End If

    Next a

'Greatest%Decrease Calculations
For b = 2 To LastRow
    
    If ws.Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    End If
    
   Next b
   
'GreatestVolume Calculations
For c = 2 To LastRow
    
    If ws.Cells(c, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = GreatestVolume
        ws.Range("P4").Value = ws.Cells(c, 9).Value
    End If
  
    Next c
 
ws.Columns("A:Q").AutoFit
    
    
Next ws

End Sub

