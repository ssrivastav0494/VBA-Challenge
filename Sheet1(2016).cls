VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockYearLoop():
Dim Ticker As String
Dim TotalVolume As Double
Dim Change As Double
Dim PercentChange As Double
Dim SummTablRow As Integer
Dim Start As Integer
Dim LastRow As Double

LastRow = Range("A" & Rows.Count).End(xlUp).Row
TotalVolume = 0
SummTablRow = 2
Start = 2
    
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "% Change"
Range("L1") = "Total Volume"

For i = 2 To LastRow
    If Cells(i + 1, 1) <> Cells(i, 1) Then
        If Cells(Start, 3) = 0 Then
        For Value = Start To i
            If Cells(Value, 3) <> 0 Then
                Start = Value
                Exit For
            End If
        Next Value
        End If
    Change = Cells(i, 6) - Cells(Start, 3)
    PercentChange = Round((Change / Cells(Start, 3)) * 100, 2)
    Ticker = Cells(i, 1)
    TotalVolume = TotalVolume + Cells(i, 7)
    Range("I" & SummTablRow) = Ticker
    Range("J" & SummTablRow) = Change
    Range("K" & SummTablRow) = Str(PercentChange) + "%"
    Range("L" & SummTablRow) = TotalVolume
    If Change >= 0 Then
    Range("J" & SummTablRow).Interior.ColorIndex = 4
    Else
    Range("J" & SummTablRow).Interior.ColorIndex = 3
    End If
    SummTablRow = SummTablRow + 1
    Start = i + 1
    TotalVolume = 0
    Change = 0
    Else
    OpeningPrice = Cells(i, 3)
    TotalVolume = TotalVolume + Cells(i, 7)
    End If

Next i
End Sub
