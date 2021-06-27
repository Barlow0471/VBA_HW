VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Lop()

'Define Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"


'Declare Variables
Dim LastRow As Double
Dim ticker As String
Dim tickercounter As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim OpenPrice As Double
Dim PercentChange As Double
Dim YearlyOpen As Double
Dim TotalVolume As Double

ClosePrice = 0
UniqueCounter = 1
PercentChange = 0

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
ticker = ""

OpenPrice = Range("C2").Value

For i = 2 To LastRow

    ticker = Range("A" & i)
    TotalVolume = TotalVolume + Cells(i, 7).Value
       
    'If statement to determine each Unique Ticker and Yearly Change
    If Range("A" & i) <> Range("A" & i + 1) Then
    
        
        UniqueCounter = UniqueCounter + 1
        Range("I" & UniqueCounter).Value = Range("A" & i).Value
        Range("L" & UniqueCounter).Value = TotalVolume
        TotalVolume = 0
        
        ClosePrice = Range("F" & i).Value
        YearlyChange = ClosePrice - OpenPrice
        Range("J" & UniqueCounter) = YearlyChange
        
        
        'If statement to determine Percent Change
        If OpenPrice = 0 Then
            PercentChange = 0
        Else
            PercentChange = YearlyChange / OpenPrice
        End If
            Range("K" & UniqueCounter).Value = PercentChange
            
        'Conditional Formatting for Green/Red cells
        If Range("J" & UniqueCounter).Value >= 0 Then
            Range("J" & UniqueCounter).Interior.ColorIndex = 4
        Else
            Range("J" & UniqueCounter).Interior.ColorIndex = 3
        End If
                   
        
    OpenPrice = Range("C" & i + 1)
    

    End If

Next i

'New loop to find Greatest % Change, Least % Change, and Greatest Volume
LastRow = Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRow
    
    'If statement to find Greatest % Increase
    If Range("K" & i).Value > Range("P2").Value Then
        Range("P2").Value = Range("K" & i).Value
        Range("O2").Value = Range("I" & i).Value
    End If
    
    'If statement to find Greatest % Decrease
    If Range("K" & i).Value < Range("P3").Value Then
        Range("P3").Value = Range("K" & i).Value
        Range("O3").Value = Range("I" & i).Value
    End If
    
    'If statement to find Greatest Total Volume
    If Range("L" & i).Value > Range("P4").Value Then
        Range("P4").Value = Range("L" & i).Value
        Range("O4").Value = Range("I" & i).Value
    End If
    
Next i

'Convert necessary fields to percentages
Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"
Columns("K").NumberFormat = "0.00%"

'Autofit columns
For Each sht In ThisWorkbook.Worksheets
    sht.Cells.EntireColumn.AutoFit
  Next sht
End Sub

