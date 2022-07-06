Attribute VB_Name = "Module1"
Sub StockMarket():

For Each ws In Worksheets
  'Define Variable Types and Variable Values
    Dim Ticker As String
    
    Dim YearChange As Double
        YearChange = 0
    Dim YearFinal As Double
        YearFinal = 0
    Dim YearInit As Double
        
    Dim StockVol As Double
        StockVol = 0
        
    Dim percentChange As Double
        percentChange = 0
        
    Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    Dim SumTableTitles(3) As String
        SumTableTitles(0) = "Ticker"
        SumTableTitles(1) = "Yearly Change"
        SumTableTitles(2) = "Percent Change"
        SumTableTitles(3) = "Total Stock Volume"

'Summary table values
    Dim SummTable As Long
        SummTable = 2
    YearInit = ws.Cells(2, 3).Value

' "For loop" that looks through ticker values
    For i = 2 To LastRow
'Enter titles for the summary table
        ws.Range("J1:M1") = SumTableTitles()
        
'Start looping for values, based on ticker value
        YearFinal = ws.Cells(i, 6).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           
' Define Ticker Value and define StockVol values, as well as
' closing value at end of year
            Ticker = ws.Cells(i, 1).Value
            StockVol = StockVol + ws.Cells(i, 7).Value
            YearChange = YearFinal - YearInit
'Deals with divide by 0 instances
            If YearInit <> 0 Then
                percentChange = ((YearChange / YearInit) * 100)
                ws.Range("L" & SummTable).Value = percentChange & "%"
            Else
                ws.Range("L" & SummTable).Value = "Error, check this value manually"

            End If
' Display in respective summary table the ticker value, StockVol, and YearChange
                ws.Range("J" & SummTable).Value = Ticker
                ws.Range("M" & SummTable).Value = StockVol
                ws.Range("K" & SummTable).Value = YearChange
'If a positive change, change color to green, otherwise change color to red
           If YearChange > 0 Then
                    ws.Range("K" & SummTable).Interior.ColorIndex = 4
                    Else
                    ws.Range("K" & SummTable).Interior.ColorIndex = 3
            End If
                
' Go to next row for the summary table
                SummTable = SummTable + 1
            
' Set StockVol and YearChange back to 0 for next line
                StockVol = 0
                YearFinal = 0
                YearInit = ws.Cells(i + 1, 3).Value
            Else
' Continue adding values to the StockVol
            StockVol = StockVol + ws.Cells(i, 7).Value
    
    End If
    Next i
    Next ws
End Sub
