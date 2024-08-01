Sub ByTicker()

'------------------------------------------------------
'  notes
'------------------------------------------------------

' _____ note #1 _____
' I began my data row counter at row 3 so I can record
' the previous open thru close numbers without
' snagging on the headers in row 1.


' _____ note #2 _____
' The column header given was "quarterly change" but
' each dataset covers a full year. I'm assuming this
' was an oversight - I changed the header to "yearly"


'------------------------------------------------------
'  variables
'------------------------------------------------------

' worksheet variable
    Dim ws As Worksheet

' record counters for original and ticker datasets
    Dim DataRow As Long
    Dim TickerRow As Long

' last row for original dataset
    Dim DataLR As Long

' datum trackers - columns I thru L
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim StockVol As LongLong

' datum trackers - columns N thru P
    Dim MaxPct As Double
    Dim MinPct As Double
    Dim MaxVol As LongLong
    
    Dim MaxTkr As String
    Dim MinTkr As String
    Dim VolTkr As String


'------------------------------------------------------
'  workflow
'------------------------------------------------------

' for loop - to cycle through sheets
For Each ws In Worksheets

    ws.Activate
  
  ' spreadhseet prep
  ' ----------------
    ' set headers in columns I thru L
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"                              'see note #2 above
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    ' set grid in columns N thru P
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    
  ' variable prep
  ' -------------
    ' reset record trackers
    DataRow = 3                                     'see note#1 above
    TickerRow = 2
    
    ' reset last row finder
    DataLR = Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    ' reset datum trackers: columns I to L
    OpenPrice = Range("C2")
    StockVol = Range("G2")
    
    ' reset datum trackers: columns N to P
    MaxPct = 0
    MinPct = 0
    MaxVol = 0
    
  ' workflow
  ' --------
    ' for loop - to cycle through given dataset
    For DataRow = 3 To DataLR
    
        ' determine whether the ticker changed in the current row
        If Cells(DataRow, 1) <> Cells(DataRow - 1, 1) Then
            
            ' record ticker in column I (found in previous data row)
            Cells(TickerRow, 9) = Cells(DataRow - 1, 1)
            
            ' set close price variable (found in previous data row)
            ClosePrice = Cells(DataRow - 1, 6)
            
            ' determine and record the change in price by dollar & percent in columns J & K
            Cells(TickerRow, 10) = Format(ClosePrice - OpenPrice, "0.00")
            Cells(TickerRow, 11) = Format(Cells(TickerRow, 10) / ClosePrice, "0.00%")
            
            ' record total stock volume in column L (found in previous row)
            Cells(TickerRow, 12) = Format(StockVol, "#,###")
            
            ' reset open price & stock volume variables (found in the current row)
            OpenPrice = Cells(DataRow, 3)
            StockVol = Cells(DataRow, 7)
            
            ' next ticker row
            TickerRow = TickerRow + 1
            
        Else
        
            StockVol = StockVol + Cells(DataRow, 7)
            
        End If
        
    Next DataRow
        
    ' for loop - new data set - determine max/mins and add conditional formatting
    For i = 2 To TickerRow
    
        ' add conditional formatting to column J
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        End If
        
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        ' determine/overwrite whether this row is the highest percent change
        If Cells(i, 11) > MaxPct Then
            MaxPct = Cells(i, 11)
            MaxTkr = Cells(i, 9)
        End If
        
        ' determine/overwrite whether this row is the lowest percent change
        If Cells(i, 11) < MinPct Then
            MinPct = Cells(i, 11)
            MinTkr = Cells(i, 9)
        End If
        
        ' determine/overwrite whether this row is the highest stock volume
        If Cells(i, 12) > MaxVol Then
            MaxVol = Cells(i, 12)
            VolTkr = Cells(i, 9)
        End If
    
    Next i
    
    ' record maximum % change, minimum % change, and max volume change to columns N to P
    Range("O2") = MaxTkr
    Range("P2") = Format(MaxPct, "0.00%")
    Range("O3") = MinTkr
    Range("P3") = Format(MinPct, "0.00%")
    Range("O4") = VolTkr
    Range("P4") = Format(MaxVol, "#,###")
    
    ' column fit
    Columns("I:P").AutoFit
    

Next ws


End Sub

