Attribute VB_Name = "Module6"
Sub wallstreet()
Dim ticker_row As Integer
ticker_row = 2

Dim LRow As Long
LRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim openc As Variant
Dim closec As Variant

Dim cls As Variant
Dim op As Variant

Dim perc As Double
Dim percy As Double

Dim ticker As String

Dim total_vol As Double
total_vol = 0

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"

'loop to find Ticker
For h = 2 To LRow
    
    If Cells(h + 1, 1).Value <> Cells(h, 1).Value Then
    ticker = Cells(h, 1).Value
    Cells(ticker_row, 9).Value = ticker
    ticker_row = ticker_row + 1
    End If

Next h

ticker_row = 2

'finding values for open then close
For i = 1 To LRow
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    openc = Cells(i + 1, 3).Value
    Cells(ticker_row, 15).Value = openc
    ticker_row = ticker_row + 1
    End If

Next i

ticker_row = 2

For j = 2 To LRow
   
    If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
    closec = Cells(j, 6).Value
    Cells(ticker_row, 16).Value = closec
    ticker_row = ticker_row + 1
    End If

Next j

ticker_row = 2

'calculating the difference for open and close values
For y = 1 To LRow
    
    If Cells(y + 1, 9).Value <> Cells(y, 9).Value Then
    op = Cells(y + 1, 15).Value
    cls = Cells(y + 1, 16).Value
    Cells(ticker_row, 10).Value = cls - op
    ticker_row = ticker_row + 1
    End If

Next y

range("K2", range("K2").End(xlDown)).NumberFormat = "0.00%"
range("J2", range("J2").End(xlDown)).NumberFormat = "$#,##0.00"

ticker_row = 2

'calculating percentage change
For p = 1 To LRow
    
    If Cells(p + 1, 9).Value <> Cells(p, 9).Value Then
    perc = Cells(p + 1, 10)
    percy = Cells(p + 1, 15)
        If (perc = 0) Or (percy = 0) Then
        Cells(ticker_row, 11).Value = 0
        Else
        Cells(ticker_row, 11).Value = perc / percy
        End If
    ticker_row = ticker_row + 1
    End If

Next p

ticker_row = 2

'cleanup since I didn't get it to work the efficient/correct way
range("O2", range("O2").End(xlDown)).Clear
range("P2", range("P2").End(xlDown)).Clear

'solving for total stock volume
For k = 2 To LRow
    
    If Cells(k + 1, 1).Value = Cells(k, 1).Value Then
    total_vol = total_vol + Cells(k, 7).Value
    Cells(ticker_row, 12).Value = total_vol
    Else
    total_vol = 0
    ticker_row = ticker_row + 1
    End If

Next k

'conditional formatting percentage change for color
For m = 2 To LRow
    
    If (Cells(m, 11).Value > 0 And Cells(m, 11).Value <> "") Then
    Cells(m, 11).Interior.ColorIndex = 4
    Else
    Cells(m, 11).Interior.ColorIndex = 3
    End If
    If Cells(m, 11).Value = "" Then
    Cells(m, 11).Interior.ColorIndex = xlNone
    End If
    
Next m

End Sub



