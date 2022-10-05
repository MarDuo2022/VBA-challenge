Attribute VB_Name = "stockcodeMD"
Sub teststock()
'.Interior.ColorIndex   for conditional formatting using VBA
For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total stock Volume"
    
    Dim numrows As Long
    'NumRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'ideally numrows = count(number of rows on each sheet)
    numrows = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
    Dim i As Long
    
    'Define row of output for the corresponding ticker
    Dim rowout As Integer
    rowout = 1

    Dim rowopenyearold As Integer
    Dim rowopenyearnew As Integer
    Dim accumulatedV As LongLong
    accumulatedV = 0
    
    For i = 2 To numrows
       If ws.Cells(i, 1) = ws.Cells(i - 1, 1) Then
        accumulatedV = accumulatedV + ws.Cells(i, 7)
        
        Else
       
        'Output ticker symbols
            rowout = rowout + 1
            ws.Cells(rowout, 9) = ws.Cells(i, 1)
        'Output yearly change
            rowopenyearold = rowopenyearnew 'Open price for previous ticker
            rowopenyearnew = i  'Open price for new ticker
            
' [For checking] MsgBox rowopenyearold
               ' MsgBox rowopenyearnew
            
            If i = 2 Then
            Else: ws.Cells(rowout - 1, 10) = ws.Cells(i - 1, 6) - ws.Cells(rowopenyearold, 3)
      
                  If ws.Cells(rowout - 1, 10) < 0 Then
                    ws.Cells(rowout - 1, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(rowout - 1, 10).Interior.ColorIndex = 4
                  End If
                  
      'Output Percent change
                  ws.Cells(rowout - 1, 11) = ws.Cells(rowout - 1, 10) / ws.Cells(rowopenyearold, 3) * 100 & " %"

              
      'Output Total volume of stock
                  ws.Cells(rowout - 1, 12) = accumulatedV
            End If
            accumulatedV = ws.Cells(i, 7)
        
       End If
    Next i
'BONUS
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    Dim pcincrease, pcdecrease As Double
    Dim totalV As LongLong
    Dim numrows2 As Integer
    Dim tickerpcincrease, tickerpcdecrease As String
    pcincrease = 0
    pcdecrease = 0
    totalV = 0
    
    numrows2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To numrows2
    ' just the number of rows of second table NumRows
    
        If pcincrease < ws.Cells(i, 11) Then
            pcincrease = ws.Cells(i, 11)
            tickerpcincrease = ws.Cells(i, 9)
        End If
        If pcdecrease > ws.Cells(i, 11) Then
            pcdecrease = ws.Cells(i, 11)
            tickerpcdecrease = ws.Cells(i, 9)
        End If
        If totalV < ws.Cells(i, 12) Then
            totalV = ws.Cells(i, 12)
            tickertotalV = ws.Cells(i, 9)
        End If
    Next i
    
    For j = 2 To numrows
        If ws.Cells(j, 11) = pcincrease Then
            ws.Range("P2").Value = tickerpcincrease
            ws.Range("Q2").Value = pcincrease * 100 & " %"
        End If
        If ws.Cells(j, 11) = pcdecrease Then
            ws.Range("P3").Value = tickerpcdecrease
            ws.Range("Q3").Value = pcdecrease * 100 & " %"
        End If
        If ws.Cells(j, 12) = totalV Then
            ws.Range("P4").Value = tickertotalV
            ws.Range("Q4").Value = totalV
        End If
    Next j
Next ws
End Sub
