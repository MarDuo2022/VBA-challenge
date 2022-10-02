Attribute VB_Name = "Module1"
Sub teststock()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total stock Volume"
    
    Dim NumRows As Integer
    NumRows = 400
        'ideally numrows = count(number of rows on each sheet)
        
    Dim i As Integer
    
    'Define row of output for the corresponding ticker
    Dim rowout As Integer
    rowout = 1

    Dim rowopenyearold As Integer
    Dim rowopenyearnew As Integer
    
    For i = 2 To NumRows
       If Cells(i, 1) = Cells(i - 1, 1) Then
       Else:
        'Output ticker symbols
            rowout = rowout + 1
            Cells(rowout, 9) = Cells(i, 1)
        'Output yearly change
            rowopenyearold = rowopenyearnew 'Open price for previous ticker
            rowopenyearnew = i  'Open price for new ticker
            
' [For checking] MsgBox rowopenyearold
               ' MsgBox rowopenyearnew
            
            If i = 2 Then
            Else: Cells(rowout - 1, 10) = Cells(i - 1, 6) - Cells(rowopenyearold, 3)
            End If
       End If
    Next i
    
End Sub
