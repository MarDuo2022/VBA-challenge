Attribute VB_Name = "Module1"
Sub teststock()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total stock Volume"
    
    Dim NumRows As Integer
    NumRows = 30000
        'ideally numrows = count(number of rows on each sheet)
        
    Dim i As Integer
    Dim k As Integer
    k = 1

'Output ticker symbols

    For i = 2 To NumRows
       If Cells(i, 1) = Cells(i - 1, 1) Then
       Else:
        k = k + 1
        Cells(k, 9) = Cells(i, 1)
       End If
    Next i
    
End Sub
