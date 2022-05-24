Sub StockParser()

    'establish variables
    Dim YearlyChange As Double
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    TotalVolume = 0
    Dim Tickerid As String
    Dim Final As Long
    Dim i As Long
    Dim i2 As Long
    Dim i3 As Long
    Dim i4 As Long
    'put values into counter variables
    i = 2
    i2 = 2
    i3 = 2
    i4 = 2
    YearlyChange = 0
    YearlyOpen = 0
    YearlyClose = 0
    
    'Create table header
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
 
    'Find last row
    Final = Cells(Rows.Count, 1).End(xlUp).Row
 
    'Create for loop to return ticker type
    For c = 2 To Final

        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
            Tickerid = Cells(c, 1).Value '
            Cells(i, 9).Value = Tickerid
            i = i + 1
        End If
    
    Next c
    
    'total stock volume
    For c = 2 To Final
    
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
        TotalVolume = TotalVolume + Cells(c, 7).Value
        Cells(i2, 12).Value = TotalVolume
        i2 = i2 + 1
        
        TotalVolume = 0
        
        Else
        
        TotalVolume = TotalVolume + Cells(c, 7).Value
        
        End If
        
        Next c
        
    ' yearly open and close
    YearlyOpen = Cells(2, 3).Value
    
    For c = 2 To Final
        
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
        YearlyClose = Cells(c, 6).Value
        YearlyChange = YearlyClose - YearlyOpen
        Cells(i3, 10).Value = YearlyChange
        i3 = i3 + 1
        
        YearlyOpen = Cells(c + 1, 3).Value
        End If
        
        YearlyClose = 0
        YearlyChange = 0
    
    Next c
    
    For c = 2 To Final
    
        If Cells(c, 10).Value <= 0 Then
        Cells(c, 10).Interior.ColorIndex = 3
        
        ElseIf Cells(c, 10).Value > 0 Then
        Cells(c, 10).Interior.ColorIndex = 4
        
        End If
    
    Next c
    
    'percent change
    YearlyOpen = Cells(2, 3).Value
    
    For c = 2 To Final
        
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
        YearlyClose = Cells(c, 6).Value
        YearlyChange = YearlyClose - YearlyOpen
        PercentChange = (YearlyChange / YearlyOpen) * 100
        Cells(i4, 11).Value = Round(PercentChange, 2) & "%"
        i4 = i4 + 1
        
        YearlyOpen = Cells(c + 1, 3).Value
        
        End If
        
        YearlyClose = 0
        YearlyChange = 0
        PercentChange = 0
    
    Next c
        
End Sub
