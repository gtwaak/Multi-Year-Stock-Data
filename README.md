# Multi-Year-Stock-Data
# Unit 2 | Assignment - The VBA of Wall Street

## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, choose your assignment from Easy, Moderate, or Hard below.

### Files

* [Test Data](Resources/alphabtical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.


### Moderate

* Create a script that will loop through all the stocks for one year for each run and take the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.


### Copyright

Coding Boot Camp © 2019. All Rights Reserved.

VBA Code

Sub counter()

    Dim ticker As String
    Dim TotalVol As Double
    Dim lastRow As Long
    Dim currRow As Double
    Dim count As Integer
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim lastRow2 As Long
    
    
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 12).Value = "TotalVol"
        ws.Cells(1, 10).Value = "yearlyChange"
        ws.Cells(1, 11).Value = "percentChange"
        
        lastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        currRow = 2
        count = 0
        
    
    'setting for loop which iterates through rows
    
    For i = 2 To lastRow
        TotalVol = TotalVol + ws.Cells(i, 7).Value
        count = count + 1
        'comparing intiital row to next rows value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        'inputing values in column if ticker is not the same
        
            ws.Cells(currRow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(currRow, 12).Value = TotalVol
            yearlyChange = ws.Cells(i, 6) - ws.Cells(i - count + 1, 3).Value
            ws.Cells(currRow, 10).Value = yearlyChange
            If yearlyChange = 0 Then
                ws.Cells(currRow, 11).Value = "0"
            Else
            
                percentChange = ws.Cells(i - count + 1, 3).Value / yearlyChange
                ws.Cells(currRow, 11).Value = percentChange
            
            End If
            
            TotalVol = 0
            currRow = currRow + 1
            count = 0
            
            
        End If
    
    Next i
       
       lastRow2 = ws.Cells(Rows.count, 11).End(xlUp).Row
    
        For i = 2 To lastRow2
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 11).Value < 0 Then
            
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        
        
        Next i
    
Next ws

End Sub

