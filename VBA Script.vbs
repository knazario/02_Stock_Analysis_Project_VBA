Sub stocks():

'Declaring all necessary variables to pull in data, cylce through loops and hold various pieces
' data during looping
Dim i As Long
Dim j As Integer
Dim ticker As String
Dim opening As Double
Dim closing As Double
Dim volume As Variant
Dim sum_table_row As Long
Dim wscounter As Integer
Dim lastRow As Long
Dim gincrease_tick As String, gdecrease_tick As String, gvolume_tick As String
Dim gvolume As Variant
Dim gincrease As Double
Dim gdecrease As Double


'Code borrowed from microsoft support site. This sets wscounter to the number of sheets in the workbook
wscounter = ActiveWorkbook.Worksheets.Count

' j loop will cycle through each sheet in the workbook and build the table to enter the info
' from the nested loop

For j = 1 To wscounter
    
    Worksheets(j).Activate
    
    'Code borrowed from a demo Adrien(Instructor) provided in class
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'set counter to 1 to offset the data being entered in the summary table to start on the 2nd row. 
    ' In nested loop, first iteration will add 1 to sum_table_row to enter info on the 2nd row. 
    ' Loop will then increase by one whenever it encounters a new ticker symbol
    sum_table_row = 1
    
    'Resetting values for greatest, increase, decrease and volume for each worksheet. Chose low number for 
    'greatest increase and high number for greatest decrease to ensure a ticker value replaces it. 
    ' Theoretically, all stocks could have a percent change higher than 0 in a year. 
    gincrease = -100
    gdecrease = 100
    gvolume = 0
    
    'Create Ticker summary table on row 1 starting at column "I" (9)
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To lastRow
' conditional to check if the ticker matches the previous ticker row
        If ticker = Cells(i, 1).Value Then
               closing = Cells(i, 6)
               volume = Cells(i, 7).Value + volume
               Cells(sum_table_row, 12).Value = volume
               Cells(sum_table_row, 10).Value = closing - opening
               Cells(sum_table_row, 11).Value = (closing - opening) / opening

        Else
            sum_table_row = sum_table_row + 1
            
            If i = 2 Then
            
            ElseIf gincrease < (closing - opening) / opening Then
                gincrease = (closing - opening) / opening
                gincrease_tick = ticker
            
            
            ElseIf gdecrease > (closing - opening) / opening Then
                gdecrease = (closing - opening) / opening
                gdecrease_tick = ticker
            
            End If
            
            If gvolume < volume Then
                gvolume = volume
                gvolume_tick = ticker
            End If
            
            ticker = Cells(i, 1).Value
            opening = Cells(i, 3).Value
            Cells(sum_table_row, 9).Value = ticker
            volume = Cells(i, 7).Value
            Cells(sum_table_row, 12).Value = volume
        End If
    Next i

    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease "
    Range("N4").Value = "Greatest Total Volume"

    Range("O2").Value = gincrease_tick
    Range("O3").Value = gdecrease_tick
    Range("O4").Value = gvolume_tick

    Range("P2").Value = gincrease
    Range("P3").Value = gdecrease
    Range("P4").Value = gvolume
Next j

End Sub

