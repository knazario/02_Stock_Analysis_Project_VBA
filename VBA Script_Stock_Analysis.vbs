Sub stocks():
'Declaring all necessary variables to pull in data, cylce through loops
'and hold various pieces data during looping
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

'Sets wscounter to the number of sheets in the workbook
wscounter = ActiveWorkbook.Worksheets.Count

' j loop cycles through each sheet in the workbook and builds the summary table header
' prior to starting nested (i) loop.
For j = 1 To wscounter
    Worksheets(j).Activate

    'Calculates last non-empty row in dataset
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'set counter to 2 to offset the data being entered in the summary table to start on the 2nd row.
    sum_table_row = 2

    ' Resetting value of ticker to arbitrary string to ensure a match does not occur on the first iteration
    'the i loop after changing sheets.
    ticker = "stocks"

    'Resetting values for greatest, increase, decrease and volume for each worksheet.
    gincrease = 0
    gdecrease = 0
    gvolume = 0
    'Resetting greatest increase and decrease ticker flag in rare instance one of these values is not replaced by the data, they will print
    'as the greatest value and should be a sign of an error.
    gincrease_tick = "All values less than 0"
    gdecrease_tick = "All values greater than 0"

    'Create Ticker summary table on row 1 starting at column "I" (9)
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

' Loop through each row in the worksheet
    For i = 2 To lastRow
        ' Checks if first row (i = 2), brings in ticker, opening and sets volume and adds ticker to summary table
        If i = 2 Then
            ticker = Cells(i, 1).Value
            opening = Cells(i, 3).Value
            volume = Cells(i, 7).Value
            Cells(sum_table_row, 9).Value = ticker

        'Checks if current row does not equal (<>) the row below. This will find the last row of each ticker symbol          
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Store closing and update volume from current row (last row of ticker), and add values to summary table
            closing = Cells(i, 6)
            volume = volume + Cells(i, 7).Value
            Cells(sum_table_row, 12).Value = volume
            Cells(sum_table_row, 10).Value = closing - opening
            Cells(sum_table_row, 11).Value = (closing - opening) / opening
            
            'checks if overall greatest increase is less than most recently calculated ticker's increase. If true, replace greatest increase values
            If gincrease < (closing - opening) / opening Then
                gincrease = (closing - opening) / opening
                gincrease_tick = ticker

            'checks if current greatest decrease is greater than most recently calculated ticker's decrease. If true, replace greatest decrease values
            ElseIf gdecrease > (closing - opening) / opening Then
                gdecrease = (closing - opening) / opening
                gdecrease_tick = ticker
            End If

            'checks if current greatest volume is less than most recently calculated ticker's total volume. If true, replace greatest volume values
            If gvolume < volume Then
                gvolume = volume
                gvolume_tick = ticker
            End If
            'Updates ticker and opening to new ticker symbol referencing row below current row ((Cells(i+1,1) is the start of new ticker symbol),
            'and reset volume to 0 for new ticker (will update on following iterations) 
            sum_table_row = sum_table_row + 1
            ticker = Cells(i + 1, 1).Value
            opening = Cells(i + 1, 3).Value
            Cells(sum_table_row, 9).Value = ticker
            volume = 0

        ' Else will be all rows (other than the first row where i = 2) where the current row and row below (i+1) match. 
        ' Only action needed is adding current row volume to totalf volume for current ticker iteration.
        Else
            volume = Cells(i, 7).Value + volume
        End If
    Next i

'Creates Greatest increase,decrease and volume summary table grid
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease "
    Range("N4").Value = "Greatest Total Volume"

'Enters values of greatest increase, decrease and volume into summary table
    Range("O2").Value = gincrease_tick
    Range("O3").Value = gdecrease_tick
    Range("O4").Value = gvolume_tick

    Range("P2").Value = gincrease
    Range("P3").Value = gdecrease
    Range("P4").Value = gvolume
Next j
End Sub