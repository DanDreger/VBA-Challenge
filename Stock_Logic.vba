'Subfunction which loops through all the sheets and runs stockAnalysis on all of them-----------------------------
Sub analyzeAllSheets()
    'initialize stocksheet
    Dim stockSheet As Worksheet
    Application.ScreenUpdating = False
    'loop through all the worksheets
    For Each stockSheet In Worksheets
        'select each sheet
        stockSheet.Select
        'call stockanalysis() on each sheet
        Call stockAnalysis
    Next
    Application.ScreenUpdating = True
End Sub
'End Subfunction



'subfunction to be run on all stockSheets ---------------------------------------------------------
Sub stockAnalysis():



'add headers at the top of the colums--------------------------------------
'declare a variant array
Dim columnHeaders(9 To 17) As String
columnHeaders(9) = "Ticker"
columnHeaders(10) = "Yearly Change"
columnHeaders(11) = "Percent Change"
columnHeaders(12) = "Total Stock Volume"
columnHeaders(13) = ""
columnHeaders(14) = ""
columnHeaders(15) = ""
columnHeaders(16) = "Ticker"
columnHeaders(17) = "Value"
'loop from columns 9 to 17
For i = 9 To 17
             'show the header in the appropriate column
             Cells(1, i).Value = columnHeaders(i)
Next i
'end add headers at the top of the colums----------------------------------------------------------


'Add labels for stocks with greatest gain, loss, and total volume-----------------------------------------------
Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decrease"
Range("o4").Value = "Greatest Total Volume"


'Establish Vaiables for the data --------------------------------------------------------------------------------------
Dim LR As Long
LR = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
Dim currentResultsRow As Integer
currentResultsRow = 2
Dim stockDailyVolume As Variant
Dim stockTotalVolume As Variant
stockTotalVolume = 0
Dim stockStartingValue As Variant
Dim stockEndingValue As Variant
Dim currentTicker As String
Dim nextTicker As String
Dim previousTicker As String
Dim percentChanged As Variant
Dim greatestIncrease As Variant
greatestIncrease = 0
Dim greatestDecrease As Variant
greatestDecrease = 0
Dim greatestVolume As Variant
greatestVolume = 0
Dim stockChange As Integer
Dim startingVolume As Long
Dim firstCompanyRow As Long
'End Establish Vaiables for the data --------------------------------------------------------------------------------------



'Loop through the data and analyze----------------------------------------------------------------------------------
For i = 2 To LR

            'establish variables for previous ticker, current ticker and next ticker
            previousTicker = Cells(i - 1, 1)
            currentTicker = Cells(i, 1)
            nextTicker = Cells(i + 1, 1)
            

            'set var for daily volume
            stockDailyVolume = Cells(i, 7).Value
            
            'Calculate stockTotalVolume part of second solution
            stockTotalVolume = stockTotalVolume + stockDailyVolume

            If (currentTicker = previousTicker And currentTicker = nextTicker) Then
            'do nothing. this saves the macro from running the rest of the elseifs for 99% of the cells
            
            'run this block the first time the loop comes upon a new stock
            ElseIf (previousTicker <> currentTicker) Then
                    stockStartingValue = Cells(i, 3).Value
                    firstCompanyRow = i
                    
            'run this block the last time a loop iterates through a given stock
            ElseIf (currentTicker <> nextTicker) Then
            
            'set stockendingvalue
            stockEndingValue = Cells(i, 6).Value

            'define yearly change
            yearlyChange = stockEndingValue - stockStartingValue

            'display yearly change
            Cells(currentResultsRow, 10) = yearlyChange
            
            'define percent changed
            percentChanged = (stockEndingValue - stockStartingValue) / stockStartingValue

            'display ticker and percentChanged
            Cells(currentResultsRow, 9) = currentTicker
            Cells(currentResultsRow, 11) = percentChanged
    
            'compare percentChanged to greatestIncrease and greatestDecrease
            If (percentChanged < 0 And percentChanged < greatestDecrease) Then
                        greatestDecrease = percentChanged
                        'display new greatest % changed
                        Cells(3, 17).Value = FormatPercent(greatestDecrease)
                        Cells(3, 16).Value = currentTicker
            ElseIf (percentChanged > 0 And percentChanged > greatestIncrease) Then
                        greatestIncrease = percentChanged
                        Cells(2, 17).Value = FormatPercent(greatestIncrease)
                        Cells(2, 16).Value = currentTicker
            End If
            
            'Display stockDailyVolume
            Cells(currentResultsRow, 12).Value = stockTotalVolume
            
            'compareTotalVolume with prevous total volume
            If stockTotalVolume > greatestVolume Then
                        greatestVolume = stockTotalVolume
                        Cells(4, 17).Value = FormatNumber(greatestVolume)
                        Cells(4, 16).Value = currentTicker
            End If
            
            'color code the total Change column
            If Cells(currentResultsRow, 10) < 0 Then
                        Cells(currentResultsRow, 10).Interior.Color = 8747775
            Else
                        Cells(currentResultsRow, 10).Interior.Color = 5296274
            End If
            
            'resetstocktotalvolume
            stockTotalVolume = 0
            
            'reset starting and ending values
            stockStartingValue = 0
            stockEndingValue = 0
            'set currentResultsRow one line down
            currentResultsRow = currentResultsRow + 1
                       
            End If
              
Next i
'End Loop through the data and analyze-------------------------------------------------------------------------------


'Format cells to display as a percentage
    Range("K:K,Q2,Q3").Select
    Range("Q3").Activate
    Selection.NumberFormat = "0.00%"

' Autofit to display data
ActiveSheet.UsedRange.EntireColumn.AutoFit


End Sub
'End Subtask------------------------------------------------------------------------------------------------------
 

