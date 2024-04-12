Sub RunStockAnalysisForAllSheets()
'After the coding for Sub StockAnalysis was configured, this loop was set up to run the macro on every sheet
    'Set variable to hold # of worksheets within workbook
    Dim SHEET As Worksheet
    'Loop that runs the Sub StockAnalysis macro for every sheet
    For Each SHEET In Worksheets
        SHEET.Activate
        Call StockAnalysis
    Next SHEET
End Sub

Sub StockAnalysis()

' creating the new columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
'apply percentage style and format to Colum K results
'apply percentage style and format to Q2 & Q3 results
'autofit created columns
    Columns("K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q2", "Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Columns("J").Select
    Selection.NumberFormat = "0.00"
    Columns("I:Q").EntireColumn.AutoFit
    
'MAIN CALCULATION LOOP
'ATTENTION: Aspects of the following code were generated with help from Academic Tutor and Xpert Learning Assistant
'ATTENTION: Aspects of the following code were generated with help from Academic Tutor and Xpert Learning Assistant
    'Declaration of variables
    Dim LastRow As Long
    Dim T As Long
    Dim R As Long
    Dim Opener As Double
    Dim Closer As Double
    Dim Percentage As Double
    
    'Calculate the last row for the given sheet
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    'MsgBox (LastRow) - previously used to confirm accuracy of the above
'Set initial values
    'Setting iterator for corresponding results columns
    T = 2
    'Setting initial stock's first daily volume to be the relative starting point for volume aggregate result
    Cells(T, 12).Value = Cells(T, 7).Value
    'Setting stock's ticker value to be the relative starting point for ticker result
    Cells(T, 9).Value = Cells(T, 1).Value
    'Setting stock's opening value to be the relative starting point for yearly change result
    Opener = Cells(T, 3).Value
    'Loop to populate Ticker, Yearly Change, Percentage Change and Total Stock Volume
    For R = 2 To LastRow
        If Cells(R, 1).Value = Cells(R + 1, 1).Value Then
            Cells(T, 12).Value = Cells(T, 12).Value + Cells(R + 1, 7).Value
        Else
        'Conditions to be met when the Ticker Value is about to change (eg. the next row presents a different stock name)!
            'Yearly Change
            Closer = Cells(R, 6).Value
            Cells(T, 10).Value = Closer - Opener
            'Percentage Change
            Percentage = (Closer - Opener) / Opener
            Cells(T, 11).Value = Percentage
            'Moving down a row for Total Stock Volume
            T = T + 1
            'Resetting the previously defined ticker name to the new Ticker name column
            Cells(T, 9).Value = Cells(R + 1, 1).Value
            'Reset the previously defined stock volume start point to start over with a new ticker name
            Cells(T, 12).Value = Cells(R + 1, 7).Value
            'Reset the previously defined opener for when the condition is not met (tracking the new stocks opening value)
            Opener = Cells(R + 1, 3).Value
        End If
    Next R
'Once the values have poulated due to our above For/Next loop, we can collect the data and provide the requested annual calculations for each stock
    'Set variables for finding the max and min of the Percentage Change results and the max of the Total Stock Volume results
    Dim Max As Double
    Dim Min As Double
    Dim MaxVol As Double
    Dim L As Long
    'Utilize Max and Min functions to assign reuslt to variables
    Max = Application.WorksheetFunction.Max(Range("K1:K" & LastRow))
    Min = Application.WorksheetFunction.Min(Range("K1:K" & LastRow))
    MaxVol = Application.WorksheetFunction.Max(Range("L1:L" & LastRow))
    'Place variable value within desired cells
    Cells(2, 17).Value = Max
    Cells(3, 17).Value = Min
    Cells(4, 17).Value = MaxVol
    'Run loop to locate Max and Min % change values and populate the match's corresponding ticker value to the outcome cell (Greatest % Increase/Decrease)
    'Do the same with a final elseif cycle for the Ticker associated with the largest Total Stock Volume
    For L = 2 To LastRow
        If Cells(L, 11).Value = Max Then
            Cells(2, 16).Value = Cells(L, 9)
        ElseIf Cells(L, 11).Value = Min Then
            Cells(3, 16).Value = Cells(L, 9).Value
        ElseIf Cells(L, 12).Value = MaxVol Then
            Cells(4, 16).Value = Cells(L, 9).Value
        End If
    Next L
    'Re-autofit column Q to display information better
     Columns("Q").EntireColumn.AutoFit
     'Run 3 Loops to set conditional formatting determining the color displayed for positive/negative/zero results within the Yearly Change, Percentage Change columns as well as the Greatest % Decrease/Increase cells
     Dim C As Long
     Dim P As Long
     Dim Q As Integer
        For C = 2 To LastRow
            If Cells(C, 10).Value > 0 Then
                Cells(C, 10).Interior.ColorIndex = 4
            ElseIf Cells(C, 10).Value = 0 Then
                Cells(C, 10).Interior.ColorIndex = 0
            ElseIf Cells(C, 10).Value < 0 Then
                Cells(C, 10).Interior.ColorIndex = 3
            End If
        Next C
        For P = 2 To LastRow
            If Cells(P, 11).Value > 0 Then
                Cells(P, 11).Interior.ColorIndex = 4
            ElseIf Cells(P, 11).Value = 0 Then
                Cells(P, 11).Interior.ColorIndex = 0
            ElseIf Cells(P, 11).Value < 0 Then
                Cells(P, 11).Interior.ColorIndex = 3
            End If
        Next P
        For Q = 2 To 3
            If Cells(Q, 17).Value > 0 Then
                Cells(Q, 17).Interior.ColorIndex = 4
            ElseIf Cells(Q, 17).Value < 0 Then
                Cells(Q, 17).Interior.ColorIndex = 3
            End If
        Next Q
    'Reselection of initial Cell on Worksheet
    Range("A1").Select
End Sub

