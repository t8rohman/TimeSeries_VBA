Attribute VB_Name = "Module3"
Option Explicit
Public nr_decompose As Integer
' This module contains the coding for the Deseasonalized Data

Sub decompose_forecast()
    Dim nr As Integer, sea_idx As Integer, sea_n As Integer, cyc_n As Integer, i As Integer, fore_n As Integer
    
    nr = Cells(Rows.Count, 2).End(xlUp).Row
    Range("B2:C" & nr).Copy
    Worksheets("seasonality_decompose").Select
    Range("B3").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    nr = nr + 1 ' to account the first 2 rows
    nr_decompose = nr
    
    Range("A3").FormulaR1C1 = "1" ' fill the t period
    Range("A3").AutoFill Range("A3:A" & nr), xlFillSeries
    
    MsgBox ("Now, you will be asked the regularity of the season. If you cancel or don't type anything, it would be automatically set to 12.")
    sea_n = Application.InputBox(Prompt:="Enter the (n) for regularity period.", Type:=1) ' calculate for seasonality
    If sea_n = 0 Then
        sea_n = 12
    End If
    
    sea_idx = sea_n \ 2 + 3
    
    Range("D1:D2").FormulaR1C1 = "(" & sea_n & ") MA" ' calculate for moving average
    Range("D" & (sea_n / 2) + 2).FormulaR1C1 = "=AVERAGE(R[-" & (sea_n / 2) - 1 & "]C[-1]:R[" & (sea_n / 2) & "]C[-1])" '
    Range("D" & (sea_n / 2) + 2).AutoFill Destination:=Range("D" & (sea_n / 2) + 2 & ":D" & nr - (sea_n / 2))
    
    With Range("E" & (sea_n / 2) + 3)
        .FormulaR1C1 = "=AVERAGE(R[-1]C[-1]:RC[-1])" ' calculate for centred moving average
        .AutoFill Destination:=Range("E" & (sea_n / 2) + 3 & ":E" & nr - (sea_n / 2))
    End With
    
    With Range("F" & (sea_n / 2) + 3)
        .FormulaR1C1 = "=(RC[-3]/RC[-1])" ' calculate for seasonal relative
        .AutoFill Destination:=Range("F" & (sea_n / 2) + 3 & ":F" & nr - (sea_n / 2))
    End With
    
    MsgBox ("Now, you will be asked how many cycles of the season. If you cancel or don't type anything, it would be automatically set to 1.")
    MsgBox ("This version is set to can only count until 5 cycles of the season")
    
    Do
        cyc_n = Application.InputBox(Prompt:="Enter the (n) for cycle of the season.", Type:=1) ' calculate for seasonality
        If cyc_n = 0 Then
            cyc_n = 1
        End If
    Loop Until cyc_n <= 5
    
    With Range("G" & sea_idx)
        .Formula = "=AVERAGE(F" & sea_idx & ",F" & sea_idx + sea_n & ",F" & sea_idx + (sea_n * 2) & _
            ",F" & sea_idx + (sea_n * 3) & ",F" & sea_idx + (sea_n * 4) & ")"
        .AutoFill Destination:=Range("G" & sea_idx & ":G" & sea_idx + sea_n - 1)
    End With
    
    For i = 0 To sea_n - 1 ' calculate deseasoned for the data after
        Range("H" & sea_idx + i).FormulaR1C1 = "=(RC[-5]/R" & sea_idx + i & "C7)"
    Next i
    With Range("H" & sea_idx & ":H" & sea_idx + sea_n - 1)
        .AutoFill Destination:=Range("H" & sea_idx & ":H" & nr)
    End With
    
    i = 0
    Do While Range("H" & 3 + i).Value = "" ' calculate deseasoned for the data before
        Range("H" & 3 + i).FormulaR1C1 = "=(RC[-5]/R" & sea_n + 3 + i & "C7)"
        i = i + 1
    Loop
    
    ' ## Creating Trend Line Values for Forecasting
    
    MsgBox ("Now, you will be asked for how many period forward you want to forecast. If you cancel or don't type anything, it would be automatically set to 6.")
    MsgBox ("This version is set to can only count until 24 (n) period forward.")
    
    Do
        fore_n = Application.InputBox(Prompt:="Enter the (n) for Period Forecast.", Type:=1)
        If fore_n = 0 Then
            fore_n = 6
        End If
    Loop Until fore_n <= 24
    
    Range("P4").Formula2R1C1 = "=LINEST(R3C8:R" & nr & "C8,R3C1:R" & nr & "C1)" ' make the regression to put back the seasonality index
    Range("A" & nr - 3 & ":B" & nr).AutoFill Range("A" & nr - 3 & ":B" & nr + fore_n), xlFillDefault
    
    Range("J3").FormulaR1C1 = "=R4C16*RC[-9]+R4C17"
    Range("J3").AutoFill Destination:=Range("J3:J" & nr + fore_n), Type:=xlFillDefault
    
    ' ## Forecasting
    
    For i = 0 To sea_n - 1 ' calculate predictions for the data after
        Range("K" & sea_idx + i).FormulaR1C1 = "=(RC[-1]*R" & sea_idx + i & "C7)"
    Next i
    With Range("K" & sea_idx & ":K" & sea_idx + sea_n - 1)
        .AutoFill Destination:=Range("K" & sea_idx & ":K" & nr + fore_n)
    End With
    
    i = 0
    Do While Range("K" & 3 + i).Value = "" ' calculate predictions for the data before
        Range("K" & 3 + i).FormulaR1C1 = "=(RC[-1]*R" & sea_n + 3 + i & "C7)"
        i = i + 1
    Loop
    
    ' ## Error
    
    With Range("L3") ' error
        .FormulaR1C1 = "=(RC[-1]-RC[-9])"
        .AutoFill Destination:=Range("L3:L" & nr)
    End With
    
    With Range("M3") ' absolute error
        .FormulaR1C1 = "=ABS(RC[-1])"
        .AutoFill Destination:=Range("M3:M" & nr)
    End With
    
    With Range("N3") ' error squarred
        .FormulaR1C1 = "=(RC[-1]^2)"
        .AutoFill Destination:=Range("N3:N" & nr)
    End With
    
    Range("Q6").Formula = "=AVERAGE(M3:M" & nr & ")"
    Range("Q7").Formula = "=AVERAGE(N3:N" & nr & ")"
    Range("Q8").Formula = "=SQRT(Q7)"
    
End Sub

Sub decompose_reset()
    Dim nr As Integer

    nr = Cells(Rows.Count, 1).End(xlUp).Row
    If WorksheetFunction.CountA(Range("A3:B3")) = 0 Then
        Exit Sub
    End If
    Range("A3:Q" & nr).ClearContents
    
    ' The label for parameter and info
    
    Range("P3").Formula2R1C1 = "Beta"
    Range("Q3").Formula2R1C1 = "Alpha"
    Range("P6").Formula2R1C1 = "MAE"
    Range("P7").Formula2R1C1 = "MSE"
    Range("P8").Formula2R1C1 = "RMSE"
    
    Range("P3:Q3").Font.Bold = True
    Range("P6:P8").Font.Bold = True

End Sub

Sub decompose_graph()

    Dim nr As Integer
    nr = Cells(Rows.Count, 1).End(xlUp).Row
    Sheets("graph").Select
    ActiveChart.ChartArea.Select
    With ActiveChart
        .SeriesCollection(1).Values = Worksheets("seasonality_decompose").Range("C3:C" & nr)
        .SeriesCollection(2).Values = Worksheets("seasonality_decompose").Range("K3:K" & nr)
        .Axes(xlCategory).CategoryNames = Worksheets("seasonality_decompose").Range("B3:B" & nr)
        .ChartTitle.Text = "Forecast Using Deseasonalized Method"
    End With

End Sub

Sub decompose_deseasonalized_graph()

    Dim nr As Integer
    Sheets("seasonality_graph").Select
    ActiveChart.ChartArea.Select
    With ActiveChart
        .SeriesCollection(1).Values = Worksheets("seasonality_decompose").Range("H3:H" & nr_decompose)
        .Axes(xlCategory).CategoryNames = Worksheets("seasonality_decompose").Range("B3:B" & nr_decompose)
        .ChartTitle.Text = "Observed Smoothed Data (Deseasonalized)"
    End With

End Sub
