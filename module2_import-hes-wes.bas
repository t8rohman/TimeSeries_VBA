Attribute VB_Name = "Module2"
Option Explicit
Option Base 1
' This module contains the coding to calculate the forecasting using Winter's and Holt's
' This also contains the coding to import the data (the main worksheet)

Sub seasonal_forecast()
    Dim nr As Integer, sea_n As Integer, wes_n As Integer

    nr = Cells(Rows.Count, 2).End(xlUp).Row
    Range("B2:C" & nr).Copy
    Worksheets("seasonality_wes").Select
    Range("A3").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    nr = nr + 1 ' to account the first 2 rows
    
    MsgBox ("Now, you will be asked the regularity of the sesason. If you cancel or don't type anything, it would be automatically set to 12.")
    sea_n = Application.InputBox(Prompt:="Enter the (n) for regularity period.", Type:=1) ' calculate for seasonality
    If sea_n = 0 Then
        sea_n = 12
    End If
    Range("J7").Value = sea_n
    sea_n = sea_n + 2 '  to account the first 2 rows
    
    Range("C" & sea_n).Formula = "=AVERAGE(B3:B" & sea_n & ")" ' calculating initial level
    
    Worksheets("seasonality_wes").Range("B3:B" & sea_n).Copy _
        (Worksheets("seasonality_wes").Range("P3"))  ' calculating the initial value for trend
    Worksheets("seasonality_wes").Range("B" & sea_n + 1 & ":B" & 2 * (sea_n - 2) + 2).Copy _
        (Worksheets("seasonality_wes").Range("Q3"))
    With Range("R3")
        .FormulaR1C1 = "=RC[-1]-RC[-2]"
        .NumberFormat = "0.0"
        .AutoFill Destination:=Range("R3:R" & sea_n), Type:=xlFillDefault
    End With
    With Range("S3")
        .FormulaR1C1 = "=RC[-1]/R7C10"
        .NumberFormat = "0.0"
        .AutoFill Destination:=Range("S3:S" & sea_n), Type:=xlFillDefault
    End With
    With Range("R" & sea_n + 2)
        .FormulaR1C1 = "Initial Trend"
        .Font.Bold = True
    End With
    Range("S" & sea_n + 2).Formula = "=AVERAGE(S3:S" & sea_n & ")"
    Range("D" & sea_n).Formula = "=S" & sea_n + 2
    
    With Range("E3") ' calculating initial seasonal
        .Formula = "=B3/$C$" & sea_n
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("E3:E" & sea_n), Type:=xlFillDefault
    End With
    
    With Range("C" & sea_n + 1) ' for level
        .Formula = "=wes_level($J$4,B" & sea_n + 1 & ",C" & sea_n & ",D" & sea_n & ",E3)"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("C" & sea_n + 1 & ":C" & nr)
    End With
    With Range("D" & sea_n + 1) ' for trend
        .Formula = "=wes_trend($J$5,C" & sea_n + 1 & ",C" & sea_n & ",D" & sea_n & ")"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("D" & sea_n + 1 & ":D" & nr)
    End With
    With Range("E" & sea_n + 1) ' for season
        .Formula = "=wes_season($J$6,B" & sea_n + 1 & ",C" & sea_n + 1 & ",E3)"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("E" & sea_n + 1 & ":E" & nr)
    End With
    
    ' ## Create forecast
    
    With Range("F" & sea_n + 1) '
        .Formula = "=wes_forecast(C" & sea_n & ",D" & sea_n & ",E3,1)"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("F" & sea_n + 1 & ":F" & nr)
    End With
    
    MsgBox ("Now, you will be asked for how many period forward you want to forecast. If you cancel or don't type anything, it would be automatically set to 6.")
    MsgBox ("Remember that period you want to forecast couldn't be a bigger number than the period regularity")
    Do
        wes_n = Application.InputBox(Prompt:="Enter the (n) for Period Forecast.", Type:=1) ' calculate for moving average
        If wes_n = 0 Then
            wes_n = 6
        ElseIf wes_n > sea_n - 2 Then
            MsgBox ("Period you want to forecast is higher than the period regularity! Enter the number again.")
        End If
    Loop Until sea_n - 2 >= wes_n
    
    Range("G" & nr + 1).FormulaR1C1 = "1" ' fill the t period
    Range("G" & nr + 1).AutoFill Range("G" & nr + 1 & ":G" & nr + wes_n), xlFillSeries
    
    Range("F" & nr + 1) = "=wes_forecast(R" & nr & "C3,R" & nr & "C4,R[-" & wes_n & "]C[-1],RC[1])" ' fill the forecast
    Range("F" & nr + 1).AutoFill Range("F" & nr + 1 & ":F" & nr + wes_n), xlFillSeries
    
    Range("A" & nr).AutoFill Destination:=Range("A" & nr & ":A" & nr + wes_n), Type:=xlFillDefault ' autofill period
    
    ' ## Calculate error
    
    With Range("L" & sea_n + 1) ' error
        .FormulaR1C1 = "=(RC[-6]-RC[-10])"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("L" & sea_n + 1 & ":L" & nr), Type:=xlFillDefault
    End With
    With Range("M" & sea_n + 1) ' absolute error
        .FormulaR1C1 = "=ABS(RC[-1])"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("M" & sea_n + 1 & ":M" & nr), Type:=xlFillDefault
    End With
    With Range("N" & sea_n + 1) ' squarred error
        .FormulaR1C1 = "=RC[-1]^2"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("N" & sea_n + 1 & ":N" & nr), Type:=xlFillDefault
    End With
    
    Range("J10").Formula = "=AVERAGE(M" & sea_n + 1 & ":M" & nr & ")"
    Range("J11").Formula = "=AVERAGE(N" & sea_n + 1 & ":N" & nr & ")"
    Range("J12").Formula = "=SQRT(J11)"
    Range("J10:J12").NumberFormat = "0.000"
    
End Sub

Sub seasonal_solver()
    SolverAdd CellRef:="$J$4", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$J$5", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$J$6", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$J$4", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$J$5", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$J$6", Relation:=3, FormulaText:="0"
    SolverOk SetCell:="$J$12", MaxMinVal:=2, ValueOf:=0, ByChange:="$J$4:$J$6", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve (True)
End Sub

Sub main_openfile()
Dim UserRange As Range, DefaultRange As String, filename As Variant, nr As Integer, min_obs As Double
Dim tWB As Workbook, aWB As Workbook
On Error Resume Next

Set tWB = ThisWorkbook

filename = Application.GetOpenFilename(Title:="Select Your File")
' Error handling
If filename = False Then
    Exit Sub
End If

MsgBox ("Select the range you want to import, consisting of time period and the observed level. Be sure to put 2 rows, with period on left and observed.")

Workbooks.Open filename
DefaultRange = Selection.Address ' Selection before subroutine is executed

Set UserRange = Application.InputBox(Prompt:="Select a range to copy to the main sheet!", Title:="Instruction", Default:=DefaultRange, Type:=8)
Err.Clear
' Error handling
If UserRange Is Nothing Then
    Set aWB = ActiveWorkbook
    aWB.Close SaveChanges:=False
    Exit Sub
ElseIf UserRange.Columns.Count <> 2 Then
    MsgBox ("The data should be in 2 columns. Press the button again to retry!")
    Set aWB = ActiveWorkbook
    aWB.Close SaveChanges:=False
    Exit Sub
End If

Set aWB = ActiveWorkbook
aWB.Worksheets(1).Activate
UserRange.Select
Selection.Copy
aWB.Close SaveChanges:=False

tWB.Worksheets("data").Activate
Range("B2").Select
ActiveSheet.Paste
Range("A1").Select

nr = Cells(Rows.Count, 2).End(xlUp).Row
Range("A2:A4").AutoFill Destination:=Range("A2:A" & nr), Type:=xlFillSeries

' modifying the data in chart

min_obs = WorksheetFunction.Min(Range("C:C"))
ActiveSheet.ChartObjects("Chart 3").Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScale = min_obs - min_obs / 10

Range("A1").Select

End Sub

Sub main_reset()
Dim nr As Integer

nr = Cells(Rows.Count, 2).End(xlUp).Row
If WorksheetFunction.CountA(Range("B2:C" & nr)) = 0 Then
    Exit Sub
End If
Range("B2:C" & nr).Clear
Range("A5:A" & nr).Clear
ActiveSheet.ListObjects("Table1").Resize Range("$A$1:$C$4")

End Sub

Sub level_forecast()
Attribute level_forecast.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim nr As Integer, ma_n As Integer, hes_n As Integer
    Dim hes_level_ref As Double, hes_trend_ref As Double

    nr = Cells(Rows.Count, 2).End(xlUp).Row
    Range("B2:C" & nr).Copy
    Worksheets("leveltrend_all").Select
    Range("A3").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    nr = nr + 1 ' to account 2 first row
    Range("C3").FormulaR1C1 = "=AVERAGE(C[-1])" ' calculate for average forecast
    Range("C3").AutoFill Destination:=Range("C3:C" & nr + 1)
    
    MsgBox ("Now, you will be asked for the (n) for Moving Average. If you cancel or don't type anything, it would be automatically set to 1.")
    
    ma_n = Application.InputBox(Prompt:="Enter the (n) for Moving Average.", Type:=1) ' calculate for moving average
    If ma_n = 0 Then
        ma_n = 1
    End If
    Range("D1:D2").FormulaR1C1 = "Moving Average (" & ma_n & ")"
    ma_n = ma_n + 3 ' to account 2 first row
    Range("D" & ma_n).FormulaR1C1 = "=AVERAGE(R[-" & ma_n - 3 & "]C[-2]:R[-1]C[-2])" '
    Range("D" & ma_n).AutoFill Destination:=Range("D" & ma_n & ":D" & nr + 1)
        
    Range("E4").FormulaR1C1 = "=(R[-1]C[-3])"  ' calculate for naive forecast
    Range("E4").AutoFill Destination:=Range("E4:E" & nr + 1)
    
    Range("F3").FormulaR1C1 = "=RC[-4]" ' calculate level for hes
    Range("F4").Formula2R1C1 = "=hes_level(R4C12,RC[-4],R[-1]C,R[-1]C[1])"
    Range("F4").Select
    Selection.AutoFill Destination:=Range("F4:F" & nr)
    
    Range("G3").FormulaR1C1 = "=R[1]C[-5]-RC[-5]" ' calculate trend for hes
    Range("G4").Formula2R1C1 = "=hes_trend(R5C12,RC[-1],R[-1]C[-1],R[-1]C)"
    Range("G4").AutoFill Destination:=Range("G4:G" & nr)
    
    Range("H4").FormulaR1C1 = "=R[-1]C[-2]+R[-1]C[-1]" ' calculate forecast for hes
    Range("H4").AutoFill Destination:=Range("H4:H" & nr + 1)
    MsgBox ("Now, you will be asked for how many period forward you want to forecast. If you cancel or don't type anything, it would be automatically set to 3.")
    hes_n = Application.InputBox(Prompt:="Enter the (n) for Period Forecast.", Type:=1)
    If hes_n = 0 Then
        hes_n = 3
    End If
    Range("I" & nr + 1).FormulaR1C1 = "1" ' fill the t period
    Range("I" & nr + 1).AutoFill Range("I" & nr + 1 & ":I" & nr + hes_n), xlFillSeries
    hes_level_ref = Range("F" & nr).Value
    hes_trend_ref = Range("G" & nr).Value
    Range("H" & nr + 2).Formula2R1C1 = "=hes_forecast(" & hes_level_ref & "," & hes_trend_ref & ",RC[1])"
    Range("H" & nr + 2).AutoFill Destination:=Range("H" & nr + 2 & ":H" & nr + hes_n)

    Range("A" & nr).AutoFill Destination:=Range("A" & nr & ":A" & nr + hes_n), Type:=xlFillDefault ' autofill period
    
    ' ## Calculate for error
    
    With Range("N3") ' error for average forecast
        .FormulaR1C1 = "=ABS(RC[-11]-RC[-12])"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("N3:N" & nr), Type:=xlFillDefault
    End With
    With Range("O3")
        .FormulaR1C1 = "=RC[-1]^2"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("O3:O" & nr), Type:=xlFillDefault
    End With
    
    With Range("P" & ma_n) ' error for moving average n
        .Formula = "=ABS(D" & ma_n & "-B" & ma_n & ")"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("P" & ma_n & ":P" & nr), Type:=xlFillDefault
    End With
    With Range("Q" & ma_n)
        .FormulaR1C1 = "=RC[-1]^2"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("Q" & ma_n & ":Q" & nr), Type:=xlFillDefault
    End With
    
    With Range("R4")
        .FormulaR1C1 = "=ABS(RC[-13]-RC[-16])" ' error for naive forecast
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("R4:R" & nr), Type:=xlFillDefault
    End With
    With Range("S4")
        Range("S4").FormulaR1C1 = "=RC[-1]^2"
        Range("S4").NumberFormat = "0.000"
        Range("S4").AutoFill Destination:=Range("S4:S" & nr), Type:=xlFillDefault
    End With
    
    With Range("T4")
        .FormulaR1C1 = "=ABS(RC[-12]-RC[-18])" ' error for hes forecast
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("T4:T" & nr), Type:=xlFillDefault
    End With
    With Range("U4")
        .FormulaR1C1 = "=RC[-1]^2"
        .NumberFormat = "0.000"
        .AutoFill Destination:=Range("U4:U" & nr), Type:=xlFillDefault
    End With
    
    ' ## Calculate RMSE
    Range("L8").Formula = "=SQRT(AVERAGE(O3:O" & nr & "))" ' average forecast
    Range("L9").Formula = "=SQRT(AVERAGE(Q3:Q" & nr & "))" ' moving average
    Range("L10").Formula = "=SQRT(AVERAGE(S3:S" & nr & "))" ' naive forecast
    Range("L11").Formula = "=SQRT(AVERAGE(U3:U" & nr & "))" ' naive forecast
    
End Sub

Sub level_solver()
    SolverAdd CellRef:="$L$4", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$L$5", Relation:=1, FormulaText:="1"
    SolverAdd CellRef:="$L$4", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="LK$5", Relation:=3, FormulaText:="0"
    SolverOk SetCell:="$L$11", MaxMinVal:=2, ValueOf:=0, ByChange:="$L$4:$L$5", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve (True)
End Sub

Sub level_reset()
Dim nr As Integer

nr = Cells(Rows.Count, 8).End(xlUp).Row
If WorksheetFunction.CountA(Range("A3:B3")) = 0 Then
    Exit Sub
End If
Range("A3:I" & nr).ClearContents
Range("N3:U" & nr).ClearContents
Range("L8:L11").ClearContents
Range("L4").Value = 0.5
Range("L5").Value = 0.5
Range("D1:D2").FormulaR1C1 = "Moving Average (n)"

End Sub

Sub seasonal_reset()
    Dim nr As Integer

    nr = Cells(Rows.Count, 6).End(xlUp).Row
    If WorksheetFunction.CountA(Range("A3:B3")) = 0 Then
        Exit Sub
    End If
    Range("A3:G" & nr & ",L3:S" & nr & ",J10:J12").ClearContents
    With Range("J4")
        .FormulaR1C1 = "0.5"
        .AutoFill Destination:=Range("J4:J6"), Type:=xlFillDefault
    End With
    Range("J7").ClearContents

End Sub
Sub graph()
'
' graph Macro
'

'
    Sheets("leveltrend_all").Select
    Sheets("leveltrend_all").Move Before:=Sheets(5)
    Sheets("graph").Select
    ActiveChart.ChartArea.Select
    ActiveChart.ChartArea.Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.FullSeriesCollection(2).Select
End Sub

Sub level_graph()
Attribute level_graph.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim nr As Integer
    nr = Cells(Rows.Count, 8).End(xlUp).Row
    Sheets("graph").Select
    ActiveChart.ChartArea.Select
    With ActiveChart
        .SeriesCollection(1).Values = Worksheets("leveltrend_all").Range("B3:B" & nr)
        .SeriesCollection(2).Values = Worksheets("leveltrend_all").Range("H3:H" & nr)
        .Axes(xlCategory).CategoryNames = Worksheets("leveltrend_all").Range("A3:A" & nr)
        .ChartTitle.Text = "Forecast Using Holt's Exponential Smoothing Method"
    End With
End Sub

Sub seasonal_graph()
    Dim nr As Integer
    nr = Cells(Rows.Count, 6).End(xlUp).Row
    Sheets("graph").Select
    ActiveChart.ChartArea.Select
    With ActiveChart
        .SeriesCollection(1).Values = Worksheets("seasonality_wes").Range("B3:B" & nr)
        .SeriesCollection(2).Values = Worksheets("seasonality_wes").Range("F3:F" & nr)
        .Axes(xlCategory).CategoryNames = Worksheets("seasonality_wes").Range("A3:A" & nr)
        .ChartTitle.Text = "Forecast Using Winter's Exponential Smoothing"
    End With
End Sub
