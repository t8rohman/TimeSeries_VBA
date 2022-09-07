Attribute VB_Name = "Module1"
Option Explicit
' This module contains the coding for the Holt's and Winter's formula

Public Function hes_level(alpha As Double, y As Double, levelt_1 As Double, trendt_1 As Double)
hes_level = (alpha * y) + (1 - alpha) * (levelt_1 + trendt_1)
End Function

Public Function hes_trend(beta As Double, level_now As Double, levelt_1 As Double, trendt_1 As Double)
hes_trend = beta * (level_now - levelt_1) + (1 - beta) * trendt_1
End Function

Public Function hes_forecast(level_now As Double, trend_now As Double, period As Integer)
hes_forecast = level_now + period * trend_now
End Function

Public Function wes_level(alpha As Double, y_now As Double, levelt_1 As Double, _
trendt_1 As Double, season_index As Double)
wes_level = alpha * (y_now / season_index) + (1 - alpha) * (levelt_1 + trendt_1)
End Function

Public Function wes_trend(beta As Double, level_now As Double, levelt_1 As Double, trendt_1 As Double)
wes_trend = beta * (level_now - levelt_1) + (1 - beta) * trendt_1
End Function

Public Function wes_season(gamma As Double, y_now As Double, level_now As Double, season_index As Double)
wes_season = gamma * (y_now / level_now) + (1 - gamma) * season_index
End Function

Public Function wes_forecast(level_now As Double, trend_now As Double, season_index As Double, period As Double)
wes_forecast = (level_now + period * trend_now) * season_index
End Function

