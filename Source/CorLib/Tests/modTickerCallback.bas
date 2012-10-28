Attribute VB_Name = "modTickerCallback"
'
' modTickerCallback
'
Option Explicit

Public Sub TickerEvent(ByRef t As Ticker, ByRef Data As Variant)
    Debug.Print "Callback: " & Data, "Interval: " & t.Interval
End Sub

