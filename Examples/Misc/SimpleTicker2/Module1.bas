Attribute VB_Name = "Module1"
Option Explicit

' The Ticker will call back into this function when
' the Interval has elapsed.
'
' The signature should be followed closely.
Public Sub TickerCallback(ByRef Ticker As Ticker, ByRef Data As Variant)
    ' We were smart enough to allow the Ticker to carry
    ' some data around for easy access during callbacks.
    Dim c As Counter
    Set c = Data
    c.Count = c.Count + 1
End Sub

