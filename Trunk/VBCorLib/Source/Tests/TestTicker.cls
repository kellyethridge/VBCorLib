VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTicker
'
Option Explicit

'Private WithEvents mTicker As Ticker
'
'
'Public Sub Setup()
'    Set mTicker = NewTicker(1000, Now, , AddressOf TickerEvent)
'End Sub
'
'Public Sub TestEnabled()
'    mTicker.Enabled = True
'End Sub
'
'Private Sub mTicker_Elapsed(ByRef Data As Variant)
'    Debug.Print "Event: " & Data
'End Sub
