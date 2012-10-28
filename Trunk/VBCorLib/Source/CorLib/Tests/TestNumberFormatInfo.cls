VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestNumberFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestNumberFormatInfo
'
Option Explicit

Public Sub TestGeneralInteger()
    Dim info As New NumberFormatInfo
    Assert.AreEqual "12345", info.Format(12345)
End Sub

Public Sub TestGeneralIntegerScaled()
    Dim info As New NumberFormatInfo
    Assert.AreEqual "1.23e+04", info.Format(12345, "g3")
End Sub
