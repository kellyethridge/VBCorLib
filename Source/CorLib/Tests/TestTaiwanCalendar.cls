VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTaiwanCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTaiwanCalendar
'
Option Explicit

Private cal As TaiwanCalendar

Public Sub Setup()
    Set cal = New TaiwanCalendar
End Sub

Public Sub TestGetYear()
    Assert.AreEqual 90, cal.GetYear(#1/1/2001#)
End Sub

Public Sub TestIsLeapYear()
    Assert.IsFalse cal.IsLeapYear(92)
    Assert.IsTrue cal.IsLeapYear(93)
End Sub

Public Sub TestToFourDigitYear()
    Assert.AreEqual 1, cal.ToFourDigitYear(1)
    Assert.AreEqual 8088, cal.ToFourDigitYear(8088)
End Sub

Public Sub TestToDateTime()
    Dim dt As cDateTime
    Set dt = Cor.NewcDateTime(2001, 1, 1, 1, 1, 1, 1)
    Assert.IsTrue dt.Equals(cal.ToDateTime(90, 1, 1, 1, 1, 1, 1))
End Sub

