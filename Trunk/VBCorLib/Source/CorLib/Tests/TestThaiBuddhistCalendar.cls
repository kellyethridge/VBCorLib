VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestThaiBuddhistCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestThaiBuddhistCalendar
'
Option Explicit

Private cal As ThaiBuddhistCalendar

Public Sub Setup()
    Set cal = New ThaiBuddhistCalendar
End Sub

Public Sub TestGetYear()
    Assert.AreEqual 2544, cal.GetYear(#1/1/2001#)
End Sub

Public Sub TestIsLeapYear()
    Assert.IsFalse cal.IsLeapYear(2542)
    Assert.IsTrue cal.IsLeapYear(2543)
End Sub

Public Sub TestToFourDigitYear()
    Assert.AreEqual 2500, cal.ToFourDigitYear(0)
    Assert.AreEqual 8000, cal.ToFourDigitYear(8000)
End Sub

Public Sub TestToDateTime()
    Dim dt As cDateTime
    Set dt = Cor.NewcDateTime(2001, 1, 1, 1, 1, 1, 1)
    Assert.IsTrue dt.Equals(cal.ToDateTime(2544, 1, 1, 1, 1, 1, 1))
End Sub
