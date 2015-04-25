VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestKoreanCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestKoreanCalendar
'
Option Explicit
Implements ICategorizable

Private cal As KoreanCalendar

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Public Sub Setup()
    Set cal = New KoreanCalendar
End Sub

Public Sub TestGetYear()
    Assert.AreEqual 4334, cal.GetYear(#1/1/2001#)
End Sub

Public Sub TestIsLeapYear()
    Assert.IsFalse cal.IsLeapYear(2336)
    Assert.IsTrue cal.IsLeapYear(2337)
End Sub

Public Sub TestToFourDigitYear()
    Assert.AreEqual 4300, cal.ToFourDigitYear(0)
    Assert.AreEqual 8000, cal.ToFourDigitYear(8000)
End Sub

Public Sub TestToDateTime()
    Dim dt As cDateTime
    Set dt = Cor.NewDateTime(2001, 1, 1, 1, 1, 1, 1)
    Assert.IsTrue dt.Equals(cal.ToDateTime(4334, 1, 1, 1, 1, 1, 1))
End Sub

