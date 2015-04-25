VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestJapaneseCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestJapaneseCalendar
'
Option Explicit
Implements ICategorizable

Private cal As JapaneseCalendar

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Public Sub Setup()
    Set cal = New JapaneseCalendar
End Sub


Public Sub TestAddDays()
    Dim d1 As cDateTime
    Dim d2 As cDateTime
    Dim i As Long
    
    Set d1 = cDateTime.FromOADate(#1/1/2000#)
    Set d2 = d1
    
    For i = 1 To 1000
        Set d1 = d1.AddDays(1)
        Set d2 = cal.AddDays(d2, 1)
        Assert.IsTrue d2.Equals(d1)
    Next i
End Sub

Public Sub TestAddMonths()
    Dim d1 As cDateTime
    Dim d2 As cDateTime
    Dim i As Long
    
    Set d1 = cDateTime.FromOADate(#1/1/2000#)
    Set d2 = d1
    
    For i = 1 To 1000
        Set d1 = d1.AddMonths(1)
        Set d2 = cal.AddMonths(d2, 1)
        Assert.IsTrue d2.Equals(d1)
    Next i
End Sub

Public Sub TestAddYears()
    Dim d1 As cDateTime
    Dim d2 As cDateTime
    Dim i As Long
    
    Set d1 = cDateTime.FromOADate(#1/1/2000#)
    Set d2 = d1
    
    For i = 1 To 1000
        Set d1 = d1.AddYears(1)
        Set d2 = cal.AddYears(d2, 1)
        Assert.IsTrue d2.Equals(d1)
    Next i
End Sub

Public Sub TestAddWeeks()
    Dim d1 As cDateTime
    Dim d2 As cDateTime
    Dim i As Long
    
    Set d1 = cDateTime.FromOADate(#1/1/2000#)
    Set d2 = d1
    
    For i = 1 To 1000
        Set d1 = d1.AddDays(7)
        Set d2 = cal.AddWeeks(d2, 1)
        Assert.IsTrue d2.Equals(d1)
    Next i
End Sub

Public Sub TestGetYear()
    Dim sr As StreamReader
    Dim d As cDateTime
    Dim i As Long
    
    Assert.AreEqual 1, cal.GetYear(#9/8/1868#)
    Assert.AreEqual 1, cal.GetYear(#7/30/1912#)
    Assert.AreEqual 1, cal.GetYear(#12/25/1926#)
    Assert.AreEqual 1, cal.GetYear(#1/8/1989#)
    Assert.AreEqual 1, cal.GetYear(#12/31/1989#)
    Assert.AreEqual 2, cal.GetYear(#1/1/1990#)
    Assert.AreEqual 2, cal.GetYear(#12/31/1990#)
    Assert.AreEqual 3, cal.GetYear(#1/1/1991#)
    
    Set sr = File.OpenText(App.Path & "\Dummy Test Files\TestJapaneseCalendar_GetYear.txt")
    
    Set d = cDateTime.FromOADate(#1/1/1900#)
    For i = 1 To 100
        Assert.AreEqual CLng(sr.ReadLine), cal.GetYear(d)
        Set d = d.AddYears(1)
    Next i
    sr.CloseReader
End Sub

Public Sub TestIsLeapYear()
    Dim d As Date
    Dim i As Long
    
    Assert.IsFalse cal.IsLeapYear(cal.GetYear(#1/1/1900#), cal.GetEra(#1/1/1900#)), "Year 1900 should return false."
    
    d = #1/1/1901#
    For i = 1 To 100
        If Year(d) Mod 4 <> 0 Then
            Assert.IsFalse cal.IsLeapYear(cal.GetYear(d), cal.GetEra(d)), "Year " & Year(d) & " should be false."
        Else
            Assert.IsTrue cal.IsLeapYear(cal.GetYear(d), cal.GetEra(d)), "Year " & Year(d) & " should be true."
        End If
        d = DateAdd("yyyy", 1, d)
    Next i
End Sub

Public Sub TestToDateTimeForCurrentEra()
    Dim dt As cDateTime
    Set dt = cal.ToDateTime(1, 1, 1, 0, 0, 0, 0)
    Assert.AreEqual #1/1/1989#, dt.ToOADate
End Sub

Public Sub TestToDateTimeForHeiseiEra()
    Dim dt As cDateTime
    Set dt = cal.ToDateTime(1, 1, 1, 0, 0, 0, 0, 4)
    Assert.AreEqual #1/1/1989#, dt.ToOADate
End Sub

Public Sub TestToDateTimeForShowaEra()
    Dim dt As cDateTime
    Set dt = cal.ToDateTime(1, 1, 1, 0, 0, 0, 0, 3)
    Assert.AreEqual #1/1/1926#, dt.ToOADate
    
    Set dt = cal.ToDateTime(64, 1, 1, 0, 0, 0, 0, 3)
    Assert.AreEqual #1/1/1989#, dt.ToOADate
End Sub

Public Sub TestToDateTimeForTaishoEra()
    Dim dt As cDateTime
    Set dt = cal.ToDateTime(1, 1, 1, 0, 0, 0, 0, 2)
    Assert.AreEqual #1/1/1912#, dt.ToOADate
    
    Set dt = cal.ToDateTime(15, 1, 1, 0, 0, 0, 0, 2)
    Assert.AreEqual #1/1/1926#, dt.ToOADate
End Sub

Public Sub TestToDateTimeForMeijiEra()
    Dim dt As cDateTime
    Set dt = cal.ToDateTime(1, 1, 1, 0, 0, 0, 0, 1)
    Assert.AreEqual #1/1/1868#, dt.ToOADate
    
    Set dt = cal.ToDateTime(45, 1, 1, 0, 0, 0, 0, 1)
    Assert.AreEqual #1/1/1912#, dt.ToOADate
End Sub













