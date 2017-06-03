VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GregorianCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Public Sub TwoDigitYearMax_With()
    Dim i As Long
    Dim cal As New GregorianCalendar
    
    i = cal.TwoDigitYearMax
    Assert.IsTrue i >= 100 And i <= 9999, "i should be between 100 and 9999."
End Sub

Public Sub TestGetWeekOfYear()
    Dim cal As New GregorianCalendar
    Dim sr As StreamReader
    Dim i As Long
    Dim dt As CorDateTime
    Dim j As Long
    
    Set sr = NewStreamReader(App.Path & "\Dummy Test Files\TestGregorianCalendar_WeekOfYear.txt")
    Set dt = CorDateTime.FromOADate(#1/1/2004#)
    
    ' tests using CalendarWeekRule.FirstDay and first day of week of DayOfWeek.Saturday
    For i = 1 To 366
        j = sr.ReadLine
        Assert.AreEqual j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Saturday), "Wrong week for " & dt.ToString & " during first test."
        Set dt = dt.AddDays(1)
    Next i
    
    ' tests using CalendarWeekRule.FirstFourDayWeek and first day of week of DayOfWeek.Wednesday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.AreEqual j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Wednesday), "Wrong week for " & dt.ToString & " during second test."
        Set dt = dt.AddDays(1)
    Next i
    
    ' tests using CalendarWeekRule.FirstFullWeek and first day of week of DayOfWeek.Monday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.AreEqual j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday), "Wrong week for " & dt.ToString & " during thrid test."
        Set dt = dt.AddDays(1)
    Next i
    
    sr.CloseReader
End Sub

Public Sub TestToFourDigitYear()
    Dim i As Long
    Dim j As Long
    Dim cal As New GregorianCalendar
    
    j = 1930
    
    For i = 30 To 129
        Assert.AreEqual j, cal.ToFourDigitYear(i Mod 100)
        j = j + 1
    Next i
End Sub
