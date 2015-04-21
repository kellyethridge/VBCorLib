VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHebrewCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHebrewCalendar
'
Option Explicit

Private cal As HebrewCalendar

Public Sub Setup()
    Set cal = New HebrewCalendar
End Sub

Public Sub TestIsLeapYear()
    Dim sr As StreamReader
    Dim i As Long
    
    Set sr = File.OpenText(App.Path & "\dummy test files\HebrewCalendar_IsLeapYear.txt")
    
    For i = 1 To 100
        Assert.AreEqual CBool(sr.ReadLine), cal.IsLeapYear(5400 + i), "Year should be true."
    Next i
    sr.CloseReader
End Sub

Public Sub TestGetDaysInYear()
    Assert.AreEqual 355, cal.GetDaysInYear(5500)
    Assert.AreEqual 383, cal.GetDaysInYear(5502)
    Assert.AreEqual 353, cal.GetDaysInYear(5550)
End Sub

Public Sub TestGetYear()
    Dim sr As StreamReader
    Dim d As cDateTime
    Dim i As Long
    
    Set sr = File.OpenText(App.Path & "\dummy test files\hebrewcalendar_getyear.txt")
    Set d = cDateTime.FromOADate(#1/1/1950#)
    
    For i = 1 To 100
        Assert.AreEqual CLng(sr.ReadLine), cal.GetYear(d)
        Set d = d.AddYears(1)
    Next i
    
    sr.CloseReader
    
End Sub

Public Sub TestGetDaysInMonth()
    Dim i As Long
    Dim ly As Long
    
    For i = 1 To 12
        Assert.AreEqual IIf((i And 1) Or (i = 2), 30, 29), cal.GetDaysInMonth(5500, i), "month " & i & " should be 30."
    Next i
    
    Assert.AreEqual 30, cal.GetDaysInMonth(5502, 1), "1/5502"
    Assert.AreEqual 29, cal.GetDaysInMonth(5502, 2), "2/5502"
    Assert.AreEqual 29, cal.GetDaysInMonth(5502, 3), "3/5502"
    Assert.AreEqual 29, cal.GetDaysInMonth(5502, 4), "4/5502"
    Assert.AreEqual 30, cal.GetDaysInMonth(5502, 5), "5/5502"
    
    For i = 6 To 13
        Assert.AreEqual IIf(i And 1, 29, 30), cal.GetDaysInMonth(5502, i), i & "/5502"
    Next i
End Sub

Public Sub TestGetMonth()
    Dim sr As StreamReader
    Dim i As Long
    Dim d As cDateTime
    
    Set sr = File.OpenText(App.Path & "\dummy test files\HebrewCalendar_GetMonth.txt")
    Set d = cDateTime.FromOADate(#1/1/2000#)
    
    For i = 1 To 100
        Assert.AreEqual CLng(sr.ReadLine), cal.GetMonth(d), "Incorrect month for #" & d.ToShortDateString & "#"
        Set d = d.AddMonths(1)
    Next i
    
    sr.CloseReader
End Sub

Public Sub TestToDateTime()
    Assert.AreEqual #5/4/2000 1:01:01 AM#, cal.ToDateTime(5760, 8, 29, 1, 1, 1, 0).ToOADate
    Assert.AreEqual #12/10/1739 1:01:01 AM#, cal.ToDateTime(5500, 3, 9, 1, 1, 1, 0).ToOADate
End Sub

Public Sub TestGetDayOfMonth()
    Assert.AreEqual 23, cal.GetDayOfMonth(#1/1/2000#), "Day of month should be 23."
    Assert.AreEqual 21, cal.GetDayOfMonth(#8/29/2002#), "Day of month should be 21."
End Sub

Public Sub TestGetDayOfWeek()
    Assert.AreEqual 6, cal.GetDayOfWeek(#1/1/2000#)
    Assert.AreEqual 1, cal.GetDayOfWeek(#1/1/2001#)
End Sub

Public Sub TestGetDayOfYear()
    Assert.AreEqual 94, cal.GetDayOfYear(#1/1/2001#)
    Assert.AreEqual 298, cal.GetDayOfYear(#7/1/2003#)
End Sub

Public Sub TestAddYears()
    Assert.AreEqual #12/28/2002 10:35:35 PM#, cal.AddYears(#1/1/2000 10:35:35 PM#, 3).ToOADate
    Assert.AreEqual #1/10/2067#, cal.AddYears(#1/1/2000#, 67).ToOADate
End Sub

Public Sub TestAddMonths()
    Assert.AreEqual #9/23/2000#, cal.AddMonths(#1/1/2000#, 9).ToOADate
    Assert.AreEqual #1/10/2047#, cal.AddMonths(#1/1/1950#, 1200).ToOADate
    Assert.AreEqual #10/30/1983#, cal.AddMonths(#1/1/2000#, -200).ToOADate
End Sub
