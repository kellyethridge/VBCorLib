Attribute VB_Name = "modCalendarHelpers"
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modCalendarHelpers
'

''
' Provides some common functions among the different calendars.
'
Option Explicit


''
' Many of the calendars calculate the week of the year the same way,
' so wrap up a function here to let them all share.
'
' @param Time The Date to calculate the week of the year from.
' @param Rule How to determine what the first week of the year is.
' @param FirstDayOfWeek Which day of the week is the start of the week.
' @param Cal The Calendar to be used for specific calculations.
' @return The week of the year.
'
Public Function InternalGetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek, ByVal Cal As Calendar) As Long
    Dim FirstWeekLength As Long
    Dim Offset          As Long
    Dim dt              As CorDateTime
    Dim WholeWeeks      As Long
    Dim doy             As Long

    Set dt = CorDateTime.GetcDateTime(Time)

    FirstWeekLength = FirstDayOfWeek - Cal.GetDayOfWeek(Cal.ToDateTime(Cal.GetYear(dt), 1, 1, 0, 0, 0, 0))
    If FirstWeekLength < 0 Then FirstWeekLength = FirstWeekLength + 7

    Select Case Rule
        Case FirstDay:              If FirstWeekLength > 0 Then Offset = 1
        Case FirstFullWeek:         If FirstWeekLength >= 7 Then Offset = 1
        Case FirstFourDayWeek:      If FirstWeekLength >= 4 Then Offset = 1
    End Select

    doy = Cal.GetDayOfYear(dt)
    If doy > FirstWeekLength Then
        WholeWeeks = (doy - FirstWeekLength) \ 7
        If WholeWeeks * 7 + FirstWeekLength < doy Then Offset = Offset + 1
    End If

    InternalGetWeekOfYear = WholeWeeks + Offset
    If InternalGetWeekOfYear = 0 Then
        Dim Year    As Long
        Dim Month   As Long
        
        Year = Cal.GetYear(dt) - 1
        Month = Cal.GetMonthsInYear(Year)
        InternalGetWeekOfYear = InternalGetWeekOfYear(Cal.ToDateTime(Year, Month, Cal.GetDaysInMonth(Year, Month), 0, 0, 0, 0), Rule, FirstDayOfWeek, Cal)
    End If
End Function

''
' Gets a numeric value from the system calendar settings.
'
' @param Cal The calendar to get the value from.
' @param CalType The type of value to get from the calendar.
' @return The numeric value from the calendar on the system.
'
Public Function GetCalendarLong(ByVal Cal As Long, ByVal CalType As Long) As Long
    If GetCalendarInfo(LOCALE_USER_DEFAULT, Cal, CalType Or CAL_RETURN_NUMBER, vbNullString, 0, GetCalendarLong) = BOOL_FALSE Then IOError Err.LastDllError
End Function

