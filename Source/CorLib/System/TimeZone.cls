VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: TimeZone
'

''
' Represents timezone information for the local machine.
'
' @see TimeZoneStatic
'
Option Explicit
Implements IObject

Private Const FIND_LAST As Long = 5

Private mDayLightName           As String
Private mStandardName           As String
Private mUtcOffsetNoDLS         As TimeSpan
Private mUtcOffsetWithDLS       As TimeSpan
Private mTimeZoneInfo           As TIME_ZONE_INFORMATION
Private mNotUsedDayLightTimes   As DayLightTime


''
' Returns the name of the daylight savings period.
'
' @return Daylight savings period name.
Public Property Get DayLightName() As String
    DayLightName = mDayLightName
End Property

''
' Returns the name of the standard time period.
'
' @return Standard time period name.
Public Property Get StandardName() As String
    StandardName = mStandardName
End Property

''
' Returns the start and end of the daylight savings period.
'
' @param Year The year which contains the daylight savings period.
' @return The daylight savings period.
Public Function GetDayLightChanges(ByVal Year As Long) As DayLightTime
    Set GetDayLightChanges = GetDayLightTimes(Year)
End Function

''
' Returns the offset from UTC to the current timezone, adjusting for daylight savings.
'
' @param time The time to get the offset for.
' @return The offset from the time to UTC.
Public Function GetUtcOffset(ByRef Time As Variant) As TimeSpan
    If IsDayLightSavingTime(Time) Then
        Set GetUtcOffset = mUtcOffsetWithDLS
    Else
        Set GetUtcOffset = mUtcOffsetNoDLS
    End If
End Function

''
' Returns whether the time is currently in the daylight savings period.
'
' @param time The time to check if in the daylight savings period.
' @param dayLightTimes the daylight period to check the time against.
' @return Indicating if the time was in the daylight savings period.
Public Function IsDayLightSavingTime(ByRef Time As Variant, Optional ByVal dayLightTimes As DayLightTime) As Boolean
    Dim dt As CorDateTime
    
    Set dt = Convert.ToDateTime(Time)
    If dayLightTimes Is Nothing Then Set dayLightTimes = GetDayLightTimes(dt.Year)
    
    With dayLightTimes
        If .Delta.Equals(TimeSpan.Zero) Then Exit Function
        If .StartTime.LessThan(.EndTime) Then
            ' northern hemisphere
            If dt.LessThan(.StartTime) Then Exit Function
            If dt.GreaterThanOrEqualTo(.EndTime) Then Exit Function
        Else
            ' southern hemisphere
            If dt.GreaterThanOrEqualTo(.StartTime) Then Exit Function
            If dt.LessThan(.EndTime) Then Exit Function
        End If
    End With
    IsDayLightSavingTime = True
End Function

''
' Returns the time with the UTC offset applied to get the local time.
'
' @param Time The UTC time to convert to local time.
' @return The local time converted from the UTC time.
Public Function ToLocalTime(ByRef Time As Variant) As CorDateTime
    Dim d As CorDateTime
    Set d = Convert.ToDateTime(Time)
    
    If d.Kind = LocalKind Then
        Set ToLocalTime = d
    Else
        Set ToLocalTime = CorDateTime.SpecifyKind(d.Add(GetUtcOffset(d)), LocalKind)
    End If
End Function

''
' Returns the local time with UTC offset removed to get the UTC time.
'
' @param Time The local time to have the UTC offset removed.
' @return The UTC time converted from the local time.
Public Function ToUniversalTime(ByRef Time As Variant) As CorDateTime
    Dim d As CorDateTime
    Set d = Convert.ToDateTime(Time)
    
    If d.Kind = UtcKind Then
        Set ToUniversalTime = d
    Else
        Set ToUniversalTime = CorDateTime.SpecifyKind(d.Subtract(GetUtcOffset(d)), UtcKind)
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim R As Long
    
    R = GetTimeZoneInformation(mTimeZoneInfo)
    If R = TIME_ZONE_ID_UNKNOWN Or mTimeZoneInfo.DaylightBias = 0 Then
        Set mNotUsedDayLightTimes = Cor.NewDayLightTime(CorDateTime.MinValue, CorDateTime.MinValue, TimeSpan.Zero)
    End If
    
    mDayLightName = SysAllocString(VarPtr(mTimeZoneInfo.DayLightName(0)))
    mStandardName = SysAllocString(VarPtr(mTimeZoneInfo.StandardName(0)))
    Set mUtcOffsetNoDLS = TimeSpan.FromMinutes(-(mTimeZoneInfo.Bias + mTimeZoneInfo.StandardBias))
    Set mUtcOffsetWithDLS = TimeSpan.FromMinutes(-(mTimeZoneInfo.Bias + mTimeZoneInfo.DaylightBias))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDayLightTimes(ByVal Year As Long) As DayLightTime
    Dim StartTime As CorDateTime
    Dim EndTime As CorDateTime
    
    If mNotUsedDayLightTimes Is Nothing Then
        Set StartTime = GetCrossOverPoint(Year, mTimeZoneInfo.DaylightDate)
        Set EndTime = GetCrossOverPoint(Year, mTimeZoneInfo.StandardDate)
        Set GetDayLightTimes = Cor.NewDayLightTime(StartTime, EndTime, TimeSpan.FromMinutes(mTimeZoneInfo.Bias))
    Else
        Set GetDayLightTimes = mNotUsedDayLightTimes
    End If
End Function

Private Function GetCrossOverPoint(ByVal Year As Long, ByRef Time As SYSTEMTIME) As CorDateTime
    Dim Ret As CorDateTime
    Dim mydow As DayOfWeek
    Dim targetdow As DayOfWeek
    
    If Time.wDay = FIND_LAST Then
        ' find last dayofweek in month
        Set Ret = Cor.NewDateTime(Year, Time.wMonth + 1, 1, Time.wHour)
        mydow = Ret.DayOfWeek
        targetdow = mydow - Time.wDayOfWeek
        If targetdow < 0 Then targetdow = targetdow + 7
        Set Ret = Ret.AddDays(-targetdow)
    Else
        ' find nth dayofweek of month
        Set Ret = Cor.NewDateTime(Year, Time.wMonth, 1, Time.wHour)
        mydow = Ret.DayOfWeek
        targetdow = Time.wDayOfWeek
        If targetdow < mydow Then targetdow = targetdow + 7
        Set Ret = Ret.AddDays(targetdow - mydow)
        If Time.wDay > 1 Then
            Set Ret = Ret.AddDays((Time.wDay - 1) * 7)
        End If
    End If
    
    Set GetCrossOverPoint = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
