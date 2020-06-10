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
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: TimeZone
'

''
' Represents a time zone.
'
' @see TimeZoneStatic
'
Option Explicit
Implements IObject

Private Const FIND_LAST As Long = 5

Private mDaylightName           As String
Private mStandardName           As String
Private mUtcOffsetNoDLS         As TimeSpan
Private mUtcOffsetWithDLS       As TimeSpan
Private mTimeZoneInfo           As TIME_ZONE_INFORMATION
Private mNotUsedDayLightTimes   As DayLightTime


''
' Returns the name of the daylight saving period.
'
' @return Daylight saving period name.
Public Property Get DaylightName() As String
    DaylightName = mDaylightName
End Property

''
' Returns the name of the standard time period.
'
' @return Standard time period name.
Public Property Get StandardName() As String
    StandardName = mStandardName
End Property

''
' Returns the start and end of the daylight saving period.
'
' @param Year The year which contains the daylight saving period.
' @return The daylight saving period.
Public Function GetDayLightChanges(ByVal Year As Long) As DayLightTime
    Set GetDayLightChanges = GetDaylightTimes(Year)
End Function

''
' Returns the offset from UTC to the current timezone, adjusting for daylight saving.
'
' @param time The time to get the offset for.
' @return The offset from the time to UTC.
Public Function GetUtcOffset(ByRef Time As Variant) As TimeSpan
    If IsDaylightSavingTime(Time) Then
        Set GetUtcOffset = mUtcOffsetWithDLS
    Else
        Set GetUtcOffset = mUtcOffsetNoDLS
    End If
End Function

''
' Returns whether the time is currently in the daylight saving period.
'
' @param Time The time to check if in the daylight saving period.
' @return Indicating if the time was in the daylight saving period.
Public Function IsDaylightSavingTime(ByRef Time As Variant) As Boolean
    Dim dt As CorDateTime

    Set dt = Convert.ToDateTime(Time)

    IsDaylightSavingTime = TimeZone.IsDaylightSavingTime(dt, GetDaylightTimes(dt.Year))
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
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim r As Long
    
    r = GetTimeZoneInformation(mTimeZoneInfo)
    
    If r = TIME_ZONE_ID_UNKNOWN Or mTimeZoneInfo.DaylightBias = 0 Then
        Set mNotUsedDayLightTimes = Cor.NewDayLightTime(CorDateTime.MinValue, CorDateTime.MinValue, TimeSpan.Zero)
    End If
    
    mDaylightName = SysAllocString(VarPtr(mTimeZoneInfo.DaylightName(0)))
    mStandardName = SysAllocString(VarPtr(mTimeZoneInfo.StandardName(0)))
    Set mUtcOffsetNoDLS = TimeSpan.FromMinutes(-(mTimeZoneInfo.Bias + mTimeZoneInfo.StandardBias))
    Set mUtcOffsetWithDLS = TimeSpan.FromMinutes(-(mTimeZoneInfo.Bias + mTimeZoneInfo.DaylightBias))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDaylightTimes(ByVal Year As Long) As DayLightTime
    Dim StartTime As CorDateTime
    Dim EndTime As CorDateTime
    
    If mNotUsedDayLightTimes Is Nothing Then
        Set StartTime = GetCrossOverPoint(Year, mTimeZoneInfo.DaylightDate)
        Set EndTime = GetCrossOverPoint(Year, mTimeZoneInfo.StandardDate)
        Set GetDaylightTimes = Cor.NewDayLightTime(StartTime, EndTime, TimeSpan.FromMinutes(mTimeZoneInfo.Bias))
    Else
        Set GetDaylightTimes = mNotUsedDayLightTimes
    End If
End Function

Private Function GetCrossOverPoint(ByVal Year As Long, ByRef Time As SYSTEMTIME) As CorDateTime
    Dim Ret         As CorDateTime
    Dim MyDow       As DayOfWeek
    Dim TargetDow   As DayOfWeek
    
    If Time.wDay = FIND_LAST Then
        ' find last dayofweek in month
        Set Ret = Cor.NewDateTime(Year, Time.wMonth + 1, 1, Time.wHour)
        MyDow = Ret.DayOfWeek
        TargetDow = MyDow - Time.wDayOfWeek
        
        If TargetDow < 0 Then
            TargetDow = TargetDow + 7
        End If
        
        Set Ret = Ret.AddDays(-TargetDow)
    Else
        ' find nth dayofweek of month
        Set Ret = Cor.NewDateTime(Year, Time.wMonth, 1, Time.wHour)
        MyDow = Ret.DayOfWeek
        TargetDow = Time.wDayOfWeek
        
        If TargetDow < MyDow Then
            TargetDow = TargetDow + 7
        End If
        
        Set Ret = Ret.AddDays(TargetDow - MyDow)
        
        If Time.wDay > 1 Then
            Set Ret = Ret.AddDays((Time.wDay - 1) * 7)
        End If
    End If
    
    Set GetCrossOverPoint = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
