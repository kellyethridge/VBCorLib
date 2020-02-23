VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZoneStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: TimeZoneStatic
'

''
' Represents static methods for the <a href="TimeZone.html">TimeZone</a> class.
'
' @see TimeZone
'
Option Explicit

Private mCurrentTimeZone As TimeZone


''
' Returns the current time zone for this machine.
'
' @return The current time zone for this machine.
Public Property Get CurrentTimeZone() As TimeZone
    Set CurrentTimeZone = mCurrentTimeZone
End Property

''
' Returns whether the time is currently in the daylight saving period.
'
' @param Time The time to check if in the daylight saving period.
' @param daylightTimes the daylight period to check the time against.
' @return <b>True</b> if <i>Time</i> is in the daylight saving period; <b>False</b> otherwise.
Public Function IsDaylightSavingTime(ByRef Time As Variant, ByVal daylightTimes As DayLightTime) As Boolean
    Dim dt As CorDateTime
    
    If daylightTimes Is Nothing Then _
        Error.ArgumentNull "daylightTimes"
    
    Set dt = Convert.ToDateTime(Time)
    
    With daylightTimes
        If .Delta.Equals(TimeSpan.Zero) Then
            Exit Function
        End If
        
        If .StartTime.LessThan(.EndTime) Then
            ' northern hemisphere
            If dt.LessThan(.StartTime) Then
                Exit Function
            End If
            
            If dt.GreaterThanOrEqualTo(.EndTime) Then
                Exit Function
            End If
        Else
            ' southern hemisphere
            If dt.GreaterThanOrEqualTo(.StartTime) Then
                Exit Function
            End If
            
            If dt.LessThan(.EndTime) Then
                Exit Function
            End If
        End If
    End With
    
    IsDaylightSavingTime = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mCurrentTimeZone = New TimeZone
End Sub
