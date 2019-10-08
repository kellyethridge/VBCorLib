VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: TimeSpan
'

''
' This class represents an amount of time.
'
' @remarks The time is represented internally in milliseconds, with a fractional portion
' to represent ticks. A tick is defined as 100 nanoseconds, or 1/10000 of a millisecond.
'
' TimeSpans are displayed in a [-][d.]hh:mm:ss[.ff] format. Bracketed fields are only
' displayed if they are not zero. A negative sign is displayed if the timespan is negative.
'
' @see Constructors
' @see TimeSpanStatic
' @see CorDateTime
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable

Private Const PropMilliseconds      As String = "Milliseconds"
Private Const DefaultMilliseconds   As Long = 0

Private mMilliseconds As Currency


''
' Adds a timespan to this instance.
'
' @param ts TimeSpan to be added.
' @return A TimeSpan that represents the sum of the two timespans.
Public Function Add(ByRef ts As TimeSpan) As TimeSpan
    Set Add = TimeSpan.Add(Me, ts)
End Function

''
' Subtracts a timespan from this instance.
'
' @param ts TimeSpan to be subtracted.
' @return TimeSpan with the subtracted value.
Public Function Subtract(ByRef ts As TimeSpan) As TimeSpan
    Set Subtract = TimeSpan.Subtract(Me, ts)
End Function

''
' Returns the absolute value of this instance.
'
' @return TimeSpan that represents the absolute value of the original TimeSpan.
Public Function Duration() As TimeSpan
    If mMilliseconds < 0@ Then
        Set Duration = TimeSpan.FromMilliseconds(-mMilliseconds)
    Else
        Set Duration = Me
    End If
End Function

''
' Returns the duration of this instance in milliseconds.
'
' @return This timespan in milliseconds.
' @remarks The returned duration can contain a fractional
' part if necessary to represent the duration.
Public Property Get TotalMilliseconds() As Currency
    TotalMilliseconds = mMilliseconds
End Property

''
' Returns the duration of this instance in seconds.
'
' @return This timespan in seconds.
' @remarks The returned duration can contain a fractional
' part if necessary to represent the duration.
Public Property Get TotalSeconds() As Double
    TotalSeconds = mMilliseconds / MillisecondsPerSecond
End Property

''
' Returns the duration of this instance in minutes.
'
' @return This timespan in minutes.
' @remarks The returned duration can contain a fractional
' part if necessary to represent the duration.
Public Property Get TotalMinutes() As Double
    TotalMinutes = mMilliseconds / MillisecondsPerMinute
End Property

''
' Returns the duration of this instance in hours.
'
' @return This timespan in hours.
Public Property Get TotalHours() As Double
    TotalHours = mMilliseconds / MillisecondsPerHour
End Property

''
' Returns the duration of this instance in days.
'
' @return This timespan in days.
' @remarks The returned duration can contain a fractional
' part if necessary to represent the duration.
Public Property Get TotalDays() As Double
    TotalDays = mMilliseconds / MilliSecondsPerDay
End Property

''
' Returns the Milliseconds portion of the duration.
'
' @return The Milliseconds portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Milliseconds() As Long
    Milliseconds = AsLong(mMilliseconds * MillisecondsPerTick) Mod 1000
End Property

''
' Returns the Seconds portion of the duration.
'
' @return The Seconds portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Seconds() As Long
    Seconds = Fix(mMilliseconds / MillisecondsPerSecond) Mod 60
End Property

''
' Returns the Minutes portion of the duration.
'
' @return The Minutes portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Minutes() As Long
    Minutes = Fix(mMilliseconds / MillisecondsPerMinute) Mod 60
End Property

''
' Returns the Hours portion of the duration.
'
' @return The Hours portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Hours() As Long
    Hours = Fix(mMilliseconds / MillisecondsPerHour) Mod 24
End Property

''
' Returns the Days portion of the duration.
'
' @return The Days portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Days() As Long
    Days = Fix(mMilliseconds / MilliSecondsPerDay)
End Property

''
' Returns the duration in ticks.
'
' @return The duration in ticks
' @remarks A tick is defined as 100 nanoseconds, or 1/10000 of a millisecond.
Public Property Get Ticks() As Variant
    Ticks = CDec(mMilliseconds) * TicksPerMillisecond
End Property

''
' Compares this TimeSpan with another timespan.
'
' @param value The TimeSpan to compare this instance against.
' @return Value indicating the relation between the two timespans.
' @remarks The return value indicates how the two timespans are
' in relation to eachother.<br>
' 1 = This instance is greater than the passed in timespan.<br>
' 0 = This instance is equal to the passed in timespan.<br>
' -1 = This instance is less than the passed in timespan.<br>
' <br><br>
' If <b>Nothing</b> is passed in 1 is returned.
Public Function CompareTo(ByRef Value As TimeSpan) As Long
    CompareTo = TimeSpan.Compare(Me, Value)
End Function

''
' Returns a TimeSpan with the opposite sign of this instance.
'
' @return The timespan with the opposite sign as this duration.
Public Function Negate() As TimeSpan
    Set Negate = TimeSpan.FromMilliseconds(-mMilliseconds)
End Function

''
' Converts the value of the current <a href="TimeSpan.html">TimeSpan</a> object to its equivalent string representation by using the specified format and culture-specific formatting information.
'
' @param ts The <a href="TimeSpan.html">TimeSpan</a> to be converted to a string.
' @param Format A standard or custom <a href="TimeSpan.html">TimeSpan</a> format string.
' @param FormatProvider An object that supplies culture-specific formatting information.
' @return The string representation of the current <a href="TimeSpan.html">TimeSpan</a> value, as specified by <i>Format</i> and <i>FormatProvider</i>.
Public Function ToString(Optional ByRef Format As String, Optional ByVal FormatProvider As IFormatProvider) As String
    ToString = TimeSpan.ToString(Me, Format, FormatProvider)
End Function

''
' Returns a boolean indicating if the value and this TimeSpan
' instance are the same value.
'
' @param Value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As TimeSpan) As Boolean
    Equals = TimeSpan.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    With AsDLong(mMilliseconds)
        GetHashCode = .LoDWord Xor .HiDWord
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Days As Long, ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, ByVal Milliseconds As Long)
    InitFromMilliseconds (Days * MilliSecondsPerDay) + (Hours * MillisecondsPerHour) + (Minutes * MillisecondsPerMinute) + (Seconds * MillisecondsPerSecond) + Milliseconds
End Sub

Friend Sub InitFromMilliseconds(ByVal Value As Currency)
    mMilliseconds = Value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mMilliseconds = PropBag.ReadProperty(PropMilliseconds, DefaultMilliseconds)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PropMilliseconds, mMilliseconds
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = ToString(Format, Provider)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    On Error GoTo Catch
    
    Dim Other As TimeSpan
    Set Other = Value
    IObject_Equals = Equals(Other)
    
Catch:
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    On Error GoTo Catch
    
    Dim Other As TimeSpan
    Set Other = Value
    IComparable_CompareTo = CompareTo(Other)
    Exit Function
    
Catch:
    Error.Argument Argument_TimeSpanRequired, "Value"
End Function
