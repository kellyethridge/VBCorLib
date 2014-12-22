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
' @see cDateTime
'
Option Explicit
Implements IObject
Implements IComparable

Private Const PROP_MILLISECONDS         As String = "Milliseconds"
Private Const DEF_TIMESPAN              As Long = 0

Private Const SECONDS_PER_MINUTE        As Long = 60
Private Const MINUTES_PER_HOUR          As Long = 60
Private Const HOURS_PER_DAY             As Long = 24

Private Const TICKS_PER_MILLISECOND     As Currency = 10000
Private Const MILLISECONDS_PER_TICK     As Currency = 0.0001@
Private Const MILLISECONDS_PER_SECOND   As Currency = 1000@
Private Const MILLISECONDS_PER_MINUTE   As Currency = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR     As Currency = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY      As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

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
    TotalSeconds = mMilliseconds / MILLISECONDS_PER_SECOND
End Property

''
' Returns the duration of this instance in minutes.
'
' @return This timespan in minutes.
' @remarks The returned duration can contain a fractional
' part if necessary to represent the duration.
Public Property Get TotalMinutes() As Double
    TotalMinutes = mMilliseconds / MILLISECONDS_PER_MINUTE
End Property

''
' Returns the duration of this instance in hours.
'
' @return This timespan in hours.
Public Property Get TotalHours() As Double
    TotalHours = mMilliseconds / MILLISECONDS_PER_HOUR
End Property

''
' Returns the duration of this instance in days.
'
' @return This timespan in days.
' @remarks The returned duration can contain a fractional
' part if necessary to represent the duration.
Public Property Get TotalDays() As Double
    TotalDays = mMilliseconds / MILLISECONDS_PER_DAY
End Property

''
' Returns the Milliseconds portion of the duration.
'
' @return The Milliseconds portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Milliseconds() As Long
    Milliseconds = AsLong(mMilliseconds * MILLISECONDS_PER_TICK) Mod 1000
End Property

''
' Returns the Seconds portion of the duration.
'
' @return The Seconds portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Seconds() As Long
    Seconds = Fix(mMilliseconds / MILLISECONDS_PER_SECOND) Mod 60
End Property

''
' Returns the Minutes portion of the duration.
'
' @return The Minutes portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Minutes() As Long
    Minutes = Fix(mMilliseconds / MILLISECONDS_PER_MINUTE) Mod 60
End Property

''
' Returns the Hours portion of the duration.
'
' @return The Hours portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Hours() As Long
    Hours = Fix(mMilliseconds / MILLISECONDS_PER_HOUR) Mod 24
End Property

''
' Returns the Days portion of the duration.
'
' @return The Days portion of the duration.
' @remarks This does not represent the entire duration of the
' timespan, only the specified portion of it.
Public Property Get Days() As Long
    Days = Fix(mMilliseconds / MILLISECONDS_PER_DAY)
End Property

''
' Returns the duration in ticks.
'
' @return The duration in ticks
' @remarks A tick is defined as 100 nanoseconds, or 1/10000 of a millisecond.
Public Property Get Ticks() As Variant
    Ticks = CDec(mMilliseconds) * TICKS_PER_MILLISECOND
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
    CompareTo = TimeSpan.CompareTo(Me, Value)
End Function

''
' Returns a TimeSpan with the opposite sign of this instance.
'
' @return The timespan with the opposite sign as this duration.
Public Function Negate() As TimeSpan
    Set Negate = TimeSpan.FromMilliseconds(-mMilliseconds)
End Function

''
' Returns a string representation of the duration.
'
' @return String representing the duration.
' @remarks A string representation is in the form of [-][d.]hh:mm:ss[.ff]
' where the bracketed fields are optional. The optional fields will be
' included in the string their value is not zero, or the TimeSpan is negative.
Public Function ToString() As String
    ' use a cached version to lower overhead. The length is set to zero.
    Dim Output As StringBuilder
    Set Output = TimeSpan.StringBuilder
    
    Dim ms As Currency
    ms = mMilliseconds
    If ms < 0 Then
        Output.AppendChar vbMinusChar
        ms = -ms
    End If
    
    Dim Days As Long
    Days = Int(ms / MILLISECONDS_PER_DAY)
    If Days <> 0 Then
        Output.Append Days
        Output.AppendChar vbPeriodChar
    End If

    AppendComponent Output, Int(ms / MILLISECONDS_PER_HOUR) Mod 24, 2
    Output.AppendChar vbColonChar
    AppendComponent Output, Int(ms / MILLISECONDS_PER_MINUTE) Mod 60, 2
    Output.AppendChar vbColonChar
    AppendComponent Output, Int(ms / MILLISECONDS_PER_SECOND) Mod 60, 2
    
    Dim MillisecondComponent    As Long
    Dim TicksComponent          As Long
    MillisecondComponent = Modulus(ms, MILLISECONDS_PER_SECOND)
    TicksComponent = (mMilliseconds - Int(mMilliseconds)) * TICKS_PER_MILLISECOND
    
    If MillisecondComponent <> 0 Then
        Output.AppendChar vbPeriodChar
        AppendComponent Output, MillisecondComponent, 3
        AppendComponent Output, TicksComponent, 4
    ElseIf TicksComponent <> 0 Then
        Output.AppendChar vbPeriodChar
        AppendComponent Output, TicksComponent, 7
    End If
    
    ToString = Output.ToString
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

Public Function LessThan(ByRef other As TimeSpan) As Boolean
    LessThan = TimeSpan.LessThan(Me, other)
End Function

Public Function GreaterThan(ByRef other As TimeSpan) As Boolean
    GreaterThan = TimeSpan.GreaterThan(Me, other)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Days As Long, ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, ByVal Milliseconds As Long)
    InitFromMilliseconds (Days * MILLISECONDS_PER_DAY) + (Hours * MILLISECONDS_PER_HOUR) + (Minutes * MILLISECONDS_PER_MINUTE) + (Seconds * MILLISECONDS_PER_SECOND) + Milliseconds
End Sub

Friend Sub InitFromMilliseconds(ByVal Value As Currency)
    mMilliseconds = Value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AppendComponent(ByRef Output As StringBuilder, ByVal Component As Long, ByVal Width As Long)
    Dim ComponentString As String
    ComponentString = Component
    
    Output.AppendChar vbZero, Width - Len(ComponentString)
    Output.AppendString ComponentString
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mMilliseconds = PropBag.ReadProperty(PROP_MILLISECONDS, DEF_TIMESPAN)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_MILLISECONDS, mMilliseconds
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    On Error GoTo errTrap
    
    Dim other As TimeSpan
    Set other = Value
    IObject_Equals = Equals(other)
    
errTrap:
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable_CompareTo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    On Error GoTo errTrap
    
    Dim other As TimeSpan
    Set other = Value
    IComparable_CompareTo = CompareTo(other)
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException(Resources.GetString(Argument_TimeSpanRequired), Resources.GetParameter(Param.Value))
End Function
