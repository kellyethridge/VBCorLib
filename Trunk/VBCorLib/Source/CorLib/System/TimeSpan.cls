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
'    Module: TimeSpan
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
    If ts Is Nothing Then
        Set Add = Me
    Else
        Set Add = TimeSpan.FromMilliseconds(mMilliseconds + ts.TotalMilliseconds)
    End If
End Function

''
' Subtracts a timespan from this instance.
'
' @param ts TimeSpan to be subtracted.
' @return TimeSpan with the subtracted value.
Public Function Subtract(ByRef ts As TimeSpan) As TimeSpan
    If ts Is Nothing Then
        Set Subtract = Me
    Else
        Set Subtract = TimeSpan.FromMilliseconds(mMilliseconds - ts.TotalMilliseconds)
    End If
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
    Dim OtherMilliseconds As Currency
    If Not Value Is Nothing Then
        OtherMilliseconds = Value.TotalMilliseconds
    End If
    
    If mMilliseconds < OtherMilliseconds Then
        CompareTo = -1
    ElseIf mMilliseconds > OtherMilliseconds Then
        CompareTo = 1
    End If
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
    Dim sb As StringBuilder
    Set sb = TimeSpan.StringBuilder
    
    Dim ms As Currency
    ms = mMilliseconds
    If ms < 0 Then
        Call sb.AppendChar(vbMinus)
        ms = -ms
    End If
    
    Dim Days As Long
    Days = Int(ms / MILLISECONDS_PER_DAY)
    If Days <> 0 Then
        Call sb.Append(Days)
        Call sb.AppendChar(vbPeriod)
    End If
    Call sb.AppendQuick(FormatNumber(Int(ms / MILLISECONDS_PER_HOUR) Mod 24))
    Call sb.AppendChar(vbColon)
    Call sb.AppendQuick(FormatNumber(Int(ms / MILLISECONDS_PER_MINUTE) Mod 60))
    Call sb.AppendChar(vbColon)
    Call sb.AppendQuick(FormatNumber(Int(ms / MILLISECONDS_PER_SECOND) Mod 60))
    
    Dim Frac As Long
    Frac = Modulus(ms, MILLISECONDS_PER_SECOND)
    If Frac <> 0 Then
        Call sb.AppendChar(vbPeriod)
        Dim strt As String
        strt = Frac * 10000
        Call sb.AppendChar(vbZero, 7 - Len(strt))
        Call sb.AppendQuick(strt)
    Else
        Dim TicksComponent As Long
        TicksComponent = (mMilliseconds - Int(mMilliseconds)) * TICKS_PER_MILLISECOND
        If TicksComponent <> 0 Then
            sb.AppendChar vbPeriod
            
            Dim TicksString As String
            TicksString = TicksComponent
            sb.AppendChar vbZero, 7 - Len(TicksString)
            sb.AppendQuick TicksString
        End If
    End If
    ToString = sb.ToString
End Function

''
' Returns a boolean indicating if the value and this TimeSpan
' instance are the same value.
'
' @param Value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As TimeSpan) As Boolean
    If Value Is Nothing Then
        Exit Function
    End If
    
    Dim Span As TimeSpan
    Set Span = Value
    Equals = (mMilliseconds = Span.TotalMilliseconds)
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

''
' Returns if this TimeSpan instance is greater than another TimeSpan instance.
'
' @param ts The TimeSpan object to compare this instance against.
' @return Returns True if this is greater that, False otherwise.
'
Public Function GreaterThan(ByRef ts As TimeSpan) As Boolean
    If ts Is Nothing Then
        GreaterThan = True
    Else
        GreaterThan = (mMilliseconds > ts.TotalMilliseconds)
    End If
End Function

Public Function LessThan(ByRef ts As TimeSpan) As Boolean
    If ts Is Nothing Then Exit Function
    LessThan = (mMilliseconds < ts.TotalMilliseconds)
End Function

Public Function GreaterThanOrEqualTo(ByRef ts As TimeSpan) As Boolean
    If ts Is Nothing Then
        GreaterThanOrEqualTo = True
    Else
        GreaterThanOrEqualTo = (mMilliseconds >= ts.TotalMilliseconds)
    End If
End Function

Public Function LessThanOrEqualTo(ByRef ts As TimeSpan) As Boolean
    If ts Is Nothing Then Exit Function
    LessThanOrEqualTo = (mMilliseconds <= ts.TotalMilliseconds)
End Function

Public Function EqualTo(ByRef ts As TimeSpan) As Boolean
    If ts Is Nothing Then Exit Function
    EqualTo = (mMilliseconds = ts.TotalMilliseconds)
End Function

Public Function NotEqualTo(ByRef ts As TimeSpan) As Boolean
    If ts Is Nothing Then
        NotEqualTo = True
    Else
        NotEqualTo = (mMilliseconds <> ts.TotalMilliseconds)
    End If
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FormatNumber(ByVal n As Long) As String
    If n < 10 Then
        FormatNumber = "0" & n
    Else
        FormatNumber = n
    End If
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
    
    Dim Other As TimeSpan
    Set Other = Value
    IObject_Equals = Equals(Other)
    
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
    
    Dim Other As TimeSpan
    Set Other = Value
    IComparable_CompareTo = CompareTo(Other)
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException(Resources.GetString(Argument_TimeSpanRequired), Resources.GetParameter(Param.Value))
End Function
