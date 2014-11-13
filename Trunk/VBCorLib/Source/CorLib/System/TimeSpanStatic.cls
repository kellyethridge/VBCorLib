VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpanStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: TimeSpanStatic
'

''
'   Provides severs static functions to help in creating TimeSpan objects.
'
' @see TimeSpan
'
Option Explicit
Private Const MILLISECONDS_PER_SECOND       As Long = 1000
Private Const SECONDS_PER_MINUTE            As Long = 60
Private Const MINUTES_PER_HOUR              As Long = 60
Private Const HOURS_PER_DAY                 As Long = 24

Private Const MILLISECONDS_PER_MINUTE       As Long = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR         As Long = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY          As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private mZero           As New TimeSpan
Private mStringBuilder  As New StringBuilder
Private mParser         As New TimeSpanParser
Private mMaxValue       As TimeSpan
Private mMinValue       As TimeSpan

Private Type OperandMilliseconds
    Left    As Currency
    Right   As Currency
End Type

''
' Returns a TimeSpan with a value of zero.
'
' @return A zero timespan.
'
Public Property Get Zero() As TimeSpan
    Set Zero = mZero
End Property

''
' Returns a TimeSpan with a maximum value.
'
' @return A maximum value timespan.
'
Public Property Get MaxValue() As TimeSpan
    Set MaxValue = mMaxValue
End Property

''
' Returns a TimeSpan with a minimum value.
'
' @return A minimum value timespan.
'
Public Property Get MinValue() As TimeSpan
    Set MinValue = mMinValue
End Property

''
' Returns a TimeSpan created from milliseconds.
'
' @param Value The number of milliseconds the timespan will represent.
' @return The TimeSpan initialized to the specified number of milliseconds.
'
Public Function FromMilliseconds(ByVal Value As Currency) As TimeSpan
    Set FromMilliseconds = New TimeSpan
    FromMilliseconds.InitFromMilliseconds Value
End Function

''
' Returns a TimeSpan created from seconds.
'
' @param Value The number of seconds the timespan will represent.
' @return The TimeSpan initialized to the specified number of seconds.
'
Public Function FromSeconds(ByVal Value As Double) As TimeSpan
    Set FromSeconds = FromMilliseconds(Value * MILLISECONDS_PER_SECOND)
End Function

''
' Returns a TimeSpan created from minutes.
'
' @param Value The number of minutes the timespan will represent.
' @return The TimeSpan initialized to the specified number of minutes.
'
Public Function FromMinutes(ByVal Value As Double) As TimeSpan
    Set FromMinutes = FromMilliseconds(Value * MILLISECONDS_PER_MINUTE)
End Function

''
' Returns a TimeSpan created from hours.
'
' @param Value The number of hours the timespan will represent.
' @return The TimeSpan initialized to the specified number of hours.
'
Public Function FromHours(ByVal Value As Double) As TimeSpan
    Set FromHours = FromMilliseconds(Value * MILLISECONDS_PER_HOUR)
End Function

''
' Returns a TimeSpan created from days.
'
' @param Value The number of days the timespan will represent.
' @return The TimeSpan initialized to the specified number of days.
'
Public Function FromDays(ByVal Value As Double) As TimeSpan
    Set FromDays = FromMilliseconds(Value * MILLISECONDS_PER_DAY)
End Function

''
' Returns a TimeSpan created from ticks.
'
' @param Value The number of ticks the timespan will represent.
' @return The TimeSpan initialized to the specified number of ticks.
'
Public Function FromTicks(ByRef Value As Variant) As TimeSpan
    Set FromTicks = FromMilliseconds(CDec(Value) / 10000)
End Function

''
' Parses a string in the format of [d.]h:mm:ss[.ff] into a TimeSpan.
'
' @param s A string containing the date to be parsed.
' @return A TimeSpan created from the parsed string.
' @remarks This method will throw a FormatException exception if the
' string cannot be successfully parsed.
'
Public Function Parse(ByVal s As String) As TimeSpan
    If Not mParser.TryParse(s, Parse) Then
        Select Case mParser.Error
            Case ParseError.FormatError
                Throw Cor.NewFormatException(Resources.GetString(Format_InvalidTimeSpan))
                
            Case ParseError.OverflowError
                Throw Cor.NewOverflowException(Resources.GetString(Overflow_TimeSpan))
                
            Case Else
                Throw New ArgumentException
        End Select
    End If
End Function

Public Function Add(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As TimeSpan
    With GetMilliseconds(Left, Right)
        Set Add = FromMilliseconds(.Left + .Right)
    End With
End Function

Public Function Subtract(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As TimeSpan
    With GetMilliseconds(Left, Right)
        Set Subtract = FromMilliseconds(.Left - .Right)
    End With
End Function

Public Function CompareTo(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Long
    With GetMilliseconds(Left, Right)
        If .Left < .Right Then
            CompareTo = -1
        ElseIf .Left > .Right Then
            CompareTo = 1
        End If
    End With
End Function

Public Function Equals(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Boolean
    With GetMilliseconds(Left, Right)
        Equals = .Left = .Right
    End With
End Function

Public Function LessThan(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Boolean
    With GetMilliseconds(Left, Right)
        LessThan = .Left < .Right
    End With
End Function

Public Function GreaterThan(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Boolean
    With GetMilliseconds(Left, Right)
        GreaterThan = .Left > .Right
    End With
End Function

''
' Attempts to parse the string into a TimeSpan object.
'
' @param s The string containing the characters to be parsed.
' @param Result The output TimeSpan object.
' @return Returns True if the string is successfully parsed, false otherwise.
' @remarks This method does not raise an error if the string cannot be parsed.
'
Public Function TryParse(ByVal s As String, ByRef RetVal As TimeSpan) As Boolean
    TryParse = mParser.TryParse(s, RetVal)
End Function

''
' Returns a TimeSpan object representing the time of day of the date.
'
' @param Value The date to retrieve the time of day for.
' @return A TimeSpan representing the time of day.
'
Public Function FromDate(ByRef Value As Variant) As TimeSpan
    Dim dt As cDateTime
    Set dt = cDateTime.GetcDateTime(Value)
    Set FromDate = dt.TimeOfDay
End Function

' This is to prevent a TimeSpan from having to create and
' destroy an object. A TimeSpan may live very short lives when
' doing calculations. If for some reason a large number of
' TimeSpans is created and also converted to a string using
' ToString, the overhead of having to create and destroy a
' StringBuilder object will cost twice as much overhead as
' creating and destroying just the TimeSpan object.
'
' This is to be used only by TimeSpan objects in an immediate
' context since it is shared by all TimeSpan objects.
Friend Property Get StringBuilder() As StringBuilder
    mStringBuilder.Length = 0
    Set StringBuilder = mStringBuilder
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mMaxValue = FromMilliseconds(922337203685477.5807@)
    Set mMinValue = FromMilliseconds(-922337203685477.5807@ - 0.0001@)
End Sub

Private Function GetMilliseconds(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As OperandMilliseconds
    If Not Left Is Nothing Then
        GetMilliseconds.Left = Left.TotalMilliseconds
    End If
    
    If Not Right Is Nothing Then
        GetMilliseconds.Right = Right.TotalMilliseconds
    End If
End Function
