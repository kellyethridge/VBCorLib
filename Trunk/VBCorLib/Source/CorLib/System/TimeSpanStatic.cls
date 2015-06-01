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
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: TimeSpanStatic
'

''
'   Provides severs static functions to help in creating TimeSpan objects.
'
' @see TimeSpan
'
Option Explicit

Private mParser         As New TimeSpanParser
Private mZero           As New TimeSpan
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
    Set FromSeconds = FromMilliseconds(Value * MillisecondsPerSecond)
End Function

''
' Returns a TimeSpan created from minutes.
'
' @param Value The number of minutes the timespan will represent.
' @return The TimeSpan initialized to the specified number of minutes.
'
Public Function FromMinutes(ByVal Value As Double) As TimeSpan
    Set FromMinutes = FromMilliseconds(Value * MillisecondsPerMinute)
End Function

''
' Returns a TimeSpan created from hours.
'
' @param Value The number of hours the timespan will represent.
' @return The TimeSpan initialized to the specified number of hours.
'
Public Function FromHours(ByVal Value As Double) As TimeSpan
    Set FromHours = FromMilliseconds(Value * MillisecondsPerHour)
End Function

''
' Returns a TimeSpan created from days.
'
' @param Value The number of days the timespan will represent.
' @return The TimeSpan initialized to the specified number of days.
'
Public Function FromDays(ByVal Value As Double) As TimeSpan
    Set FromDays = FromMilliseconds(Value * MillisecondsPerDay)
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
                Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidTimeSpan))
                
            Case ParseError.OverflowError
                Throw Cor.NewOverflowException(Environment.GetResourceString(Overflow_TimeSpan))
                
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
    Dim dt As CorDateTime
    Set dt = CorDateTime.GetDateTime(Value)
    Set FromDate = dt.TimeOfDay
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mMaxValue = FromMilliseconds(922337203685477.5807@)
    Set mMinValue = FromMilliseconds(-922337203685477.5807@ - 0.0001@)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMilliseconds(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As OperandMilliseconds
    If Not Left Is Nothing Then
        GetMilliseconds.Left = Left.TotalMilliseconds
    End If
    
    If Not Right Is Nothing Then
        GetMilliseconds.Right = Right.TotalMilliseconds
    End If
End Function
