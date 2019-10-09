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

Private mParser     As New TimeSpanParser
Private mFormatter  As New TimeSpanFormatter
Private mZero       As New TimeSpan
Private mMaxValue   As TimeSpan
Private mMinValue   As TimeSpan

Private Type OperandMilliseconds
    Left    As Currency
    Right   As Currency
End Type

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> with a value of zero.
Public Property Get Zero() As TimeSpan
    Set Zero = mZero
End Property

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> with a maximum value.
'
' @remarks The value of this field is equivalent to the maximum <b>Currency</b> milliseconds.
' The string representation of this value is positive 10675199.02:48:05.4775807, or slightly more than 10,675,199 days.
Public Property Get MaxValue() As TimeSpan
    Set MaxValue = mMaxValue
End Property

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> with a minimum value.
'
' @remarks The value of this field is equivalent to the minimum <b>Currency</b> milliseconds.
' The string representation of this value is negative 10675199.02:48:05.4775808, or slightly more than negative 10,675,199 days.
Public Property Get MinValue() As TimeSpan
    Set MinValue = mMinValue
End Property

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> that represents a specified time, where the specification is in units of milliseconds.
'
' @param Value The number of milliseconds the <a href="TimeSpan.html">TimeSpan</a> will represent.
' @return An object that represents <i>Value</i>.
Public Function FromMilliseconds(ByVal Value As Currency) As TimeSpan
    Set FromMilliseconds = New TimeSpan
    FromMilliseconds.InitFromMilliseconds Value
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> that represents a specified time, where the specification is in units of seconds.
'
' @param Value The number of seconds the <a href="TimeSpan.html">TimeSpan</a> will represent.
' @return An object that represents <i>Value</i>.
Public Function FromSeconds(ByVal Value As Double) As TimeSpan
    Set FromSeconds = FromMilliseconds(Value * MillisecondsPerSecond)
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> that represents a specified time, where the specification is in units of minutes.
'
' @param Value The number of minutes the <a href="TimeSpan.html">TimeSpan</a> will represent.
' @return An object that represents <i>Value</i>.
Public Function FromMinutes(ByVal Value As Double) As TimeSpan
    Set FromMinutes = FromMilliseconds(Value * MillisecondsPerMinute)
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> that represents a specified time, where the specification is in units of hours.
'
' @param Value The number of hours the <a href="TimeSpan.html">TimeSpan</a> will represent.
' @return An object that represents <i>Value</i>.
Public Function FromHours(ByVal Value As Double) As TimeSpan
    Set FromHours = FromMilliseconds(Value * MillisecondsPerHour)
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> that represents a specified time, where the specification is in units of days.
'
' @param Value The number of days the <a href="TimeSpan.html">TimeSpan</a> will represent.
' @return An object that represents <i>Value</i>.
Public Function FromDays(ByVal Value As Double) As TimeSpan
    Set FromDays = FromMilliseconds(Value * MilliSecondsPerDay)
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> that represents a specified time, where the specification is in units of ticks.
'
' @param Value The number of ticks the <a href="TimeSpan.html">TimeSpan</a> will represent.
' @return An object that represents <i>Value</i>.
Public Function FromTicks(ByRef Value As Variant) As TimeSpan
    Set FromTicks = FromMilliseconds(CDec(Value) / 10000)
End Function

''
' Parses a string in the format of [d.]h:mm:ss[.ff] into a TimeSpan.
'
' @param s A string containing the date to be parsed.
' @return A TimeSpan created from the parsed string.
' @remarks This method will throw a FormatException exception if the string cannot be successfully parsed.
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

''
' Returns a new <a href="TimeSpan.html">TimeSpan</a> object whose value is the sum of two <a href="TimeSpan.html">TimeSpan</a> objects.
'
' @param Left The time interval to be added to.
' @param Right The time interval to add.
' @return A new object that represents the value of <i>Left</i> plus the value of <i>Right</i>.
Public Function Add(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As TimeSpan
    With GetMilliseconds(Left, Right)
        Set Add = FromMilliseconds(.Left + .Right)
    End With
End Function

''
' Returns a new <a href="TimeSpan.html">TimeSpan</a> object whose value is the difference between the two <a href="TimeSpan.html">TimeSpan</a> objects.
'
' @param Left The time interval to subtract from.
' @param Right The time interval to be subracted.
' @return A new time interval whose value is the result of the value of <i>Left</i> minus the value of <i>Right</i>.
Public Function Subtract(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As TimeSpan
    With GetMilliseconds(Left, Right)
        Set Subtract = FromMilliseconds(.Left - .Right)
    End With
End Function

''
' Compares two <a href="TimeSpan.html">TimeSpan</a> values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value.
'
' @param t1 The first time interval to compare.
' @param t2 The second time interval to compare.
' @return One of the following values.
' <table style="width: auto">
'   <tr>
'       <th>Value</th><th>Description</th>
'   </tr>
'   <tr>
'       <td>-1</td><td><i>t1</i> is shorter than <i>t2</i>.</td>
'   </tr>
'   <tr>
'       <td>0</td><td><i>t1</i> is equal to <i>t2</i>.</td>
'   </tr>
'   <tr>
'       <td>1</td><td><i>t1</i> is longer than <i>t2</i>.</td>
'   </tr>
' </table>
Public Function Compare(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Long
    With GetMilliseconds(t1, t2)
        If .Left < .Right Then
            Compare = -1
        ElseIf .Left > .Right Then
            Compare = 1
        End If
    End With
End Function

''
' Returns a boolean indicating if the <i>Left</i> <a href="TimeSpan.html">TimeSpan</a> instance is equal to the <i>Right</i> instance.
'
' @param Left The <a href="TimeSpan.html">TimeSpan</a> on the left-hand side of the equality.
' @param Right The <a href="TimeSpan.html">TimeSpan</a> on the right-hand side of the equality.
' @return Returns <b>True</b> if the <i>Left</i> instance is equal to the <i>Right</i> instance, <b>False</b> otherwise.
' @remarks This method will treat a value of <b>Nothing</b> as zero.
Public Function Equals(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Boolean
    Equals = Compare(Left, Right) = 0
End Function

''
' Returns a boolean indicating if the <i>Left</i> <a href="TimeSpan.html">TimeSpan</a> instance is less-than the <i>Right</i> instance.
'
' @param Left The <a href="TimeSpan.html">TimeSpan</a> on the left-hand side of the equality.
' @param Right The <a href="TimeSpan.html">TimeSpan</a> on the right-hand side of the equality.
' @return Returns <b>True</b> if the <i>Left</i> instance is less-than the <i>Right</i> instance, <b>False</b> otherwise.
' @remarks This method will treat a value of <b>Nothing</b> as zero.
Public Function LessThan(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Boolean
    LessThan = Compare(Left, Right) < 0
End Function

''
' Returns a boolean indicating if the <i>Left</i> <a href="TimeSpan.html">TimeSpan</a> instance is greater-than the <i>Right</i> instance.
'
' @param Left The <a href="TimeSpan.html">TimeSpan</a> on the left-hand side of the equality.
' @param Right The <a href="TimeSpan.html">TimeSpan</a> on the right-hand side of the equality.
' @return Returns <b>True</b> if the <i>Left</i> instance is greater-than the <i>Right</i> instance, <b>False</b> otherwise.
' @remarks This method will treat a value of <b>Nothing</b> as zero.
Public Function GreaterThan(ByRef Left As TimeSpan, ByRef Right As TimeSpan) As Boolean
    GreaterThan = Compare(Left, Right) > 0
End Function

''
' Attempts to parse the string into a TimeSpan object.
'
' @param s A string that specifies the time interval to convert.
' @param Result When this method returns, contains an object that represents the time interval specified by <i>s</i>, or <a href="TimeSpanStatic.get.Zero.html">TimeSpan.Zero</a> if the conversion failed.
' @return <b>True</b> if <i>s</i> was converted successfully; otherwise, <b>False</b>. This operation returns <b>false</b> if the <i>s</i> parameter empty,
' has an invalid format, represents a time interval that is less than <a href="TimeSpanStatic.get.MinValue.html">TimeSpan.MinValue</a> or greater than
' <a href="TimeSpanStatic.get.MaxValue.html">TimeSpan.MaxValue</a>, or has at least one days, hours, minutes, or seconds component outside its valid range.
' @remarks <p>The <a href="TimeSpanStatic.TryParse.html">TryParse</a> method is like the <a href="TimeSpanStatic.Parse.html">TimeSpan.Parse(String)</a> method, except that it does not throw an exception if the conversion fails.</p>
' <p>The <i>s</i> parameter contains a time interval specification in the form:</p>
' <p>[<i>ws</i>][-]{ <i>d</i> | <i>d</i>.<i>hh</i>:<i>mm</i>[:<i>ss</i>[.<i>ff</i>]] | <i>hh</i>:<i>mm</i>[:<i>ss</i>[.<i>ff</i>]] }[<i>ws</i>]</p>
' <p>Elements in square brackets ([ and ]) are optional. One selection from the list of alternatives enclosed in braces ({ and }) and separated by vertical bars (|) is required. The following table describes each element.</p>
' <table style="width: auto">
'   <tr>
'       <th>Element</th><th>Description</th>
'   </tr>
'   <tr>
'       <td><i>ws</i></td><td>Optional white space.</td>
'   </tr>
'   <tr>
'       <td>-</td><td>An optional minus sign, which indicates a negative <a href="TimeSpan.html">TimeSpan</a>.</td>
'   </tr>
'   <tr>
'       <td><i>d</i></td><td>Days, ranging from 0 to 10675199.</td>
'   </tr>
'   <tr>
'       <td>.</td><td>A culture-sensitive symbol that separates days from hours. The invariant format uses a period (".") character.</td>
'   </tr>
'   <tr>
'       <td><i>hh</i></td><td>Hours, ranging from 0 to 23.</td>
'   </tr>
'   <tr>
'       <td>:</td><td>The culture-sensitive time separator symbol. The invariant format uses a colon (":") character.</td>
'   </tr>
'   <tr>
'       <td><i>mm</i></td><td>Minutes, ranging from 0 to 59.</td>
'   </tr>
'    <tr>
'        <td>ss</td><td>Optional seconds, ranging from 0 to 59.</td>
'    </tr>
'    <tr>
'        <td>.</td><td>A culture-sensitive symbol that separates seconds from fractions of a second. The invariant format uses a period (".") character.</td>
'    </tr>
'    <tr>
'        <td><i>ff</i></td><td>Optional fractional seconds, consisting of one to seven decimal digits.</td>
'    </tr>
' </table>
' <p>The components of <i>s</i> must collectively specify a time interval that is greater than or equal to <a href="TimeSpanStatic.get.MinValue.html">TimeSpan.MinValue</a> and less than or equal to <a href="TimeSpanStatic.get.MaxValue.html">TimeSpan.MaxValue</a>.</p>
' <p>The <a href="TimeSpanStatic.Parse.html">Parse(String)</a> method tries to parse <i>s</i> by using each of the culture-specific formats for the current culture.</p>
Public Function TryParse(ByRef s As String, ByRef Result As TimeSpan) As Boolean
    TryParse = mParser.TryParse(s, Result)
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> object representing the time of day of the date.
'
' @param Value The date to retrieve the time-of-day from.
' @return An object representing the time-of-day.
Public Function FromDate(ByRef Value As Variant) As TimeSpan
    Set FromDate = CorDateTime.GetTimeOfDay(Value)
End Function

''
' Converts the value of the current <a href="TimeSpan.html">TimeSpan</a> object to its equivalent string representation by using the specified format and culture-specific formatting information.
'
' @param ts The <a href="TimeSpan.html">TimeSpan</a> to be converted to a string.
' @param Format A standard or custom <a href="TimeSpan.html">TimeSpan</a> format string.
' @param FormatProvider An object that supplies culture-specific formatting information.
' @return The string representation of the current <a href="TimeSpan.html">TimeSpan</a> value, as specified by <i>Format</i> and <i>FormatProvider</i>.
Public Function ToString(ByVal ts As TimeSpan, Optional ByRef Format As String, Optional ByVal FormatProvider As IFormatProvider) As String
    Dim ms As Currency
    
    If Not ts Is Nothing Then
        ms = ts.TotalMilliseconds
    End If
    
    ToString = mFormatter.Format(ms, Format, FormatProvider)
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
