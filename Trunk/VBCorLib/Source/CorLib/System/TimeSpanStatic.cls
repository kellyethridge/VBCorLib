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

Private mZero           As TimeSpan
Private mMaxValue       As TimeSpan
Private mMinValue       As TimeSpan
Private mStringBuilder  As StringBuilder
Private mMSPlaces()     As Currency
Private mParser         As New TimeSpanParser


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
    Call FromMilliseconds.InitFromMilliseconds(Value)
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
                Throw New FormatException
                
            Case ParseError.OverflowError
                Throw New OverflowException
                
            Case Else
                Throw New ArgumentException
        End Select
    End If
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

Private Function GetSign(ByRef Buf As WordBuffer, ByRef Pos As Long, ByRef RetVal As Currency) As Boolean
    If Pos < Buf.SA.cElements Then
        If Buf.Data(0) = vbMinus Then
            RetVal = -1@
            Pos = Pos + 1
        Else
            RetVal = 1@
        End If
        
        GetSign = True
    End If
End Function

Private Function GetNumber(ByRef Buf As WordBuffer, ByRef Pos As Long, ByRef RetVal As Long, ByVal MaxDigits As Long) As Boolean
    Dim Result  As Long
    Dim Index   As Long
    Index = Pos
    
    Do While Index < Buf.SA.cElements
        Select Case Buf.Data(Index)
            Case vbZero To vbNine
                Result = Result * 10 + (Buf.Data(Index) - vbZero)
                
            Case Else
                Exit Do
        
        End Select
        
        Index = Index + 1
    Loop
    
    If Index > Pos Then
        Pos = Index
        RetVal = Result
        GetNumber = True
    End If
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
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NextIsColonOrEnd(ByRef Buf As WordBuffer, ByVal Pos As Long) As Boolean
    Dim Result As Boolean
    
    If Pos >= Buf.SA.cElements Then
        Result = True
    Else
        Result = (Buf.Data(Pos) = vbColon)
    End If
    
    NextIsColonOrEnd = Result
End Function

Private Function GetFraction(ByRef Buf As WordBuffer, ByVal Pos As Long, ByRef Result As Currency) As Long
    Dim ch      As Integer
    Dim Ret     As Currency
    Dim place   As Long
    Dim Index   As Long
    
    If Pos >= Buf.SA.cElements Then Exit Function
    
    Index = Pos
    ch = Buf.Data(Index)
    
    Do While ch <> 0 And place < 7
        Select Case ch
            Case vbZero To vbNine
                Ret = Ret + mMSPlaces(place) * (ch - vbZero)
                place = place + 1
                
            Case Else
                Exit Function
                
        End Select
        
        Index = Index + 1
        ch = Buf.Data(Index)
    Loop
    
    Result = Ret
    GetFraction = Index - Pos
End Function

Private Function GetComponent(ByRef Buf As WordBuffer, ByVal Pos As Long, ByRef Result As Long) As Long
    Dim ch      As Integer
    Dim Ret     As Long
    Dim Index   As Long
    
    If Pos >= Buf.SA.cElements Then Exit Function
    
    Index = Pos
    ch = Buf.Data(Index)
    
    Do
        Select Case ch
            Case vbZero To vbNine
                Ret = Ret * 10 + ch - vbZero
            
            Case vbNullPtr, vbPeriod, vbColon
                Result = Ret
                GetComponent = Index - Pos
                Exit Function
                
            Case Else
                Exit Function
        End Select
        
        Index = Index + 1
        ch = Buf.Data(Index)
    Loop
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mZero = New TimeSpan
    Set mMaxValue = FromMilliseconds(922337203685477.5807@)
    Set mMinValue = FromMilliseconds(-922337203685477.5807@ - 0.0001@)
    Set mStringBuilder = New StringBuilder
    mMSPlaces = Cor.NewCurrencys(100@, 10@, 1@, 0.1@, 0.01@, 0.001@, 0.0001@)
End Sub
