VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpanParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2008 Kelly Ethridge
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
'    Module: TimeSpanParser
'

Option Explicit
Private Const MILLISECONDS_PER_SECOND       As Long = 1000
Private Const SECONDS_PER_MINUTE            As Long = 60
Private Const MINUTES_PER_HOUR              As Long = 60
Private Const HOURS_PER_DAY                 As Long = 24

Private Const MILLISECONDS_PER_MINUTE       As Long = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR         As Long = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY          As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Public Enum ParseError
    None
    OverflowError
    FormatError
End Enum

Private mBuffer     As WordBuffer
Private mIndex      As Long
Private mChar       As Long
Private mError      As ParseError



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get Error() As ParseError
    Error = mError
End Property

Friend Function TryParse(ByRef s As String, ByRef RetVal As TimeSpan) As Boolean
    Init s
    NextChar
    SkipBlanks
    
    Dim Sign As Currency
    If Char = vbMinus Then
        Sign = -1@
        NextChar
    Else
        Sign = 1@
    End If
    
    Dim Result As Currency
    Select Case NextNonDigit()
        Case vbColon
            If Not TryParseTime(Result) Then
                Exit Function
            End If
    
        Case vbZero To vbNine, vbPeriod, vbNullPtr
            Const MaxDays As Long = 10675199
            Dim DaysParsed As Long
            If Not TryParseNumber(MaxDays, DaysParsed) Then
                Exit Function
            End If
            Result = Result + DaysParsed * MILLISECONDS_PER_DAY
            
            If Char = vbPeriod Then
                NextChar
                
                Dim TimeParsed As Currency
                If Not TryParseTime(TimeParsed) Then
                    Exit Function
                End If
                
                Result = Result + TimeParsed
            End If
            
        Case Else
            mError = FormatError
            Exit Function
    End Select
    
    SkipBlanks
    
    If mIndex < mBuffer.SA.cElements - 1 Then
        If mError = ParseError.None Then
            mError = ParseError.FormatError
        End If
        
        Exit Function
    End If
    
    Set RetVal = TimeSpan.FromMilliseconds(Result * Sign)
    TryParse = True
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init(ByRef s As String)
    InitWordBuffer mBuffer, StrPtr(s), Len(s)
    mIndex = -1
    mError = ParseError.None
    mChar = 0
End Sub

Private Property Get Char() As Integer
    Char = mChar
End Property

Private Sub NextChar()
    mIndex = mIndex + 1
    
    If mIndex < mBuffer.SA.cElements Then
        mChar = mBuffer.Data(mIndex)
    Else
        mIndex = mIndex - 1
        mChar = 0
    End If
End Sub

Private Sub SkipBlanks()
    Do While Char = vbSpace
        NextChar
    Loop
End Sub

Private Function NextNonDigit() As Integer
    Dim i As Long
    For i = mIndex To mBuffer.SA.cElements - 1
        Select Case mBuffer.Data(i)
            Case vbZero To vbNine
            Case Else
                NextNonDigit = mBuffer.Data(i)
                Exit Function
        End Select
    Next i
End Function

Private Function TryParseTime(ByRef Milliseconds As Currency) As Boolean
    Dim Result As Currency
    
    Const MaxHours As Long = 23
    Dim HoursParsed As Long
    If Not TryParseNumber(MaxHours, HoursParsed) Then
        Exit Function
    End If
    
    Result = HoursParsed * MILLISECONDS_PER_HOUR
    
    ' when parsing time we require at least hours and minutes.
    If Char <> vbColon Then
        mError = ParseError.FormatError
        Exit Function
    End If
    
    NextChar
    
    Const MaxMinutes As Long = 59
    Dim MinutesParsed As Long
    If Not TryParseNumber(MaxMinutes, MinutesParsed) Then
        Exit Function
    End If
    
    Result = Result + MinutesParsed * MILLISECONDS_PER_MINUTE
    
    ' check for possible seconds
    If Char = vbColon Then
        NextChar
        
        ' defined fraction immediately after seconds colon instead
        ' 00:00:.1234
        If Char <> vbPeriod Then
            ' fraction not defined immediately, so seconds must be.
            Const MaxSeconds As Long = 59
            Dim SecondsParsed As Long
            If Not TryParseNumber(MaxSeconds, SecondsParsed) Then
                Exit Function
            End If
            
            Result = Result + SecondsParsed * MILLISECONDS_PER_SECOND
        End If
        
        If Char = vbPeriod Then
            NextChar
            
            Dim FractionParsed As Long
            If Not TryParseNumber(10000000, FractionParsed) Then
                Exit Function
            End If
            
            Do While FractionParsed < 10000000
                FractionParsed = FractionParsed * 10
            Loop
            FractionParsed = FractionParsed \ 10
            
            Result = Result + (CCur(FractionParsed) * 0.0001@)
        End If
    End If
    
    Milliseconds = Result
    TryParseTime = True
End Function

Private Function TryParseNumber(ByVal MaxValue As Long, ByRef Result As Long) As Boolean
    Dim Number          As Long
    Dim OriginalIndex   As Long
    
    OriginalIndex = mIndex
    Do
        Select Case Char
            Case vbZero To vbNine
                Number = Number * 10 + (Char - vbZero)
                
            Case Else
                Exit Do
                
        End Select
        
        NextChar
    Loop
    
    If Number <= MaxValue Then
        Result = Number
        TryParseNumber = True
    ElseIf mIndex = OriginalIndex Then
        mError = ParseError.FormatError
    Else
        mError = ParseError.OverflowError
    End If
End Function

