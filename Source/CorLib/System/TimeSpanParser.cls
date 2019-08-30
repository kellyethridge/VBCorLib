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
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: TimeSpanParser
'

Option Explicit
Private Const MillisecondsPerSecond       As Long = 1000
Private Const SecondsPerMinute            As Long = 60
Private Const MinutesPerHour              As Long = 60
Private Const HoursPerDay                 As Long = 24

Private Const MillisecondsPerMinute       As Long = MillisecondsPerSecond * SecondsPerMinute
Private Const MillisecondsPerHour         As Long = MillisecondsPerMinute * MinutesPerHour
Private Const MilliSecondsPerDay          As Currency = MillisecondsPerHour * HoursPerDay

Public Enum ParseError
    None
    OverflowError
    FormatError
End Enum

Private mChars()    As Integer
Private mIndex      As Long
Private mChar       As Integer
Private mError      As ParseError


Friend Property Get Error() As ParseError
    Error = mError
End Property

Friend Function TryParse(ByRef s As String, ByRef RetVal As TimeSpan) As Boolean
    Dim Sign    As Long
    Dim Result  As Currency
    
    Init s
    NextChar
    SkipBlanks
    
    If Char = vbMinusChar Then
        Sign = -1
        NextChar
    Else
        Sign = 1
    End If
    
    Select Case NextNonDigit()
        Case vbColonChar
            If Not TryParseTime(Result) Then
                GoTo Finally
            End If
    
        Case vbZeroChar To vbNineChar, vbPeriodChar, vbNullPtr
            Const MaxDays As Long = 10675199
            Dim DaysParsed As Long
            
            If Not TryParseNumber(MaxDays, DaysParsed) Then
                GoTo Finally
            End If
            
            Result = Result + DaysParsed * MilliSecondsPerDay
            
            If Char = vbPeriodChar Then
                NextChar
                
                Dim TimeParsed As Currency
                If Not TryParseTime(TimeParsed) Then
                    GoTo Finally
                End If
                
                Result = Result + TimeParsed
            End If
            
        Case Else
            mError = FormatError
            GoTo Finally
    End Select
    
    SkipBlanks
    
    If mIndex < UBound(mChars) Then
        If mError = ParseError.None Then
            mError = ParseError.FormatError
        End If
        
        GoTo Finally
    End If
    
    Set RetVal = TimeSpan.FromMilliseconds(Result * Sign)
    TryParse = True
    
Finally:
    FreeChars mChars
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init(ByRef s As String)
    mChars = AllocChars(s)
    mIndex = -1
    mError = ParseError.None
    mChar = 0
End Sub

Private Property Get Char() As Integer
    Char = mChar
End Property

Private Sub NextChar()
    mIndex = mIndex + 1
    
    If mIndex <= UBound(mChars) Then
        mChar = mChars(mIndex)
    Else
        mIndex = mIndex - 1
        mChar = 0
    End If
End Sub

Private Sub SkipBlanks()
    Do While Char = vbSpaceChar
        NextChar
    Loop
End Sub

Private Function NextNonDigit() As Integer
    Dim i As Long
    
    If UBound(mChars) >= 0 Then
        For i = mIndex To UBound(mChars)
            Select Case mChars(i)
                Case vbZeroChar To vbNineChar
                Case Else
                    NextNonDigit = mChars(i)
                    Exit Function
            End Select
        Next
    Else
        NextNonDigit = -1
    End If
End Function

Private Function TryParseTime(ByRef Milliseconds As Currency) As Boolean
    Dim Result As Currency
    
    Const MaxHours As Long = 23
    Dim HoursParsed As Long
    If Not TryParseNumber(MaxHours, HoursParsed) Then
        Exit Function
    End If
    
    Result = HoursParsed * MillisecondsPerHour
    
    ' when parsing time we require at least hours and minutes.
    If Char <> vbColonChar Then
        mError = ParseError.FormatError
        Exit Function
    End If
    
    NextChar
    
    Const MaxMinutes As Long = 59
    Dim MinutesParsed As Long
    If Not TryParseNumber(MaxMinutes, MinutesParsed) Then
        Exit Function
    End If
    
    Result = Result + MinutesParsed * MillisecondsPerMinute
    
    ' check for possible seconds
    If Char = vbColonChar Then
        NextChar
        
        ' defined fraction immediately after seconds colon instead
        ' 00:00:.1234
        If Char <> vbPeriodChar Then
            ' fraction not defined immediately, so seconds must be.
            Const MaxSeconds As Long = 59
            Dim SecondsParsed As Long
            If Not TryParseNumber(MaxSeconds, SecondsParsed) Then
                Exit Function
            End If
            
            Result = Result + SecondsParsed * MillisecondsPerSecond
        End If
        
        If Char = vbPeriodChar Then
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
            Case vbZeroChar To vbNineChar
                Number = Number * 10 + (Char - vbZeroChar)
                
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


