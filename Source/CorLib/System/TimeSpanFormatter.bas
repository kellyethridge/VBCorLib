Attribute VB_Name = "TimeSpanFormatter"
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
' Module: TimeSpanFormatter
'

Option Explicit

Private Const MaxDayWidth       As Long = 8
Private Const MaxHourWidth      As Long = 2
Private Const MaxMinuteWidth    As Long = 2
Private Const MaxSecondWidth    As Long = 2
Private Const MaxFractionWidth  As Long = 7

Private Enum FormatType
    Standard
    Minimum
    Full
End Enum

Private mFormatChars()      As Integer
Private mOutput             As StringBuilder
Private mTotalMilliseconds  As Currency
Private mDays               As Long
Private mHours              As Long
Private mMinutes            As Long
Private mSeconds            As Long
Private mFraction           As Long


Public Function FormatTimeSpan(ByVal TotalMilliseconds As Currency, ByRef FormatString As String, ByVal FormatProvider As IFormatProvider) As String
    InitComponents TotalMilliseconds
    Set mOutput = StringBuilderCache.Acquire
    
    On Error GoTo Catch
    
    Select Case FormatString
        Case "", "c", "t", "T"
            FormatStandard FormatType.Standard, FormatProvider
        Case "g"
            FormatStandard FormatType.Minimum, FormatProvider
        Case "G"
            FormatStandard FormatType.Full, FormatProvider
        Case Else
            If Len(FormatString) = 1 Then _
                Error.Format Format_InvalidString
            
            mFormatChars = AllocChars(FormatString)
            FormatCustom
    End Select
    
    GoSub Finally
    FormatTimeSpan = StringBuilderCache.GetStringAndRelease(mOutput)
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars mFormatChars
    Return
End Function

Private Sub InitComponents(ByVal TotalMilliseconds As Currency)
    Dim Milliseconds    As Long
    Dim Ticks           As Long
    Dim ms              As Currency
    
    mTotalMilliseconds = TotalMilliseconds
    ms = Abs(TotalMilliseconds)
    mDays = Int(ms / MilliSecondsPerDay)
    mHours = Int(ms / MillisecondsPerHour) Mod HoursPerDay
    mMinutes = Int(ms / MillisecondsPerMinute) Mod MinutesPerHour
    mSeconds = Int(ms / MillisecondsPerSecond) Mod SecondsPerMinute
    Milliseconds = Modulus(ms, MillisecondsPerSecond)
    Ticks = (ms - Int(ms)) * TicksPerMillisecond
    mFraction = Milliseconds * 10000 + Ticks
End Sub

Private Sub FormatCustom()
    Dim Index           As Long
    Dim MaxIndex        As Long
    Dim Count           As Long
    Dim InWholeMode     As Boolean
    Dim InEscapeMode    As Boolean
    
    MaxIndex = UBound(mFormatChars)
    
    Do While Index <= MaxIndex
        If InWholeMode Then
            Select Case mFormatChars(Index)
                Case vbPercentChar
                    Error.Format Format_InvalidString
                Case vbLowerDChar
                    AppendComponent mDays, 1
                Case vbLowerHChar
                    AppendComponent mHours, 1
                Case vbLowerMChar
                    AppendComponent mMinutes, 1
                Case vbLowerSChar
                    AppendComponent mSeconds, 1
                Case vbLowerFChar
                    AppendFraction 1, False
                Case vbUpperFChar
                    AppendFraction 1, True
                Case Else
                    Error.Format Format_InvalidString
            End Select
        
            InWholeMode = False
            Index = Index + 1
        ElseIf InEscapeMode Then
            mOutput.AppendChar mFormatChars(Index)
            InEscapeMode = False
            Index = Index + 1
        Else
            Select Case mFormatChars(Index)
                Case vbPercentChar
                    InWholeMode = True
                    Count = 1
                Case vbBackslashChar
                    InEscapeMode = True
                    Count = 1
                Case vbLowerDChar
                    Count = CountChars(mFormatChars, vbLowerDChar, Index)
                    If Count > MaxDayWidth Then FormatError
                    AppendComponent mDays, Count
                Case vbLowerHChar
                    Count = CountChars(mFormatChars, vbLowerHChar, Index)
                    If Count > MaxHourWidth Then FormatError
                    AppendComponent mHours, Count
                Case vbLowerMChar
                    Count = CountChars(mFormatChars, vbLowerMChar, Index)
                    If Count > MaxMinuteWidth Then FormatError
                    AppendComponent mMinutes, Count
                Case vbLowerSChar
                    Count = CountChars(mFormatChars, vbLowerSChar, Index)
                    If Count > MaxSecondWidth Then FormatError
                    AppendComponent mSeconds, Count
                Case vbLowerFChar
                    Count = CountChars(mFormatChars, vbLowerFChar, Index)
                    If Count > MaxFractionWidth Then FormatError
                    AppendFraction Count, False
                Case vbUpperFChar
                    Count = CountChars(mFormatChars, vbUpperFChar, Index)
                    If Count > MaxFractionWidth Then FormatError
                    AppendFraction Count, True
                Case vbSingleQuoteChar
                    Count = AppendStringLiteral(Index)
                Case Else
                    Error.Format Format_InvalidString
            End Select
            
            Index = Index + Count
        End If
    Loop
    
    If InWholeMode Or InEscapeMode Then
        Error.Format Format_InvalidString
    End If
End Sub

Private Sub FormatError()
    Error.Format Format_InvalidString
End Sub

Private Function AppendStringLiteral(ByVal Index As Long) As Long
    Dim LiteralStart    As Long
    Dim MaxIndex        As Long
    Dim Count           As Long
    
    LiteralStart = Index + 1
    MaxIndex = UBound(mFormatChars)
    
    Do While Index < MaxIndex
        Index = Index + 1
        
        If mFormatChars(Index) = vbSingleQuoteChar Then
            mOutput.Append mFormatChars, LiteralStart, Count
            AppendStringLiteral = Count + 2 ' we add two for the single quotes.
            Exit Function
        End If
        
        Count = Count + 1
    Loop
    
    Throw Cor.NewFormatException(Environment.GetResourceString(Format_BadQuote, "'"))
End Function

Private Sub FormatStandard(ByVal FormatType As FormatType, ByVal Provider As IFormatProvider)
    If mTotalMilliseconds < 0 Then
        mOutput.AppendChar vbMinusChar
        mTotalMilliseconds = -mTotalMilliseconds
    End If
    
    If mDays <> 0 Or FormatType = Full Then
        mOutput.Append mDays
        mOutput.AppendChar IIfLong(FormatType = Standard, vbPeriodChar, vbColonChar)
    End If

    AppendComponent mHours, IIfLong(FormatType = Minimum, 1, 2)
    mOutput.AppendChar vbColonChar
    AppendComponent mMinutes, 2
    mOutput.AppendChar vbColonChar
    AppendComponent mSeconds, 2
    
    If mFraction <> 0 Or FormatType = Full Then
        mOutput.AppendString GetFractionSeparator(FormatType, Provider)
        AppendComponent mFraction, MaxFractionWidth
    End If
    
    If (FormatType = Minimum) And (mFraction) Then
        Dim NewLength As Long
        
        NewLength = mOutput.Length
        
        Do While mOutput(NewLength - 1) = vbZeroChar
            NewLength = NewLength - 1
        Loop
        
        mOutput.Length = NewLength
    End If
End Sub

Private Function GetFractionSeparator(ByVal FormatType As FormatType, ByVal Provider As IFormatProvider) As String
    Dim Nfi As NumberFormatInfo
    
    If FormatType = Standard Then
        GetFractionSeparator = "."
    Else
        Set Nfi = NumberFormatInfo.GetInstance(Provider)
        GetFractionSeparator = CultureTable.GetString(Nfi.LCID, SNUMBERDECIMALSEPARATOR)
    End If
End Function

Private Sub AppendComponent(ByVal Component As Long, ByVal Width As Long)
    Dim ComponentString As String
    
    ComponentString = Component
    Width = Width - Len(ComponentString)
    
    If Width > 0 Then
        mOutput.AppendChar vbZeroChar, Width
    End If
    
    mOutput.AppendString ComponentString
End Sub

Private Sub AppendFraction(ByVal Width As Long, ByVal Minimize As Boolean)
    Dim CharsToChop As Long
    Dim ChoppedValue As Long
    
    CharsToChop = MaxFractionWidth - Width
    ChoppedValue = mFraction
    
    Do While CharsToChop
        ChoppedValue = ChoppedValue \ 10
        CharsToChop = CharsToChop - 1
    Loop
    
    If Minimize Then
        Do While ChoppedValue Mod 10 = 0 And ChoppedValue > 0
            ChoppedValue = ChoppedValue \ 10
            Width = Width - 1
        Loop
        
        If ChoppedValue = 0 Then
            Exit Sub
        End If
    End If
    
    AppendComponent ChoppedValue, Width
End Sub



