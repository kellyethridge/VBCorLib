VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int64Parser"
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
' Module: Int64Parser
'

Option Explicit

Private mError As ParseError


Friend Sub ThrowError()
    Select Case mError
        Case ParseError.FormatError
            Error.Format Format_InvalidString
        Case ParseError.OverflowError
            Throw Cor.NewOverflowException(Environment.GetResourceString(Overflow_Int64))
    End Select
End Sub

Friend Function TryParse(ByRef s As String, ByVal Style As NumberStyles, ByVal Provider As IFormatProvider, ByRef Result As Int64) As Boolean
    Dim Info As NumberFormatInfo
    Dim HasPrefix As Boolean
    
    mError = ParseError.None
    Result.LowPart = 0
    Result.HighPart = 0
    
    If IsHexSpecifier(s, Style, HasPrefix) Then
        TryParse = TryParseHex(s, HasPrefix, Style, Result)
    Else
        Set Info = NumberFormatInfo.GetInstance(Provider)
        TryParse = TryParseDecimal(s, Style, Info, Result)
    End If
    
    If Not TryParse Then
        Result.LowPart = 0
        Result.HighPart = 0
    End If
End Function

Private Function IsHexSpecifier(ByRef s As String, ByVal Style As NumberStyles, ByRef HasPrefix As Boolean) As Boolean
    Select Case Left$(s, 2)
        Case "&h", "&H", "0x", "0X"
            HasPrefix = True
            IsHexSpecifier = True
        Case Else
            IsHexSpecifier = (Style And NumberStyles.AllowHexSpecifier)
    End Select
End Function

Private Function TryParseDecimal(ByRef s As String, ByVal Style As NumberStyles, ByVal Info As NumberFormatInfo, ByRef n As Int64) As Boolean
    Const NotFound          As Long = 0
    Const FoundBefore       As Long = 1
    Const FoundAfter        As Long = 2
    Dim Chars()             As Integer
    Dim Index               As Long
    Dim MaxIndex            As Long
    Dim Ch                  As Long
    Dim InWhitespace        As Boolean
    Dim NeedsClosingParen   As Boolean
    Dim NegativeSign        As Integer
    Dim PositiveSign        As Integer
    Dim CurrencySymbol      As Integer
    Dim Thousands           As Integer
    Dim CurrencyThousands   As Integer
    Dim DecimalPoint        As Integer
    Dim CurrencyPoint       As Integer
    Dim FoundDigit          As Boolean
    Dim FoundCurrency       As Long
    Dim FoundSign           As Long
    Dim FoundDecimal        As Boolean
    Dim FoundClosingParen   As Boolean
    Dim FoundExponent       As Boolean
    Dim Exponent            As Long
    Dim DecimalPlaces       As Long
    Dim Char64              As Int64
    Dim IsNegative          As Boolean
    Dim IsMaxValue          As Boolean
    
    NegativeSign = AscW(Info.NegativeSign)
    PositiveSign = AscW(Info.PositiveSign)
    CurrencySymbol = AscW(Info.CurrencySymbol)
    DecimalPoint = AscW(Info.NumberDecimalSeparator)
    CurrencyPoint = AscW(Info.CurrencyDecimalSeparator)
    Thousands = AscW(Info.NumberGroupSeparator)
    CurrencyThousands = AscW(Info.CurrencyGroupSeparator)
    Chars = AllocChars(s)
    MaxIndex = UBound(Chars)
    
    Do While Index <= MaxIndex
        Ch = Chars(Index)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                If (FoundCurrency = FoundAfter) Or (FoundSign = FoundAfter) Or InWhitespace Then
                    GoTo FormatError
                End If
                
                If Not FoundDecimal Then
                    If Not CheckedInPlaceMultiplyBy10(n) Then
                        GoTo Finally
                    End If
                    
                    ' After multiplying by ten we want to simply
                    ' add the digit to the current number. However,
                    ' there is a possibility that the result could
                    ' create a number higher than allowed without
                    ' the addition carrying into the next column.
                    '
                    ' Ex. the max positive value is 9,223,372,036,854,775,807.
                    ' if we successfully multiplied by ten and have a current
                    ' value of 9,223,372,036,854,775,800 we cannot allow an
                    ' 8 or 9 to be added, so we check and overflow if necessary.
                    If n.HighPart = &H7FFFFFFF Then
                        If n.LowPart = &HFFFFFFF8 Then
                            If IsNegative Then
                                If Ch = vbEightChar Then
                                    IsMaxValue = True
                                    GoTo Continue
                                ElseIf Ch = vbNineChar Then
                                    GoTo OverflowError
                                End If
                            Else
                                If Ch = vbEightChar Or Ch = vbNineChar Then
                                    GoTo OverflowError
                                End If
                            End If
                        End If
                    End If
                    
                    InPlaceAdd n, Ch - vbZeroChar
                Else
                    If Style And NumberStyles.AllowExponent Then
                        ' we are allowing a possible exponent value to be parsed,
                        ' so if we have encountered a decimal point, we need to
                        ' continue parsing digits and also count how many
                        ' decimal places we have gone.
                        If Not CheckedInPlaceMultiplyBy10(n) Then
                            GoTo Finally
                        End If
                        
                        InPlaceAdd n, Ch - vbZeroChar
                        DecimalPlaces = DecimalPlaces + 1
                    Else
                        If Ch <> vbZeroChar Then
                            GoTo FormatError
                        End If
                    End If
                End If
                
                FoundDigit = True
            Case NegativeSign, PositiveSign
                If (FoundSign <> NotFound) Or NeedsClosingParen Then
                    GoTo FormatError
                End If
                
                If Not FoundDigit Then
                    If (Style And NumberStyles.AllowLeadingSign) = 0 Then
                        GoTo FormatError
                    End If
                    
                    FoundSign = FoundBefore
                Else
                    If (Style And NumberStyles.AllowTrailingSign) = 0 Then
                        GoTo FormatError
                    End If
                    
                    FoundSign = FoundAfter
                End If
                
                If Ch = NegativeSign Then
                    IsNegative = True
                End If
            Case vbTabChar, vbLfChar, vbVtChar, vbFfChar, vbCrChar, vbSpaceChar
                If Not FoundDigit Then
                    If ((Style And NumberStyles.AllowLeadingWhite) = 0) Or NeedsClosingParen Then
                        GoTo FormatError
                    End If
                
                    InWhitespace = FoundSign
                Else
                    InWhitespace = True
                End If
            Case vbOpenParenChar
                If ((Style And NumberStyles.AllowParentheses) = 0) Or (FoundSign <> NotFound) Then
                    GoTo FormatError
                End If
                
                NeedsClosingParen = True
            Case vbCloseParenChar
                If (Style And NumberStyles.AllowParentheses) = 0 Or (Not NeedsClosingParen) Then
                    GoTo FormatError
                End If
                
                IsNegative = True
                NeedsClosingParen = False
                FoundClosingParen = True
            Case CurrencySymbol
                If ((Style And NumberStyles.AllowCurrencySymbol) = 0) Or (FoundCurrency <> NotFound) Then
                    GoTo FormatError
                End If
                
                FoundCurrency = IIfLong(Statics.Int64.IsZero(n), FoundBefore, FoundAfter)
            Case DecimalPoint, CurrencyPoint
                If ((Style And NumberStyles.AllowDecimalPoint) = 0) Or FoundDecimal Or FoundClosingParen Then
                    GoTo FormatError
                End If
                
                If FoundCurrency Then
                    If Ch = DecimalPoint Then
                        GoTo FormatError
                    End If
                End If
                
                FoundDecimal = True
            Case Thousands, CurrencyThousands
                If ((Style And NumberStyles.AllowThousands) = 0) Or (Not FoundDigit) Then
                    GoTo FormatError
                End If
                
                If FoundCurrency Then
                    If Ch = Thousands Then
                        GoTo FormatError
                    End If
                End If
            Case vbLowerEChar, vbUpperEChar
                If ((Style And NumberStyles.AllowExponent) = 0) Or FoundExponent Then
                    GoTo FormatError
                End If
            
                If Not TryParseExponent(Chars, Index, Exponent) Then
                    GoTo FormatError
                End If
                
                ' we understand that parsing the exponent will advance the index passed the end of the digits
                ' and that we will be advancing the index as well, in effect advancing it twice, so we need
                ' to counter that last advancement.
                Index = Index - 1
                FoundExponent = True
            Case Else
                GoTo FormatError
        End Select
        
Continue:
        Index = Index + 1
    Loop
    
    If InWhitespace Then
        If (Style And NumberStyles.AllowTrailingWhite) = 0 Then
            GoTo FormatError
        End If
    End If
    
    If NeedsClosingParen Then
        GoTo FormatError
    End If
    
    If FoundExponent Then
        If Exponent < DecimalPlaces Then
            GoTo FormatError
        End If
    
        Do While Exponent - DecimalPlaces > 0
            CheckedInPlaceMultiplyBy10 n
            Exponent = Exponent - 1
        Loop
    End If
    
    If IsNegative Then
        If IsMaxValue Then
            n.HighPart = &H80000000
            n.LowPart = 0
        Else
            n = Statics.Int64.Negate(n)
        End If
    End If
    
    TryParseDecimal = True
    GoTo Finally
    
FormatError:
    mError = FormatError
    GoTo Finally
OverflowError:
    mError = OverflowError
Finally:
    FreeChars Chars
End Function

Private Function CheckedInPlaceMultiplyBy10(ByRef n As Int64) As Boolean
    Dim Words As QWordArray
    Dim k As Long
    Dim i As Long
    
    AssignQWord n, Words
    
    For i = 0 To 3
        k = (Words.Words(i) And &HFFFF&) * 10 + (((k And &HFFFF0000) \ vbShift16Bits) And &HFFFF&)
        
#If Release Then
        Words.Words(i) = k And &HFFFF&
#Else
        Words.Words(i) = AsWord(k)
#End If
    Next
    
    AssignQWord Words, n
    
    If k And &HFFFF8000 Then
        mError = ParseError.OverflowError
    Else
        CheckedInPlaceMultiplyBy10 = True
    End If
End Function

Private Sub InPlaceAdd(ByRef n As Int64, ByVal Value As Long)
#If Release Then
    Dim LowPart As Long
        
    Debug.Assert Value >= 0 And Value <= 9
    
    LowPart = n.LowPart + Value
    
    If (LowPart - &H80000000) < (n.LowPart - &H80000000) Then
        n.HighPart = n.HighPart + 1
    End If
    
    n.LowPart = LowPart
#Else
    AsCurr(n) = AsCurr(n) + AsCurr(Value)
#End If
End Sub

Private Function TryParseExponent(ByRef Chars() As Integer, ByRef Index As Long, ByRef Exponent As Long) As Boolean
    Dim Ch          As Integer
    Dim MaxIndex    As Long
    
    MaxIndex = UBound(Chars)
    Index = Index + 1
    
    If Index > MaxIndex Then
        Exit Function
    End If
    
    If Chars(Index) = vbPlusChar Then
        Index = Index + 1
    End If
    
    Do While Index <= MaxIndex
        Ch = Chars(Index)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                Exponent = Exponent * 10 + (Ch - vbZeroChar)
                TryParseExponent = True
            Case Else
                Exit Function
        End Select
        
        Index = Index + 1
    Loop
End Function

Private Function TryParseHex(ByRef s As String, ByVal HasPrefix As Boolean, ByVal Style As NumberStyles, ByRef n As Int64) As Boolean
    Const vbLowerHexOffset As Long = 87
    Const vbUpperHexOffset As Long = 55
    
    Dim Chars()     As Integer
    Dim FoundDigits As Boolean
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Ch As Long
    Dim Section As Long
    Dim MinIndex As Long
    
    MinIndex = IIfLong(HasPrefix, 2, 0)
    Chars = AllocChars(s)
    j = 1
    Section = 1
    
    For i = UBound(Chars) To MinIndex Step -1
        Select Case Chars(i)
            Case vbZeroChar
                If Section > 4 Then
                    GoTo Continue
                End If
                
                Ch = 0
                FoundDigits = True
            Case vbOneChar To vbNineChar
                Ch = Chars(i) - vbZeroChar
                FoundDigits = True
            Case vbLowerAChar To vbLowerFChar
                Ch = Chars(i) - vbLowerHexOffset
                FoundDigits = True
            Case vbUpperAChar To vbUpperFChar
                Ch = Chars(i) - vbUpperHexOffset
                FoundDigits = True
            Case vbTabChar, vbLfChar, vbVtChar, vbFfChar, vbCrChar, vbSpaceChar
                If ((Style And NumberStyles.AllowLeadingWhite) = 0) And FoundDigits Then
                    GoTo FormatError
                End If

                If ((Style And NumberStyles.AllowTrailingWhite) = 0) And Not FoundDigits Then
                    GoTo FormatError
                End If

                GoTo Continue
            Case Else
                GoTo FormatError
        End Select

        k = k Or (Ch * j)
        j = j * vbShift4Bits

        If j = vbShift16Bits Then
            Select Case Section
                Case 1
                    n.LowPart = k
                Case 2
                    n.LowPart = n.LowPart Or ShiftLeftInt32(k, 16)
                Case 3
                    n.HighPart = k
                Case 4
                    n.HighPart = n.HighPart Or ShiftLeftInt32(k, 16)
                Case Else
                    GoTo OverflowError
            End Select
            
            Section = Section + 1
            j = 1
            k = 0
        End If

Continue:
    Next i

    If k Then
        Select Case Section
            Case 1
                n.LowPart = k
            Case 2
                n.LowPart = n.LowPart Or ShiftLeftInt32(k, 16)
            Case 3
                n.HighPart = k
            Case 4
                n.HighPart = n.HighPart Or ShiftLeftInt32(k, 16)
            Case Else
                GoTo OverflowError
        End Select
    End If
    TryParseHex = True
    
FormatError:
    mError = ParseError.FormatError
    GoTo Finally
OverflowError:
    mError = ParseError.OverflowError
Finally:
    FreeChars Chars
End Function

