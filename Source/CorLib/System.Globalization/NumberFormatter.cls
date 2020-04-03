VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: NumberFormatter
'
Option Explicit
Private Const UnknownPrecision          As Long = -1
Private Const DefaultIntegerPrecision   As Long = 10
Private Const DefaultDoublePrecision    As Long = 15
Private Const DefaultSinglePrecision    As Long = 7
Private Const DefaultExponentPrecision  As Long = 6

Private Const SCALE_NAN                 As Long = &H80000000
Private Const SCALE_INF                 As Long = &H7FFFFFFF

Private Const NumberBufferSize          As Long = 512
Private Const DefaultOutputSize         As Long = 1024

Private Enum FormatSpecifier
    CustomSpecifier = -1
    GeneralUpperSpecifier = 71      ' G
    GeneralLowerSpecifier = 103     ' g
    DecimalUpperSpecifier = 68      ' D
    DecimalLowerSpecifier = 100     ' d
    NumberUpperSpecifier = 78       ' N
    NumberLowerSpecifier = 110      ' n
    HexUpperSpecifier = 88          ' X
    HexLowerSpecifier = 120         ' x
    ExponentUpperSpecifier = 69     ' E
    ExponentLowerSpecifier = 101    ' e
    FixedUpperSpecifier = 70        ' F
    FixedLowerSpecifier = 102       ' f
    CurrencyUpperSpecifier = 67     ' C
    CurrencyLowerSpecifier = 99     ' c
    PercentUpperSpecifier = 80      ' P
    PercentLowerSpecifier = 112     ' p
    CharSpecifier = vbDollarChar    ' $
End Enum

Private Type Specifiers
    Format      As FormatSpecifier
    Precision   As Long
End Type

' This structure is derived from using the ECVT function.
' The ECVT function converts an IEEE Double to its 3 components
' of digits, precision, and sign.
Private Type NumberType
    DecimalPosition As Long
    Precision       As Long
    IsNegative      As BOOL
    Digits()        As Byte
    DigitsSA        As SafeArray1d
End Type

Private Type OutputBuffer
    Buffer() As Integer
    Position As Long
End Type

Private mNumberBufferPtr    As Long
Private mNumber             As NumberType
Private mDefaultOutput      As OutputBuffer
Private mFormatInfo         As NumberFormatInfo


Friend Function Format(ByRef Value As Variant, ByRef FormatText As String, ByVal Provider As IFormatProvider) As String
    Set mFormatInfo = NumberFormatInfo.GetInstance(Provider)
    
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            Format = Int32ToString(Value, FormatText)
        Case vbDouble
            Format = DoubleToString(Value, FormatText)
        Case vbSingle
            Format = SingleToString(Value, FormatText)
        Case vbDecimal
            Format = FormatVBDecimal(Value, FormatText)
        Case Else
            Format = Value
    End Select
End Function

Friend Function FormatInt64(ByRef Value As Int64, ByRef Format As String, ByVal Provider As IFormatProvider) As String
    Dim Specifiers As Specifiers
    
    Set mFormatInfo = NumberFormatInfo.GetInstance(Provider)
    ClearOutput
    Specifiers = ParseSpecifiers(Format)
    FormatInt64 = Int64ToString(Value, Specifiers, Format)
End Function

Friend Function FormatBigNumber(ByRef Value As BigNumber, ByRef Format As String, ByVal Provider As IFormatProvider) As String
    Const MaxDigitsPerElement       As Long = 5
    Const ExponentCeremonyLength    As Long = 12
    Const DigitsPerCommaSeparator   As Long = 3
    Dim Specifiers As Specifiers
    Dim MaxLength  As Long
    
    Set mFormatInfo = NumberFormatInfo.GetInstance(Provider)
    ClearOutput
    Specifiers = ParseSpecifiers(Format)
    
    MaxLength = Value.Precision * MaxDigitsPerElement + ExponentCeremonyLength + (Value.Precision \ DigitsPerCommaSeparator)
    
    If MaxLength > Len1D(mDefaultOutput.Buffer) Then
        Dim LargeOutput As OutputBuffer
        ReDim LargeOutput.Buffer(0 To MaxLength)
        FormatBigNumber = BigNumberToString(Value, Specifiers, Format, LargeOutput)
    Else
        FormatBigNumber = BigNumberToString(Value, Specifiers, Format, mDefaultOutput)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitNumber
    InitOutput
End Sub

Private Sub Class_Terminate()
    ReleaseNumber
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitNumber()
    mNumberBufferPtr = CoTaskMemAlloc(NumberBufferSize)
    
    If mNumberBufferPtr = vbNullPtr Then _
        Throw New OutOfMemoryException
    
    mNumber.DigitsSA.cbElements = 1
    mNumber.DigitsSA.cDims = 1
    SAPtr(mNumber.Digits) = VarPtr(mNumber.DigitsSA)
End Sub

Private Sub InitNumberBuffer()
    InitCustomNumberBuffer mNumberBufferPtr, NumberBufferSize
End Sub

Private Sub InitCustomNumberBuffer(ByVal BufferPtr As Long, ByVal Size As Long)
    mNumber.DigitsSA.PVData = BufferPtr
    mNumber.DigitsSA.cElements = Size
End Sub

Private Sub ReleaseNumber()
    SAPtr(mNumber.Digits) = vbNullPtr
    CoTaskMemFree mNumberBufferPtr
End Sub

Private Sub InitOutput()
    ReDim mDefaultOutput.Buffer(0 To DefaultOutputSize - 1)
End Sub

Private Function Int32ToString(ByRef Value As Variant, ByRef Format As String) As String
    Dim Specifiers As Specifiers
    
    Specifiers = ParseSpecifiers(Format)
    
    If Specifiers.Format = CharSpecifier Then
        Int32ToString = ChrW$(AsWord(CLng(Value)))
        Exit Function
    End If
    
    If Specifiers.Format <> CustomSpecifier And _
        Specifiers.Format <> HexUpperSpecifier And _
        Specifiers.Format <> HexLowerSpecifier Then
        Int32ToNumber Value
    End If
    
    Select Case Specifiers.Format
        Case FormatSpecifier.GeneralLowerSpecifier
            FormatGeneralInteger Specifiers.Precision, vbLowerEChar, mDefaultOutput
        Case FormatSpecifier.GeneralUpperSpecifier
            FormatGeneralInteger Specifiers.Precision, vbUpperEChar, mDefaultOutput
        Case FormatSpecifier.FixedUpperSpecifier, FormatSpecifier.FixedLowerSpecifier
            FormatFixedInt32 Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.DecimalUpperSpecifier, FormatSpecifier.DecimalLowerSpecifier
            FormatDecimal Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.NumberUpperSpecifier, FormatSpecifier.NumberLowerSpecifier
            FormatNumber Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.CurrencyUpperSpecifier, FormatSpecifier.CurrencyLowerSpecifier
            FormatCurrency Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.ExponentLowerSpecifier
            FormatExponentInt32 Specifiers.Precision, vbLowerEChar, mDefaultOutput
        Case FormatSpecifier.ExponentUpperSpecifier
            FormatExponentInt32 Specifiers.Precision, vbUpperEChar, mDefaultOutput
        Case FormatSpecifier.PercentUpperSpecifier, FormatSpecifier.PercentLowerSpecifier
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            If Specifiers.Precision = UnknownPrecision Then
                FormatPercent mFormatInfo.PercentDecimalDigits, mDefaultOutput
            Else
                FormatPercent Specifiers.Precision, mDefaultOutput
            End If
        Case FormatSpecifier.HexLowerSpecifier, FormatSpecifier.HexUpperSpecifier
            Int32ToString = FormatHex(Value, Specifiers)
            Exit Function
        Case Else
            If Value < 0 Then
                If InStr(1, Format, ";") = 0 Then
                    Int32ToString = mFormatInfo.NegativeSign & VBA.Format$(Abs(Value), Format)
                    Exit Function
                End If
            End If
                
            Int32ToString = VBA.Format$(Value, Format)
            Exit Function
    End Select
    
    Int32ToString = GetStringOutput
End Function

Private Function Int64ToString(ByRef Value As Int64, ByRef Specifiers As Specifiers, ByRef Format As String) As String
    If Specifiers.Format = FormatSpecifier.CharSpecifier Then
        Int64ToString = ChrW$(AsWord(Value.LowPart))
        Exit Function
    End If
    
    If Specifiers.Format <> CustomSpecifier And _
        Specifiers.Format <> HexUpperSpecifier And _
        Specifiers.Format <> HexLowerSpecifier Then
        Int64ToNumber Value
    End If
    
    Select Case Specifiers.Format
        Case FormatSpecifier.GeneralLowerSpecifier
            FormatGeneralInteger 19, vbLowerEChar, mDefaultOutput
        Case FormatSpecifier.GeneralUpperSpecifier
            FormatGeneralInteger 19, vbUpperEChar, mDefaultOutput
        Case FormatSpecifier.CurrencyLowerSpecifier, FormatSpecifier.CurrencyUpperSpecifier
            FormatCurrency Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.DecimalUpperSpecifier, FormatSpecifier.DecimalLowerSpecifier
            FormatDecimal Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.ExponentLowerSpecifier
            FormatExponentInt32 Specifiers.Precision, vbLowerEChar, mDefaultOutput
        Case FormatSpecifier.ExponentUpperSpecifier
            FormatExponentInt32 Specifiers.Precision, vbUpperEChar, mDefaultOutput
        Case FormatSpecifier.FixedUpperSpecifier, FormatSpecifier.FixedLowerSpecifier
            FormatFixedDouble Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.NumberUpperSpecifier, FormatSpecifier.NumberLowerSpecifier
            FormatNumber Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.PercentUpperSpecifier, FormatSpecifier.PercentLowerSpecifier
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            If Specifiers.Precision = UnknownPrecision Then
                FormatPercent mFormatInfo.PercentDecimalDigits, mDefaultOutput
            Else
                FormatPercent Specifiers.Precision, mDefaultOutput
            End If
        Case FormatSpecifier.HexLowerSpecifier, FormatSpecifier.HexUpperSpecifier
            Int64ToString = FormatHexInt64(Value, Specifiers)
            Exit Function
            
        Case Else
            Dim DecimalValue As Variant
            
            DecimalValue = Statics.Int64.ToDecimal(Value)

            If DecimalValue < 0 Then
                If InStr(1, Format, ";") = 0 Then
                    Int64ToString = mFormatInfo.NegativeSign & VBA.Format$(Abs(DecimalValue), Format)
                    Exit Function
                End If
            End If
            
            Int64ToString = VBA.Format$(DecimalValue, Format)
            Exit Function
    End Select
    
    Int64ToString = GetStringOutput
End Function

Private Function GetStringOutput() As String
    If mDefaultOutput.Position = 0 Then
        GetStringOutput = "0"
    Else
        GetStringOutput = SysAllocStringLen(VarPtr(mDefaultOutput.Buffer(0)), mDefaultOutput.Position)
    End If
End Function

Private Function BigNumberToString(ByRef Value As BigNumber, ByRef Specifiers As Specifiers, ByRef Format As String, ByRef Output As OutputBuffer) As String
    Dim BufferPtr As Long
    
    If Specifiers.Format <> CustomSpecifier And _
        Specifiers.Format <> HexUpperSpecifier And _
        Specifiers.Format <> HexLowerSpecifier Then
        BufferPtr = BigNumberToNumber(Value)
    End If
    
    Select Case Specifiers.Format
        Case FormatSpecifier.GeneralLowerSpecifier
            FormatGeneralInteger CorMath.Max(mNumber.Precision, Specifiers.Precision), vbLowerEChar, Output, True
        Case FormatSpecifier.GeneralUpperSpecifier
            FormatGeneralInteger CorMath.Max(mNumber.Precision, Specifiers.Precision), vbUpperEChar, Output, True
        Case FormatSpecifier.CurrencyLowerSpecifier, FormatSpecifier.CurrencyUpperSpecifier
            FormatCurrency Specifiers.Precision, Output
        Case FormatSpecifier.DecimalUpperSpecifier, FormatSpecifier.DecimalLowerSpecifier
            FormatDecimal Specifiers.Precision, Output
        Case FormatSpecifier.ExponentLowerSpecifier
            FormatExponentInt32 Specifiers.Precision, vbLowerEChar, Output
        Case FormatSpecifier.ExponentUpperSpecifier
            FormatExponentInt32 Specifiers.Precision, vbUpperEChar, Output
        Case FormatSpecifier.FixedUpperSpecifier, FormatSpecifier.FixedLowerSpecifier
            FormatFixedDouble Specifiers.Precision, Output
        Case FormatSpecifier.NumberUpperSpecifier, FormatSpecifier.NumberLowerSpecifier
            FormatNumber Specifiers.Precision, Output
        Case FormatSpecifier.PercentUpperSpecifier, FormatSpecifier.PercentLowerSpecifier
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            If Specifiers.Precision = UnknownPrecision Then
                FormatPercent mFormatInfo.PercentDecimalDigits, Output
            Else
                FormatPercent Specifiers.Precision, Output
            End If
        Case FormatSpecifier.HexLowerSpecifier, FormatSpecifier.HexUpperSpecifier
            BigNumberToString = FormatHexBigNumber(Value, Specifiers, Output)
            Exit Function
            
        Case Else
            BigNumberToString = Format
            Exit Function
    End Select
    
    CoTaskMemFree BufferPtr
    
    If Output.Position = 0 Then
        BigNumberToString = "0"
    Else
        BigNumberToString = SysAllocStringLen(VarPtr(Output.Buffer(0)), Output.Position)
    End If
End Function

Private Function FormatVBDecimal(ByRef Value As Variant, ByRef Format As String) As String
    Dim Specifiers As Specifiers
    
    ClearOutput
    Specifiers = ParseSpecifiers(Format)
    FormatVBDecimal = DecimalToString(Value, Specifiers, Format)
End Function

Private Function DecimalToString(ByRef Value As Variant, ByRef Specifiers As Specifiers, ByRef Format As String) As String
    Dim ExpChar As Integer
    
    If Specifiers.Format = HexLowerSpecifier Or Specifiers.Format = HexUpperSpecifier Then _
        Error.Format Format_BadFormatSpecifier
    
    If Specifiers.Format <> CustomSpecifier Then
        DecimalToNumber Value
    End If
    
    Select Case Specifiers.Format
        Case FormatSpecifier.GeneralLowerSpecifier, FormatSpecifier.GeneralUpperSpecifier
            ExpChar = IIfLong(Specifiers.Format = FormatSpecifier.GeneralLowerSpecifier, vbLowerEChar, vbUpperEChar)
        
            If Specifiers.Precision = UnknownPrecision Or Specifiers.Precision = 0 Then
                Specifiers.Precision = 29
            End If
        
            RoundNumber Specifiers.Precision
            FormatGeneralDouble Specifiers.Precision, -29, ExpChar, mDefaultOutput
        Case FormatSpecifier.CurrencyLowerSpecifier, FormatSpecifier.CurrencyUpperSpecifier
            FormatCurrency Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.DecimalUpperSpecifier, FormatSpecifier.DecimalLowerSpecifier
            FormatDecimal Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.ExponentLowerSpecifier
            FormatExponentInt32 Specifiers.Precision, vbLowerEChar, mDefaultOutput
        Case FormatSpecifier.ExponentUpperSpecifier
            FormatExponentInt32 Specifiers.Precision, vbUpperEChar, mDefaultOutput
        Case FormatSpecifier.FixedUpperSpecifier, FormatSpecifier.FixedLowerSpecifier
            FormatFixedDouble Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.NumberUpperSpecifier, FormatSpecifier.NumberLowerSpecifier
            FormatNumber Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.PercentUpperSpecifier, FormatSpecifier.PercentLowerSpecifier
            If Specifiers.Precision = UnknownPrecision Then
                Specifiers.Precision = mFormatInfo.PercentDecimalDigits
            End If
            
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            RoundNumber Specifiers.Precision + mNumber.DecimalPosition
            FormatPercent Specifiers.Precision, mDefaultOutput
        Case Else
            If Value < 0 Then
                If InStr(1, Format, ";") = 0 Then
                    DecimalToString = mFormatInfo.NegativeSign & VBA.Format$(Abs(Value), Format)
                    Exit Function
                End If
            End If
            
            DecimalToString = VBA.Format$(Value, Format)
            Exit Function
    End Select
    
    If mDefaultOutput.Position = 0 Then
        DecimalToString = "0"
    Else
        DecimalToString = SysAllocStringLen(VarPtr(mDefaultOutput.Buffer(0)), mDefaultOutput.Position)
    End If
End Function

Private Sub DecimalToNumber(ByRef Value As Variant)
    Const MaxDecimalPrecision As Long = 29
    Dim DigitIndex  As Long
    Dim IsMaxNeg    As Boolean
    Dim Bits()      As Integer
    Dim PtrSrc      As Long
    
    ReDim Bits(0 To 5)
    InitNumberBuffer
    mNumber.IsNegative = IIfLong(Value < 0, BOOL_TRUE, BOOL_FALSE)
    PtrSrc = VarPtr(Value)
    
    ' this would be an extremely rare occurrence since a decimal
    ' datatype can only exist within a variant variable. The passed
    ' in parameter would literally need to be a second variant
    ' variable containing a reference to a decimal variant variable.
    If VariantType(Value) And VT_BYREF Then
        ' if a value was passed by reference we need to
        ' dereference the memory address stored in the
        ' passed in variant value.
        PtrSrc = MemLong(PtrSrc + VARIANTDATA_OFFSET)
    End If
    
    ' VB Decimal layout
    ' bytes 0-1: Variant datatype information
    ' bytes 2-3: precision and sign bytes
    ' bytes 4-7: the 32 highest bits in the 96bit value
    ' bytes 8-11: the 32 lowest bits in the 96bit value
    ' bytes 12-15: the 32 middle bits in the 96bit value
    AssignQWord ByVal PtrSrc + 8, Bits(0) ' map VB lowest and middle bits to .NET lowest and middle bits
    AssignDWord ByVal PtrSrc + 4, Bits(4) ' map VB highest bits to .NET highest bits
            
    Do While ((Bits(0) Or Bits(1) Or Bits(2) Or Bits(3) Or Bits(4) Or Bits(5)) <> 0) And (DigitIndex < MaxDecimalPrecision)
        mNumber.Digits(DigitIndex) = InPlaceDivideBy10(Bits) + vbZeroChar
        DigitIndex = DigitIndex + 1
    Loop
    
    mNumber.Digits(DigitIndex) = 0
    mNumber.DecimalPosition = DigitIndex - MemByte(PtrSrc + 2)
    mNumber.Precision = DigitIndex
    
    If mNumber.Precision > 0 Then
        Dim i As Long
        Dim j As Long
        Dim t As Byte
        
        j = mNumber.Precision - 1
        
        For i = 0 To (mNumber.Precision \ 2) - 1
            t = mNumber.Digits(i)
            mNumber.Digits(i) = mNumber.Digits(j)
            mNumber.Digits(j) = t
            j = j - 1
        Next
    End If
End Sub

Private Function SingleToString(ByVal Value As Single, ByRef Format As String) As String
    Dim Specifiers As Specifiers
    
    Specifiers = ParseSpecifiers(Format)

    Select Case Specifiers.Format
        Case FormatSpecifier.GeneralLowerSpecifier, FormatSpecifier.GeneralUpperSpecifier, FormatSpecifier.CustomSpecifier
            ' do nothing
        Case Else
            DoubleToNumber Value, DefaultSinglePrecision
            
            If Not IsValidNumber(SingleToString) Then
                Exit Function
            End If
    End Select

    Select Case Specifiers.Format
        Case FormatSpecifier.FixedLowerSpecifier, FormatSpecifier.FixedUpperSpecifier
            FormatFixedDouble Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.ExponentLowerSpecifier, FormatSpecifier.ExponentUpperSpecifier
            FormatScientific Specifiers.Precision, Specifiers.Format, mDefaultOutput
        Case FormatSpecifier.NumberLowerSpecifier, FormatSpecifier.NumberUpperSpecifier
            FormatNumber Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.CurrencyLowerSpecifier, FormatSpecifier.CurrencyUpperSpecifier
            FormatCurrency Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.GeneralLowerSpecifier, FormatSpecifier.GeneralUpperSpecifier
            If Specifiers.Precision < 0 Then
                Specifiers.Precision = DefaultSinglePrecision
            End If
            
            DoubleToNumber Value, Specifiers.Precision
            
            If Not IsValidNumber(SingleToString) Then
                Exit Function
            End If
            
            If mNumber.Digits(0) = vbZeroChar Then
                SingleToString = "0"
                Exit Function
            End If
            
            If Specifiers.Precision > 7 Then
                Specifiers.Precision = 9
            End If
            
            RoundNumber Specifiers.Precision
            FormatGeneralDouble mNumber.Precision, -3, Specifiers.Format - 2, mDefaultOutput
        
        Case FormatSpecifier.PercentLowerSpecifier, FormatSpecifier.PercentUpperSpecifier
            If Specifiers.Precision = UnknownPrecision Then
                Specifiers.Precision = mFormatInfo.PercentDecimalDigits
            End If
            
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            RoundNumber Specifiers.Precision + mNumber.DecimalPosition
            FormatPercent Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.CustomSpecifier
            If Value < 0! Then
                If InStr(1, Format, ";") = 0 Then
                    SingleToString = mFormatInfo.NegativeSign & VBA.Format$(Abs(Value), Format)
                    Exit Function
                End If
            End If
            
            SingleToString = VBA.Format$(Value, Format)
            Exit Function
        Case Else
            Error.Format Format_BadFormatSpecifier
    End Select
    
    If mDefaultOutput.Position = 0 Then
        SingleToString = "0"
    Else
        SingleToString = SysAllocStringLen(VarPtr(mDefaultOutput.Buffer(0)), mDefaultOutput.Position)
    End If
End Function

Private Function DoubleToString(ByVal Value As Double, ByRef Format As String) As String
    Dim Specifiers As Specifiers
    
    Specifiers = ParseSpecifiers(Format)
    
    If Specifiers.Format <> GeneralLowerSpecifier And _
        Specifiers.Format <> GeneralUpperSpecifier And _
        Specifiers.Format <> CustomSpecifier Then
        DoubleToNumber Value, DefaultDoublePrecision
        
        If Not IsValidNumber(DoubleToString) Then
            Exit Function
        End If
    End If
    
    Select Case Specifiers.Format
        Case FormatSpecifier.FixedLowerSpecifier, FormatSpecifier.FixedUpperSpecifier
            FormatFixedDouble Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.ExponentLowerSpecifier, FormatSpecifier.ExponentUpperSpecifier
            FormatScientific Specifiers.Precision, Specifiers.Format, mDefaultOutput
        Case FormatSpecifier.NumberLowerSpecifier, FormatSpecifier.NumberUpperSpecifier
            FormatNumber Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.CurrencyLowerSpecifier, FormatSpecifier.CurrencyUpperSpecifier
            FormatCurrency Specifiers.Precision, mDefaultOutput
        Case FormatSpecifier.GeneralLowerSpecifier, FormatSpecifier.GeneralUpperSpecifier
            If Specifiers.Precision = UnknownPrecision Or Specifiers.Precision = 0 Then
                Specifiers.Precision = DefaultDoublePrecision
            End If
            
            DoubleToNumber Value, Specifiers.Precision
            
            If Not IsValidNumber(DoubleToString) Then Exit Function
            
            If mNumber.Digits(0) = vbZeroChar Then
                DoubleToString = "0"
                Exit Function
            End If
            
            RoundNumber Specifiers.Precision
            FormatGeneralDouble Specifiers.Precision, -3, Specifiers.Format - 2, mDefaultOutput
        
        Case FormatSpecifier.PercentLowerSpecifier, FormatSpecifier.PercentUpperSpecifier
            If Specifiers.Precision = UnknownPrecision Then
                Specifiers.Precision = mFormatInfo.PercentDecimalDigits
            End If
            
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            RoundNumber Specifiers.Precision + mNumber.DecimalPosition
            FormatPercent Specifiers.Precision, mDefaultOutput
        
        Case FormatSpecifier.CustomSpecifier
            If Value < 0# Then
                If InStr(1, Format, ";") = 0 Then
                    DoubleToString = mFormatInfo.NegativeSign & VBA.Format$(Abs(Value), Format)
                    Exit Function
                End If
            End If
            
            DoubleToString = VBA.Format$(Value, Format)
            Exit Function
        Case Else
            Error.Format Format_BadFormatSpecifier
    End Select
    
    DoubleToString = GetStringOutput
End Function

Private Sub FormatExponentInt32(ByVal FormatPrecision As Long, ByVal FormatType As Long, ByRef Output As OutputBuffer)
    Dim Places As Long
    
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = 6
    End If
    
    Places = mNumber.DecimalPosition - 1
    mNumber.DecimalPosition = 1
    
    FormatFixedDouble FormatPrecision, Output
    AppendExponent Places, FormatType, 3, Output
End Sub

Private Sub FormatNumber(ByVal FormatPrecision As Long, ByRef Output As OutputBuffer)
    Dim Pattern         As String
    Dim PatternChars()  As Integer
    
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = mFormatInfo.NumberDecimalDigits
    End If
    
    RoundNumber mNumber.DecimalPosition + FormatPrecision
    
    If mNumber.IsNegative Then
        Pattern = NumberFormatInfo.NumberNegativePatterns(mFormatInfo.NumberNegativePattern)
    Else
        Pattern = "#"
    End If
    
    PatternChars = AllocChars(Pattern)
    
    Dim i   As Long
    Dim Ch  As Integer
    With mFormatInfo
        For i = 0 To UBound(PatternChars)
            Ch = PatternChars(i)
            
            Select Case Ch
                Case vbPoundChar
                    FormatGroups FormatPrecision, .NumberGroupSizes, .NumberGroupSeparator, .NumberDecimalSeparator, Output
                Case vbMinusChar
                    AppendString .NegativeSign, Output
                Case Else
                    AppendChar Ch, Output
            End Select
        Next i
    End With
    
    FreeChars PatternChars
End Sub

Private Sub FormatCurrency(ByVal DecimalPlaces As Long, ByRef Output As OutputBuffer)
    Dim Pattern         As String
    Dim PatternChars()  As Integer
    
    If DecimalPlaces = UnknownPrecision Then
        DecimalPlaces = mFormatInfo.CurrencyDecimalDigits
    End If
    
    RoundNumber mNumber.DecimalPosition + DecimalPlaces
    
    If mNumber.IsNegative Then
        Pattern = NumberFormatInfo.CurrencyNegativePatterns(mFormatInfo.CurrencyNegativePattern)
    Else
        Pattern = NumberFormatInfo.CurrencyPositivePatterns(mFormatInfo.CurrencyPositivePattern)
    End If
    
    PatternChars = AllocChars(Pattern)
    
    Dim i   As Long
    Dim Ch  As Integer
    With mFormatInfo
        For i = 0 To UBound(PatternChars)
            Ch = PatternChars(i)
            Select Case Ch
                Case vbPoundChar
                    FormatGroups DecimalPlaces, .CurrencyGroupSizes, .CurrencyGroupSeparator, .CurrencyDecimalSeparator, Output
                Case vbMinusChar
                    AppendString .NegativeSign, Output
                Case vbDollarChar
                    AppendString .CurrencySymbol, Output
                Case Else
                    AppendChar Ch, Output
            End Select
        Next i
    End With
    
    FreeChars PatternChars
End Sub

Private Sub FormatPercent(ByVal DecimalPlaces As Long, ByRef Output As OutputBuffer)
    Dim Pattern         As String
    Dim PatternChars()  As Integer
    
    If mNumber.IsNegative Then
        Pattern = NumberFormatInfo.PercentNegativePatterns(mFormatInfo.PercentNegativePattern)
    Else
        Pattern = NumberFormatInfo.PercentPositivePatterns(mFormatInfo.PercentPositivePattern)
    End If
    
    PatternChars = AllocChars(Pattern)
    
    Dim i   As Long
    Dim Ch  As Integer
    With mFormatInfo
        For i = 0 To UBound(PatternChars)
            Ch = PatternChars(i)
            Select Case Ch
                Case vbPoundChar
                    FormatGroups DecimalPlaces, .PercentGroupSizes, .PercentGroupSeparator, .PercentDecimalSeparator, Output
                Case vbMinusChar
                    AppendString .NegativeSign, Output
                Case vbPercentChar
                    AppendString .PercentSymbol, Output
                Case Else
                    AppendChar Ch, Output
            End Select
        Next i
    End With
    
    FreeChars PatternChars
End Sub

''
' Formats the whole portion of a number with the group separator.
' This is typically a comma. The default is groups of 3 digits.
'
' @param DecimalPlaces The number of digits that should appear after the period.
' @param Groupsizes A set of group sizes used to group together digits.
' @param GroupSeparator The seperator character inserted between groups.
' @param DecimalSeparator The separator character used to separate the whole part from the fractional part.
'
Private Sub FormatGroups(ByVal DecimalPlaces As Long, ByRef GroupSizes() As Long, ByRef GroupSeparator As String, ByRef DecimalSeparator As String, ByRef Output As OutputBuffer)
    Dim DigitIndex                  As Long
    Dim NonrepeatingGroupSizeSum    As Long
    Dim Ch                          As Integer
    Dim RepeatingGroupDigitCount    As Long
    Dim DigitCount                  As Long
    Dim GroupSize                   As Long
    Dim i                           As Long
    Dim HasGroups                   As Boolean
    
    HasGroups = UBound(GroupSizes) >= LBound(GroupSizes)
    
    With mNumber
        If .DecimalPosition > 0 Then
            For i = LBound(GroupSizes) To UBound(GroupSizes) - 1
                NonrepeatingGroupSizeSum = NonrepeatingGroupSizeSum + GroupSizes(i)
            Next

            RepeatingGroupDigitCount = .DecimalPosition - NonrepeatingGroupSizeSum
            
            If RepeatingGroupDigitCount < 0 Then
                RepeatingGroupDigitCount = 0
            End If

            If HasGroups Then
                GroupSize = GroupSizes(UBound(GroupSizes))
            End If

            If RepeatingGroupDigitCount = GroupSize Then
                DigitCount = -1
            ElseIf GroupSize > 0 Then
                DigitCount = RepeatingGroupDigitCount Mod GroupSize
                
                If DigitCount = 0 Then
                    DigitCount = GroupSize
                End If
            Else
                DigitCount = RepeatingGroupDigitCount
            End If

            DigitIndex = 0
            Ch = .Digits(DigitIndex)
            
            Do While RepeatingGroupDigitCount > 0
                If DigitCount = 0 And HasGroups Then
                    AppendString GroupSeparator, Output
                    DigitCount = GroupSize
                End If
                
                If Ch = vbTerminator Then
                    AppendChar vbZeroChar, Output
                Else
                    AppendChar Ch, Output
                    Ch = .Digits(DigitIndex + 1)
                End If
                
                DigitIndex = DigitIndex + 1
                DigitCount = DigitCount - 1
                RepeatingGroupDigitCount = RepeatingGroupDigitCount - 1
            Loop
        Else
            AppendChar vbZeroChar, Output
        End If

        If (DigitIndex = 0 Or Ch <> vbTerminator) Or (DigitIndex < .DecimalPosition And Ch = vbTerminator) Then
            Dim GroupSizeIndex As Long
            
            If HasGroups Then
                GroupSizeIndex = UBound(GroupSizes) - 1
                
                If GroupSizeIndex < 0 Then
                    GroupSizeIndex = 0
                End If
                
                GroupSize = NonrepeatingGroupSizeSum - GroupSizes(GroupSizeIndex)
            End If
            
            DigitCount = .DecimalPosition - DigitIndex
            Ch = .Digits(DigitIndex)

            For i = DigitIndex To .DecimalPosition - 1
                If HasGroups Then
                    If DigitCount = NonrepeatingGroupSizeSum Then
                        AppendString GroupSeparator, Output
                    ElseIf DigitCount = GroupSize Then
                        AppendString GroupSeparator, Output
                        GroupSizeIndex = GroupSizeIndex - 1
                        
                        If GroupSizeIndex < 0 Then
                            GroupSizeIndex = 0
                        End If
                        
                        GroupSize = GroupSize - GroupSizes(GroupSizeIndex)
                    End If
                End If
                
                If Ch = 0 Then
                    AppendChar vbZeroChar, Output
                Else
                    AppendChar Ch, Output
                    Ch = .Digits(DigitIndex + 1)
                End If
                
                DigitIndex = DigitIndex + 1
                DigitCount = DigitCount - 1
            Next
        End If
        
        ' Append any digits that need to appear after a decimal point
        If DecimalPlaces > 0 Then
            AppendString DecimalSeparator, Output
            DigitIndex = .DecimalPosition
            
            ' negative scale represents numbers such as 0.000000123,
            ' so we need to insert the zeros before the digits.
            If DigitIndex < 0 Then
                AppendZeros -DigitIndex, Output
                DecimalPlaces = DecimalPlaces + DigitIndex
                DigitIndex = 0
            End If
            
            If Ch <> vbTerminator Then
                Ch = .Digits(i)
            End If
            
            Do While DecimalPlaces > 0
                If Ch = 0 Then
                    AppendChar vbZeroChar, Output
                Else
                    AppendChar Ch, Output
                    DigitIndex = DigitIndex + 1
                    Ch = .Digits(DigitIndex)
                End If
                
                DecimalPlaces = DecimalPlaces - 1
            Loop
        End If
    End With
End Sub

''
' Creates a hex output, upper or lowercase, with optional leading zeros.
'
' @param Value The value to be converted to the hex representation.
' @param digits The number of digits the output should be. This allows
' for a 16-bit integer to be only 4 digits, while 32-bit to be 8 digits.
' @param CaseOffset This is the alpha offset to upper or lower case
' alpha characters that are in the output string.
'
Private Function FormatHex(ByRef Value As Variant, ByRef Specifiers As Specifiers) As String
    Const HexLetterOffset As Long = 10
    Dim Digits      As Long
    Dim Ch          As Integer
    Dim HexValue    As Long
    Dim LetterBase  As Long
    
    ' This is to allow negative vbInteger values to
    ' remain at 4 hex characters instead of being
    ' stretched out to 8 characters like vbLong.
    If VarType(Value) = vbInteger Then
        HexValue = CLng(Value) And &HFFFF&
    Else
        HexValue = Value
    End If

    If Specifiers.Precision = UnknownPrecision Then
        Digits = 1
    Else
        Digits = Specifiers.Precision
    End If
    
    LetterBase = IIfLong(Specifiers.Format = HexLowerSpecifier, vbLowerAChar, vbUpperAChar) - HexLetterOffset
    mDefaultOutput.Position = DefaultOutputSize - 1
    
    Do While HexValue <> 0
        Ch = HexValue And &HF
        
        If Ch > 9 Then
            mDefaultOutput.Buffer(mDefaultOutput.Position) = LetterBase + Ch
        Else
            mDefaultOutput.Buffer(mDefaultOutput.Position) = vbZeroChar + Ch
        End If
        
        mDefaultOutput.Position = mDefaultOutput.Position - 1
        HexValue = ShiftRightInt32(HexValue, 4)
    Loop
    
    ' Add leading zeros if needed.
    Do While mDefaultOutput.Position >= DefaultOutputSize - Digits
        mDefaultOutput.Buffer(mDefaultOutput.Position) = vbZeroChar
        mDefaultOutput.Position = mDefaultOutput.Position - 1
    Loop
    
    mDefaultOutput.Position = mDefaultOutput.Position + 1
    FormatHex = SysAllocStringLen(VarPtr(mDefaultOutput.Buffer(mDefaultOutput.Position)), DefaultOutputSize - mDefaultOutput.Position)
End Function

Private Function FormatHexInt64(ByRef Value As Int64, ByRef Specifiers As Specifiers) As String
    Const HexLetterOffset As Long = 10
    Dim Digits      As Long
    Dim Ch          As Integer
    Dim HexValue    As Int64
    Dim LetterBase  As Long
    
    HexValue = Value

    If Specifiers.Precision = UnknownPrecision Then
        Digits = 1
    Else
        Digits = Specifiers.Precision
    End If
    
    LetterBase = IIfLong(Specifiers.Format = HexLowerSpecifier, vbLowerAChar, vbUpperAChar) - HexLetterOffset
    mDefaultOutput.Position = DefaultOutputSize - 1
    
    Do While Not Statics.Int64.IsZero(HexValue)
        Ch = HexValue.LowPart And &HF
        
        If Ch > 9 Then
            mDefaultOutput.Buffer(mDefaultOutput.Position) = LetterBase + Ch
        Else
            mDefaultOutput.Buffer(mDefaultOutput.Position) = vbZeroChar + Ch
        End If
        
        mDefaultOutput.Position = mDefaultOutput.Position - 1
        HexValue = Statics.Int64.ShiftRight(HexValue, 4)
    Loop
    
    ' Add leading zeros if needed.
    Do While mDefaultOutput.Position >= DefaultOutputSize - Digits
        mDefaultOutput.Buffer(mDefaultOutput.Position) = vbZeroChar
        mDefaultOutput.Position = mDefaultOutput.Position - 1
    Loop
    
    mDefaultOutput.Position = mDefaultOutput.Position + 1
    FormatHexInt64 = SysAllocStringLen(VarPtr(mDefaultOutput.Buffer(mDefaultOutput.Position)), DefaultOutputSize - mDefaultOutput.Position)
End Function

Private Function FormatHexBigNumber(ByRef Number As BigNumber, ByRef Specifiers As Specifiers, ByRef Output As OutputBuffer) As String
    Dim CharCount       As Long
    Dim HexCharOffset   As Long
    
    HexCharOffset = IIfLong(Specifiers.Format = HexLowerSpecifier, vbLowerAChar, vbUpperAChar) - &HA
    Output.Position = UBound(Output.Buffer)
    
    Dim k As Long
    Dim i As Long
    For i = 0 To Number.Precision - 1
        k = Number.Digits(i) And &HFFFF&
        
        Dim j As Long
        Dim Digit As Long
        For j = 1 To 4
            Digit = k And &HF
            
            If Digit < &HA Then
                Output.Buffer(Output.Position) = vbZeroChar + Digit
            Else
                Output.Buffer(Output.Position) = HexCharOffset + Digit
            End If

            k = k \ vbShift4Bits
            Output.Position = Output.Position - 1
            CharCount = CharCount + 1
        Next
    Next
    
    If Output.Position < UBound(Output.Buffer) Then
        Dim Position As Long
        Dim LastChar As Long
        
        Position = Output.Position + 1
        LastChar = Output.Buffer(Position)
        
        Select Case LastChar
            Case vbLowerFChar, vbUpperFChar, vbZeroChar
                Do While Position < UBound(Output.Buffer) And Output.Buffer(Position) = LastChar
                    Position = Position + 1
                    CharCount = CharCount - 1
                Loop
                
                If Output.Buffer(Position) >= vbEightChar And Number.Sign = 1 Then
                    Position = Position - 1
                    CharCount = CharCount + 1
                    Output.Buffer(Position) = vbZeroChar
                End If
        
                Output.Position = Position - 1
        End Select
    End If
            
    Dim MinCharCount    As Long
    Dim FillChar        As Long
    
    MinCharCount = IIfLong(Specifiers.Precision = UnknownPrecision, 1, Specifiers.Precision)
    FillChar = IIfLong(Number.Sign = -1, HexCharOffset + &HF, vbZeroChar)
    
    Do While CharCount < MinCharCount
        Output.Buffer(Output.Position) = FillChar
        Output.Position = Output.Position - 1
        CharCount = CharCount + 1
    Loop
            
    FormatHexBigNumber = SysAllocStringLen(VarPtr(Output.Buffer(Output.Position + 1)), Len1D(Output.Buffer) - Output.Position - 1)
End Function

''
' Formats an integer value with enough leading zeros as necessary
' to be as long as the number of digits specified.
'
' @param digits The number of digits to be displayed, filling the
' number with leading zeros if necessary.
'
Private Sub FormatDecimal(ByVal TotalDigits As Long, ByRef Output As OutputBuffer)
    With mNumber
        ' Add the negative sign if necessary.
        If .IsNegative Then
            AppendString mFormatInfo.NegativeSign, Output
        End If
        
        Dim ZerosToAdd As Long
        ZerosToAdd = TotalDigits - .Precision
        AppendZeros ZerosToAdd, Output
        
        ' Fill the rest of the buffer with the digits.
        Dim i As Long
        For i = 0 To .Precision - 1
            AppendChar .Digits(i), Output
        Next i
    End With
End Sub

''
' This will format the common number to a 0.000e+000 type format.
'
' @param digits The number of digits represented before the rest
' are represented by the exponent.
' @param expchar The exponent character "e" or "E".
' @remarks This will take the common number and format it into the
' common output buffer in scientific notation.
'
Private Sub FormatScientific(ByVal Digits As Long, ByVal ExpChar As Integer, ByRef Output As OutputBuffer)
    If Digits = UnknownPrecision Then
        Digits = DefaultExponentPrecision
    End If
    
    RoundNumber Digits + 1
    
    With mNumber
        If .IsNegative Then
            AppendString mFormatInfo.NegativeSign, Output
        End If
        
        ' We need to represent "digits" number of digits in the precision.
        ' After the first digit is ouptput we put the decimal separator, since
        ' all scientific notation has a leading digit-period-digits format.
        Dim i           As Long
        Dim Ch          As Integer
        Dim OutOfDigits As Boolean
        
        Do While Digits >= 0
            If i = 1 Then
                AppendString mFormatInfo.NumberDecimalSeparator, Output
            End If
            
            If Not OutOfDigits Then
                Ch = .Digits(i)
            End If
            
            ' Simply copy the digit to the output buffer or place
            ' a 0 in the output buffer if we have run out of digits
            ' but required more digits to be displayed.
            If Ch = 0 Then
                AppendChar vbZeroChar, Output
                OutOfDigits = True
            Else
                AppendChar Ch, Output
            End If
            
            i = i + 1
            Digits = Digits - 1
        Loop
        
        AppendExponent .DecimalPosition - 1, ExpChar, 3, Output
    End With
End Sub

''
' Rounds a number by hand to avoid banker's rounding.
'
' @param digits The number of digits to round to.
' @remarks Since we already have the number in the common buffer of
' digits, we can simply check each digit directly for rounding.
'
Private Sub RoundNumber(ByVal Digits As Long)
    With mNumber
        If Digits > .Precision Then
            Exit Sub
        End If
    
        ' are we rounding up or down?
        If .Digits(Digits) >= vbFiveChar Then
            ' We don't round 9's directly, since they would just
            ' carry over, so we keep skipping them until there
            ' are no more. Then we'll round up the first non-nine digit.
            ' This will also lower the precision, since we had
            ' to keep dealing with the carry.
            Do While Digits > 0
                If .Digits(Digits - 1) <> vbNineChar Then
                    Exit Do
                End If
                
                Digits = Digits - 1
            Loop
            
            ' If we didn't reach the leftmost digit, then round
            ' the digit 1-left of us up.
            If Digits > 0 Then
                .Digits(Digits - 1) = .Digits(Digits - 1) + 1
            Else
                ' otherwise we ran out of digits, so we have to
                ' scale up by 1 and add a 1. ie. 9999 -> 10000
                .DecimalPosition = .DecimalPosition + 1
                .Digits(Digits) = vbOneChar
                Digits = 1
            End If
        Else
            ' rounding down while each digit to the left = 0, since
            ' we don't need those 0 anymore.
            Do While Digits > 0
                If .Digits(Digits - 1) <> vbZeroChar Then
                    Exit Do
                End If
                
                Digits = Digits - 1
            Loop
        End If
        
        If Digits = 0 Then
            .DecimalPosition = 0
            .IsNegative = BOOL_FALSE
        End If
        
        .Digits(Digits) = vbNullPtr
    End With
End Sub

Private Function BigNumberToNumber(ByRef Value As BigNumber) As Long
    Const MaxDigitsPerInteger As Long = 5
    Dim Number  As BigNumber
    Dim Index   As Long
    Dim MaxDigits As Long
    
    Select Case Value.Sign
        Case -1
            BigNumberMath.Negate Value, Number
            mNumber.IsNegative = BOOL_TRUE
        Case Else
            Number = Value
            mNumber.IsNegative = BOOL_FALSE
    End Select
    
    MaxDigits = Number.Precision * MaxDigitsPerInteger + 1
    
    If MaxDigits > 500 Then
        BigNumberToNumber = CoTaskMemAlloc(MaxDigits)
        InitCustomNumberBuffer BigNumberToNumber, MaxDigits
    Else
        InitNumberBuffer
    End If
    
    Index = UBound(mNumber.Digits) - 1
    
    Do While Number.Sign <> 0
        Dim Digit As Long
        Digit = BigNumberMath.SingleInPlaceDivideBy10(Number)
        mNumber.Digits(Index) = Digit + vbZeroChar
        Index = Index - 1
    Loop
    
    mNumber.Precision = UBound(mNumber.Digits) - Index - 1
    mNumber.DecimalPosition = mNumber.Precision
    CopyMemory mNumber.Digits(0), mNumber.Digits(Index + 1), mNumber.Precision + 1
End Function

''
' This will parse the format type and optional numerical value.
'
' @param Format The string format to be parsed (ie. "D99").
' @param Specifiers The structure to place the parse results.
' @remarks A format of "D99" will return DecimalFormat in Specifiers.Format and
' 99 in Specifiers.Precision. We indicate no digits were supplied by using -1.
Private Function ParseSpecifiers(ByRef Format As String) As Specifiers
    Dim FormatChars()   As Integer
    Dim FormatLength    As Long
    
    FormatLength = Len(Format)
    
    Select Case FormatLength
        Case 0
            ParseSpecifiers.Format = GeneralUpperSpecifier
            ParseSpecifiers.Precision = UnknownPrecision
        Case Is > 3
            ParseSpecifiers.Format = CustomSpecifier
            ParseSpecifiers.Precision = UnknownPrecision
        Case Else
            FormatChars = AllocChars(Format)
            
            Select Case FormatChars(0)
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    ParseSpecifiers.Format = FormatChars(0)
                    
                    If FormatLength = 1 Then
                        ParseSpecifiers.Precision = UnknownPrecision
                        GoTo Finally
                    End If
                    
                    Dim Precision   As Long
                    Dim Ch          As Long
                    Dim i           As Long
                    For i = 1 To UBound(FormatChars)
                        Ch = FormatChars(i)
                        
                        Select Case Ch
                            Case vbZeroChar To vbNineChar
                                Precision = Precision * 10 + Ch - vbZeroChar
                            Case Else
                                ParseSpecifiers.Format = CustomSpecifier
                                ParseSpecifiers.Precision = UnknownPrecision
                                GoTo Finally
                        End Select
                    Next i
                    
                    ParseSpecifiers.Precision = Precision
                Case Else
                    If Len(Format) = 1 And Asc(Format) = vbDollarChar Then
                        ParseSpecifiers.Format = CharSpecifier
                    Else
                        ParseSpecifiers.Format = CustomSpecifier
                        ParseSpecifiers.Precision = UnknownPrecision
                    End If
            End Select
    End Select
    
Finally:
    FreeChars FormatChars
End Function

''
' Takes an integer value and makes a character string.
'
' @param Value The value to convert to a string.
' @remarks We do this by hand because we need the digits
' to be placed in a common buffer. Converting the value to
' a string then copying that string to the buffer takes a
' significantly longer time, believe it or not.
'
Private Sub Int32ToNumber(ByVal Value As Long)
    Dim IsMaxNeg    As Boolean
    Dim DigitIndex  As Long
    Dim OneTenth    As Long
    
    ClearOutput
    InitNumberBuffer
    
    With mNumber
        If Value < 0 Then
            If Value = &H80000000 Then
                ' We will overflow if we try to get the absolute
                ' value of the minimum (maximum?) negative number, so add one
                ' and set a flag indicating what we have done.
                Value = -(Value + 1)
                IsMaxNeg = True
            Else
                Value = -Value
            End If
            .IsNegative = BOOL_TRUE
        Else
            .IsNegative = BOOL_FALSE
        End If
        
        ' We determine the number of characters first because we
        ' will be adding the characters in reverse order to the
        ' number buffer. This allows us to build the string up
        ' backwards. Otherwise we would have to reverse the buffer
        ' once the string of digits has been placed in it.
        Select Case Value
            Case 0:                 DigitIndex = 0
            Case Is < 10:           DigitIndex = 1
            Case Is < 100:          DigitIndex = 2
            Case Is < 1000:         DigitIndex = 3
            Case Is < 10000:        DigitIndex = 4
            Case Is < 100000:       DigitIndex = 5
            Case Is < 1000000:      DigitIndex = 6
            Case Is < 10000000:     DigitIndex = 7
            Case Is < 100000000:    DigitIndex = 8
            Case Is < 1000000000:   DigitIndex = 9
            Case Else:              DigitIndex = 10
        End Select
                
        .Precision = DigitIndex         ' Precision is the number of digits represented.
        .DecimalPosition = DigitIndex   ' DecimalPosition is the placement of the decimal point from the left.
                                        ' Zero would be before the first digit.
                                    
        .Digits(DigitIndex) = vbNullPtr  ' terminating character in the buffer.
        
        ' chop off each digit from the right and
        ' append it into a string buffer.
        Do While Value > 0
            DigitIndex = DigitIndex - 1
            OneTenth = Value \ 10
            
            ' This is about 2x faster than using Mod.
            .Digits(DigitIndex) = (Value - (OneTenth * 10)) + vbZeroChar
            Value = OneTenth
        Loop
        
        If IsMaxNeg Then
            .Digits(.Precision - 1) = vbEightChar
        End If
    End With
End Sub

Private Sub Int64ToNumber(ByRef Value As Int64)
    Dim DigitIndex  As Long
    Dim Words       As QWordArray
    Dim IsMaxNeg    As Boolean
    
    InitNumberBuffer
    
    If Value.HighPart < 0 Then
        mNumber.IsNegative = BOOL_TRUE
        
        If Value.HighPart = &H80000000 And Value.LowPart = 0 Then
            IsMaxNeg = True
            Words.Words(0) = &HFFFF
            Words.Words(1) = &HFFFF
            Words.Words(2) = &HFFFF
            Words.Words(3) = &H7FFF
        Else
            AssignQWord Statics.Int64.Negate(Value), Words
        End If
    Else
        mNumber.IsNegative = BOOL_FALSE
        AssignQWord Value, Words
    End If
        
    Do While (Words.Words(0) Or Words.Words(1) Or Words.Words(2) Or Words.Words(3)) <> 0
        mNumber.Digits(DigitIndex) = InPlaceDivideBy10(Words.Words) + vbZeroChar
        DigitIndex = DigitIndex + 1
    Loop
    
    mNumber.Digits(DigitIndex) = vbNullPtr
    mNumber.DecimalPosition = DigitIndex
    mNumber.Precision = DigitIndex
    
    If mNumber.Precision > 0 Then
        Dim i As Long
        Dim j As Long
        Dim t As Byte
        
        j = mNumber.Precision - 1
        
        For i = 0 To (mNumber.Precision \ 2) - 1
            t = mNumber.Digits(i)
            mNumber.Digits(i) = mNumber.Digits(j)
            mNumber.Digits(j) = t
            j = j - 1
        Next

        If IsMaxNeg Then
            mNumber.Digits(mNumber.Precision - 1) = vbEightChar
        End If
    End If
End Sub

Private Function InPlaceDivideBy10(ByRef Bits() As Integer) As Long
    Dim r As Long
    Dim i As Long
    Dim d As Long

    For i = UBound(Bits) To 0 Step -1
        r = (r * vbShift16Bits) + (Bits(i) And &HFFFF&)
        d = r \ 10

#If Release Then
        Bits(i) = d And &HFFFF&
#Else
        Bits(i) = AsWord(d)
#End If

        r = r - (d * 10)
    Next

    InPlaceDivideBy10 = r
End Function

Private Sub DoubleToNumber(ByVal Value As Double, ByVal FormatPrecision As Long)
    ClearOutput
    
    With AsDLong(Value)
        If (.HiDWord And &H7FF00000) = &H7FF00000 Then
            InitNumberBuffer
        
            If (.HiDWord And &HFFFFF) Or (.LoDWord) Then
                mNumber.DecimalPosition = SCALE_NAN
            Else
                mNumber.DecimalPosition = SCALE_INF
                mNumber.IsNegative = (Value < 0#)
            End If
        Else
            InitCustomNumberBuffer Helper.Ecvt(Value, FormatPrecision, mNumber.DecimalPosition, mNumber.IsNegative), FormatPrecision + 1
            mNumber.Precision = IIfLong(mNumber.Digits(0) = vbZeroChar, 0, FormatPrecision)
        End If
    End With
End Sub

Private Sub FormatGeneralInteger(ByVal FormatPrecision As Long, ByVal ExpChar As Long, ByRef Output As OutputBuffer, Optional ByVal IsBigInteger As Boolean = False)
    Dim MaxPrecision As Long
    Dim i As Long
    
    If Not IsBigInteger Then
        If FormatPrecision = UnknownPrecision Or FormatPrecision = 0 Then
            FormatPrecision = DefaultIntegerPrecision
        End If
    End If
    
    If mNumber.IsNegative Then
        AppendString mFormatInfo.NegativeSign, Output
    End If
    
    ' If the requested precision is less than the
    ' actuall number of digits, then we have to
    ' go the scientific format route.
    Dim HasExponent As Boolean
    If FormatPrecision < mNumber.Precision Then
        MaxPrecision = FormatPrecision
        mNumber.DecimalPosition = 1
        RoundNumber FormatPrecision
        HasExponent = True
    Else
        If IsBigInteger Then
            If FormatPrecision > mNumber.Precision Then
                For i = 1 To FormatPrecision - mNumber.Precision
                    AppendChar vbZeroChar, Output
                Next
            End If
        End If
        
        MaxPrecision = mNumber.Precision
    End If
    
    For i = 0 To MaxPrecision - 1
        If i = mNumber.DecimalPosition Then
            AppendString mFormatInfo.NumberDecimalSeparator, Output
        End If
        
        AppendChar mNumber.Digits(i), Output
    Next i
    
    If HasExponent Then
        AppendExponent mNumber.Precision - 1, ExpChar, 2, Output
    End If
End Sub

Private Sub FormatGeneralDouble(ByVal FormatPrecision As Long, ByVal ExpThreshold As Long, ByVal ExpChar As Integer, ByRef Output As OutputBuffer)
    Dim Ch As Integer
    Dim DP As Long
    Dim HasExponent As Boolean
    Dim i As Long
    
    With mNumber
        If .IsNegative Then
            AppendString mFormatInfo.NegativeSign, Output
        End If
        
        DP = .DecimalPosition
        
        If FormatPrecision < .DecimalPosition Or .DecimalPosition < ExpThreshold Then
            HasExponent = True
            DP = 1
        ElseIf DP <= 0 Then
            AppendChar vbZeroChar, Output
            DP = 0
        End If
        
        Ch = .Digits(i)
        
        Do While Ch <> 0
            If i = DP Then
                AppendString mFormatInfo.NumberDecimalSeparator, Output
                
                If Not HasExponent Then
                    AppendZeros DP - .DecimalPosition, Output
                End If
            End If
            
            AppendChar Ch, Output
            i = i + 1
            Ch = .Digits(i)
        Loop
        
        If HasExponent Then
            AppendExponent .DecimalPosition - 1, ExpChar, 2, Output
        Else
            AppendZeros .DecimalPosition - i, Output
        End If
    End With
End Sub

Private Sub FormatFixedInt32(ByVal FormatPrecision As Long, ByRef Output As OutputBuffer)
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = mFormatInfo.NumberDecimalDigits
    End If
    
    With mNumber
        If .IsNegative Then
            AppendString mFormatInfo.NegativeSign, Output
        End If
        
        Dim i As Long
        If .Precision = 0 Then
            AppendChar vbZeroChar, Output
        Else
            For i = 0 To .Precision - 1
                AppendChar .Digits(i), Output
            Next i
        End If
        
        If FormatPrecision > 0 Then
            AppendString mFormatInfo.NumberDecimalSeparator, Output
            AppendZeros FormatPrecision, Output
        End If
    End With
End Sub

Private Sub FormatFixedDouble(ByVal FormatPrecision As Long, ByRef Output As OutputBuffer)
    Dim i   As Long
    Dim Ch  As Integer
        
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = mFormatInfo.NumberDecimalDigits
    End If
        
    With mNumber
        RoundNumber FormatPrecision + .DecimalPosition
        
        If .IsNegative Then
            AppendString mFormatInfo.NegativeSign, Output
        End If
        
        If .Precision > 0 And .DecimalPosition > 0 Then
            For i = 0 To .DecimalPosition - 1
                AppendChar .Digits(i), Output
            Next i
        Else
            AppendChar vbZeroChar, Output
            i = .DecimalPosition
        End If
        
        If FormatPrecision > 0 Then
            AppendString mFormatInfo.NumberDecimalSeparator, Output
            
            Do While FormatPrecision > 0
                If i >= 0 Then
                    Ch = .Digits(i)
                End If
                
                If .Precision = 0 Or Ch = 0 Then
                    AppendChar vbZeroChar, Output
                    
                    If i < 0 Then
                        i = i + 1
                    End If
                Else
                    AppendChar Ch, Output
                    i = i + 1
                End If

                FormatPrecision = FormatPrecision - 1
            Loop
        End If
    End With
End Sub

Private Function IsValidNumber(ByRef RetVal As String) As Boolean
    Select Case mNumber.DecimalPosition
        Case SCALE_NAN
            RetVal = mFormatInfo.NaNSymbol
        Case SCALE_INF
            If mNumber.IsNegative Then
                RetVal = mFormatInfo.NegativeInfinitySymbol
            Else
                RetVal = mFormatInfo.PositiveInfinitySymbol
            End If
        Case Else
            IsValidNumber = True
    End Select
End Function

''
' This adds the 'e+000' portion of a scientific formatted output.
'
Private Sub AppendExponent(ByVal Places As Long, ByVal ExpChar As Integer, ByVal MinDigits As Long, ByRef Output As OutputBuffer)
    AppendChar ExpChar, Output
    
    If mNumber.Precision = 0 Then
        Places = 0
    End If
    
    If Places < 0 Then
        Places = -Places
        AppendString mFormatInfo.NegativeSign, Output
    Else
        AppendString mFormatInfo.PositiveSign, Output
    End If
    
    ' Determine how many characters is in Places.
    Dim Digits As Long
    Select Case Places
        Case 0:         Digits = 0
        Case Is < 10:   Digits = 1
        Case Is < 100:  Digits = 2
        Case Else:      Digits = 3
    End Select
    
    ' To make sure we have the MinDigits of characters in
    ' the exponent string, we pad with zeros to make up for
    ' for Places not having enough digits.
    AppendZeros MinDigits - Digits, Output
    
    ' Move our buffer pointer out to the end.
    Output.Position = Output.Position + Digits
    
    ' We add the right-most character to the buffer. We add
    ' them in reverse order, starting from the rightside and
    ' moving back towards the left. This is the easiest way
    ' mathematically without having to reverse the characters.
    Dim pos As Long
    pos = Output.Position - 1
    
    Do While Places > 0
        Dim OneTenth As Long
        OneTenth = Places \ 10
        Output.Buffer(pos) = (Places - (OneTenth * 10)) + vbZeroChar
        Places = OneTenth
        pos = pos - 1
    Loop
End Sub

Private Sub ClearOutput()
    mDefaultOutput.Position = 0
End Sub

Private Sub AppendChar(ByVal Char As Integer, ByRef Buffer As OutputBuffer)
    Buffer.Buffer(Buffer.Position) = Char
    Buffer.Position = Buffer.Position + 1
End Sub

Private Sub AppendZeros(ByVal Count As Long, ByRef Buffer As OutputBuffer)
    Dim i As Long
    
    If Count < 1 Then
        Exit Sub
    End If
    
    For i = Buffer.Position To Buffer.Position + Count - 1
        Buffer.Buffer(i) = vbZeroChar
    Next i
    
    Buffer.Position = Buffer.Position + Count
End Sub

Private Sub AppendString(ByRef s As String, ByRef Buffer As OutputBuffer)
    Dim Chars() As Integer
    Dim i       As Long
    Dim Length  As Long
    
    Length = Len(s)
    
    Select Case Length
        Case 1
            Buffer.Buffer(Buffer.Position) = AscW(s)
        Case Is < 6
            Chars = AllocChars(s)
            
            For i = 0 To Length - 1
                Buffer.Buffer(Buffer.Position + i) = Chars(i)
            Next
            
            FreeChars Chars
        Case Else
            CopyMemory Buffer.Buffer(Buffer.Position), ByVal StrPtr(s), Length * vbSizeOfChar
    End Select
    
    Buffer.Position = Buffer.Position + Length
End Sub

