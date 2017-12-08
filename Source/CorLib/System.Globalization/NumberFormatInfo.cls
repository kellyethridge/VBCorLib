VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: NumberFormatInfo
'

''
' Provides numeric formatting information.
'
' @remarks<br>
' An instance of this class can be used to provide formatting for numeric data
' in places that allow for an <b>IFormatProvider</b> to be supplied. StringBuilder.AppendFormat
' is an example of when a <b>NumberFormatInfo</b> object can be used. Also, an
' instance of <b>NumberFormatInfo</b> can be used stand-alone to format numeric data
' by calling the <b>Format</b> method.
' <p>To format a number using <b>NumberFormatInfo</b>, call the <b>Format</b> method passing
' in the value to be formated.
' @see NumberFormatInfoStatic
' @see IFormatProvider
' @see ICloneable
' @include "..\DocIncludes\NumberFormatInfo.html"
Option Explicit
Implements IObject
Implements IFormatProvider
Implements ICloneable

Private Const PROP_POSITIVESIGN                 As String = "PositiveSign"
Private Const PROP_NEGATIVESIGN                 As String = "NegativeSign"
Private Const PROP_NUMBERDECIMALSEPARATOR       As String = "NumberDecimalSeparator"
Private Const PROP_NUMBERDECIMALDIGITS          As String = "NumberDecimalDigits"
Private Const PROP_NUMBERGROUPSIZES             As String = "NumberGroupSizes"
Private Const PROP_NUMBERNEGATIVEPATTERN        As String = "NumberNegativePattern"
Private Const PROP_NUMBERGROUPSEPARATOR         As String = "NumberGroupSeparator"
Private Const PROP_CURRENCYDECIMALSEPARATOR     As String = "CurrencyDecimalSeparator"
Private Const PROP_CURRENCYDECIMALDIGITS        As String = "CurrencyDecimalDigits"
Private Const PROP_CURRENCYGROUPSIZES           As String = "CurrencyGroupSizes"
Private Const PROP_CURRENCYNEGATIVEPATTERN      As String = "CurrencyNegativePattern"
Private Const PROP_CURRENcYGROUPSEPARATOR       As String = "CurrencyGroupSeparator"
Private Const PROP_CURRENCYSYMBOL               As String = "CurrencySymbol"
Private Const PROP_CURRENCYPOSITIVEPATTERN      As String = "CurrencyPositivePattern"
Private Const PROP_PERCENTDECIMALSEPARATOR      As String = "PercentDecimalSeparator"
Private Const PROP_PERCENTDECIMALDIGITS         As String = "PercentDecimalDigits"
Private Const PROP_PERCENTGROUPSIZES            As String = "PercentGroupSizes"
Private Const PROP_PERCENTNEGATIVEPATTERN       As String = "PercentNegativePattern"
Private Const PROP_PERCENTGROUPSEPARATOR        As String = "PercentGroupSeparator"
Private Const PROP_PERCENTSYMBOL                As String = "PercentSymbol"
Private Const PROP_PERCENTPOSITIVEPATTERN       As String = "PercentPositivePattern"
Private Const PROP_PERMILLESYMBOL               As String = "PerMilleSymbol"
Private Const PROP_NANSYMBOL                    As String = "NaNSymbol"
Private Const PROP_POSITIVEINFINITYSYMBOL       As String = "PositiveInfinitySymbol"
Private Const PROP_NEGATIVEINFINITYSYMBOL       As String = "NegativeInfinitySymbol"
Private Const PROP_ISREADONLY                   As String = "IsReadOnly"

Private Const UnknownPrecision          As Long = -1
Private Const DefaultIntegerPrecision   As Long = 10
Private Const DefaultDoublePrecision    As Long = 15
Private Const DefaultSinglePrecision    As Long = 7
Private Const DefaultExponentPrecision  As Long = 6

Private Const CURRENCY_FORMAT       As Long = 67
Private Const DECIMAL_FORMAT        As Long = 68
Private Const EXPONENT_FORMAT       As Long = 69
Private Const FIXED_FORMAT          As Long = 70
Private Const GENERAL_FORMAT        As Long = 71
Private Const NUMBER_FORMAT         As Long = 78
Private Const PERCENT_FORMAT        As Long = 80
Private Const HEX_FORMAT            As Long = 88

Private Const NO_DIGITS             As Long = -1

Private Const SCALE_NAN             As Long = &H80000000
Private Const SCALE_INF             As Long = &H7FFFFFFF

Private Const BUFFER_SIZE           As Long = 1024


''
' Output patterns for negative numbers.
'
' @param [(n)] Negative value inside perenthesis. Negative symbol is not included.
' @param [-n] Negative symbol before the value.
' @param [- n] Negative symbol before the value, separated by a space.
' @param [n-] Negative symbol after the value.
' @param [n -] Negative symbol after the value, separated by a space.
Public Enum NumberNegativePatterns
    [(n)] = 0
    [-n] = 1
    [- n] = 2
    [n-] = 3
    [n -] = 4
End Enum

''
' Output patterns for negative currency values. The <strong>$</strong> is the currency
' symbol, then <strong>-</strong> is the negative symbol and <strong>n</strong> is the number.
'
' @param [($n)]
' @param [-$n]
' @param [$-n]
' @param [$n-]
' @param [(n$)]
' @param [-n$]
' @param [n-$]
' @param [n$-]
' @param [-n $]
' @param [-$ n]
' @param [n $-]
' @param [$ n-]
' @param [$ -n]
' @param [n- $]
' @param [($ n)]
' @param [(n $)]
Public Enum CurrencyNegativePatterns
    [($n)] = 0
    [-$n] = 1
    [$-n] = 2
    [$n-] = 3
    [(n$)] = 4
    [-n$] = 5
    [n-$] = 6
    [n$-] = 7
    [-n $] = 8
    [-$ n] = 9
    [n $-] = 10
    [$ n-] = 11
    [$ -n] = 12
    [n- $] = 13
    [($ n)] = 14
    [(n $)] = 15
End Enum

''
' Output patterns for positive currency values. The <strong>$</strong> is the currency symbol and <strong>n</strong> is the number.
'
' @param [$n]
' @param [n$]
' @param [$ n]
' @param [n $]
Public Enum CurrencyPositivePatterns
    [$n] = 0
    [n$] = 1
    [$ n] = 2
    [n $] = 3
End Enum

''
' Output patterns for positive percentage values. The <strong>#</strong> is the number and the <strong>%</strong> is the percentage symbol.
'
' @param [n %]
' @param [n%]
' @param [%n]
' @param [% n]
Public Enum PercentPositivePatterns
    [n %] = 0
    [n%] = 1
    [%n] = 2
    [% n] = 3
End Enum

''
' Output patterns for negative percentage values. The <b>-</b> is the negative symbol,
' <strong>#</strong> is the number and the <strong>%</strong> is the percentage symbol.
'
' @param [-n %]
' @param [-n%]
' @param [-%n]
' @param [%-n]
' @param [%n-]
' @param [n-%]
' @param [n%-]
' @param [-% n]
' @param [n %-]
' @param [% n-]
' @param [% -n]
' @param [n- %]
Public Enum PercentNegativePatterns
    [-n %] = 0
    [-n%] = 1
    [-%n] = 2
    [%-n] = 3
    [%n-] = 4
    [n-%] = 5
    [n%-] = 6
    [-% n] = 7
    [n %-] = 8
    [% n-] = 9
    [% -n] = 10
    [n- %] = 11
End Enum

' These are the format types supported by this
' class. They are the uppercase of the format type
' letters used in the format string.
'
Private Enum FormatType
    CustomFormat = -1
    GeneralFormat = 71      ' G
    DecimalFormat = 68      ' D
    NumberFormat = 78       ' N
    HexFormat = 88          ' X
    ExponentFormat = 69     ' E
    FixedFormat = 70        ' F
    CurrencyFormat = 67     ' C
    PercentFormat = 80      ' P
End Enum

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
End Enum

Private Type Specifiers
    Format      As FormatSpecifier
    Precision   As Long
End Type

' This structure is derived from using the ECVT function.
' The ECVT function converts an IEEE Double to its 3 components
' of digits, precision, and sign.
'
Private Type NumberType
    DecimalPosition As Long
    Precision       As Long
    IsNegative      As BOOL
    Digits()        As Byte
    DigitsSA        As SafeArray1d
End Type

Private Type PropsType
    PositiveSign                As String
    NegativeSign                As String
    NumberDecimalSeparator      As String
    NumberDecimalDigits         As Long
    NumberGroupSizes()          As Long
    NumberNegativePattern       As NumberNegativePatterns
    NumberGroupSeparator        As String
    CurrencyDecimalSeparator    As String
    CurrencyDecimalDigits       As Long
    CurrencyGroupSizes()        As Long
    CurrencyNegativePattern     As CurrencyNegativePatterns
    CurrencyGroupSeparator      As String
    CurrencySymbol              As String
    CurrencyPositivePattern     As CurrencyPositivePatterns
    PercentDecimalSeparator     As String
    PercentDecimalDigits        As Long
    PercentGroupSizes()         As Long
    PercentNegativePattern      As PercentNegativePatterns
    PercentGroupSeparator       As String
    PercentSymbol               As String
    PercentPositivePattern      As PercentPositivePatterns
    PerMilleSymbol              As String
    NaNSymbol                   As String
    PositiveInfinitySymbol      As String
    NegativeInfinitySymbol      As String
    IsReadOnly                  As Boolean
End Type

' sets of string representation of the different output patterns.
Private mCurrencyPositivePatterns() As String
Private mCurrencyNegativePatterns() As String
Private mNumberNegativePatterns()   As String
Private mPercentPositivePatterns()  As String
Private mPercentNegativePatterns()  As String

' utility variables
Private mPtrNumberBuffer    As Long
Private mNumber             As NumberType
Private mAppendStringBuffer As CharBuffer

' this is the set of NumberFormatInfo properties.
Private mProps  As PropsType
Private mLoaded As Boolean


' we'll assume every numeric output will fit in 1024 characters.
Private mOutput(BUFFER_SIZE - 1)    As Integer
Private mOutputPos                  As Long


' used to parse the format specifier components.
Private mFormatSpecifier As CharBuffer


''
' Returns if this instance is read-only.
'
' @return Value indicating if this instance is read-only.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

Friend Sub MarkReadOnly()
    mProps.IsReadOnly = True
End Sub

''
' Returns a clone of this instance.
'
' @return A clone of the original instance.
' @remarks The <b>IsReadOnly</b> is passed to the clone, aswell.
Public Function Clone() As NumberFormatInfo
    Set Clone = New NumberFormatInfo
    Clone.InitClone mProps
End Function

''
' Returns a string representation of negative infinty.
'
' @return Negative infinity string.
' @remarks Negative infinity is achieved by dividing a negative
' number by zero. The default is "-Infinity".
Public Property Get NegativeInfinitySymbol() As String
    VerifyLoaded
    NegativeInfinitySymbol = mProps.NegativeInfinitySymbol
End Property

''
' Sets the negative infinity representation.
'
' @param RHS The new string that represents negative infinity.
' @remarks Negative infinity is achieved by dividing a negative
' number by zero. The default is "-Infinity".
Public Property Let NegativeInfinitySymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NegativeInfinitySymbol = RHS
End Property

''
' Returns a string representation of positive infinity.
'
' @return Positive infinity string.
' @remarks Positive infinity is achieved by dividing a positive
' number by zero. The default is "Infinity".
Public Property Get PositiveInfinitySymbol() As String
    VerifyLoaded
    PositiveInfinitySymbol = mProps.PositiveInfinitySymbol
End Property

''
' Sets the positive infinity representation.
'
' @param RHS The new string representation of positive infinity.
' @remarks Positive infinity is achieved by dividing a positive
' number by zero. The default is "Infinity".
Public Property Let PositiveInfinitySymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PositiveInfinitySymbol = RHS
End Property

''
' Returns the permille representation.
'
' @return Permille string representation.
' @remarks From <a href="http://mathworld.wolfram.com/Permil.html">MathWorld</a><br>
' The use of permille (a.k.a. parts per thousand) is a way of
' expressing ratios in terms of whole numbers. Given a ratio or fraction,
' it is converted to a permil-age by multiplying by 1000 and appending a "‰" (Unicode &h2030).
' For example, if an investment grows from a number P=13.00 to a number A=22.50,
' then A is 22.50/13.00=1.7308 times as much as P, or 1730.8‰.
Public Property Get PerMilleSymbol() As String
    VerifyLoaded
    PerMilleSymbol = mProps.PerMilleSymbol
End Property

''
' Sets the permille representation.
'
' @param RHS The new permille string representation.
' @remarks From From <a href="http://mathworld.wolfram.com/Permil.html">MathWorld</a><br>
' The use of permille (a.k.a. parts per thousand) is a way of
' expressing ratios in terms of whole numbers. Given a ratio or fraction,
' it is converted to a permil-age by multiplying by 1000 and appending a "‰" (Unicode &H2030).
' For example, if an investment grows from a number P=13.00 to a number A=22.50,
' then A is 22.50/13.00=1.7308 times as much as P, or 1730.8‰.
Public Property Let PerMilleSymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PerMilleSymbol = RHS
End Property

''
' Returns the string representation for Not-A-Number.
'
' @return The string representation.
' @remarks Not-A-Number is achieved by dividing zero by zero. The default
' NaN symbol is "NaN".
Public Property Get NaNSymbol() As String
    VerifyLoaded
    NaNSymbol = mProps.NaNSymbol
End Property

''
' Sets the string representation for Not-A-Number.
'
' @param RHS The new string NaN representation.
' @remarks Not-A-Number is achieved by dividing zero by zero. The default
' NaN symbol is "NaN".
Public Property Let NaNSymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NaNSymbol = RHS
End Property

''
' Returns the decimal symbol for percent formatted numbers.
'
' @return The symbol for decimal separation.
Public Property Get PercentDecimalSeparator() As String
    PercentDecimalSeparator = mProps.PercentDecimalSeparator
End Property

''
' Sets the decimal symbol for percent formatted numbers.
'
' @param RHS The new decimal symbol.
Public Property Let PercentDecimalSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentDecimalSeparator = RHS
End Property

''
' Returns the number of digits to be displayed after the
' decimal point for percent formatted numbers.
'
' @return The number of digits to display after the decimal point.
' @remarks <p>This is used when a specifier of "p" is used with no
' associated digits (ie "p5").</p>
' <p>If the number is less than the precision
' of the value, then the value is rounded to fit within the number of
' digits specified. If omitted the value is rounded to the number of
' digits returned by this property.</p>
' <p>If the specified number of digits is larger than the precision of
' value, then the formatted value will be padded with zeros.</p>
' <p>The default number of digits is 2.</p>
Public Property Get PercentDecimalDigits() As Long
    VerifyLoaded
    PercentDecimalDigits = mProps.PercentDecimalDigits
End Property

''
' Sets the number of digits after the decimal point for percent formatted numbers.
'
' @param RHS The number of digits to display after the decimal point.
' @remarks <p>This is used when a specifier of "p" is used with no
' associated digits (ie "p5").</p>
' <p>If the number is less than the precision
' of the value, then the value is rounded to fit within the number of
' digits specified. If omitted the value is rounded to the number of
' digits returned by this property.</p>
' <p>If the specified number of digits is larger than the precision of
' value, then the formatted value will be padded with zeros.</p>
' <p>The default number of digits is 2.</p>
Public Property Let PercentDecimalDigits(ByVal RHS As Long)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99))
        
    mProps.PercentDecimalDigits = RHS
End Property

''
' Returns the current group sizes array for percent formatted numbers.
'
' @return An array containing grouping sizes for a percent formatted number.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Get PercentGroupSizes() As Variant
    VerifyLoaded
    PercentGroupSizes = mProps.PercentGroupSizes
End Property

''
' Sets the group sizes for percent formatted numbers.
'
' @param RHS A vbLong array containing the grouping sizes.
' @remarks A group size is the number of digits that are to be groupsed together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Let PercentGroupSizes(ByRef RHS As Variant)
    VerifyWritable
    VerifyLoaded
    If CorArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException("PercentGroupSizes", Environment.GetResourceString(ArgumentNull_Array))
    
    VerifyGroupSizes RHS
    mProps.PercentGroupSizes = RHS
End Property

''
' Returns the selected negative pattern used in percent number formatting.
'
' @return Then negative number pattern to be used in formating negative numbers.
'
Public Property Get PercentNegativePattern() As PercentNegativePatterns
    VerifyLoaded
    PercentNegativePattern = mProps.PercentNegativePattern
End Property

''
' Sets the negative pattern to be used in formatting negative percent numbers.
'
' @param RHS The pattern to be used.
'
Public Property Let PercentNegativePattern(ByVal RHS As PercentNegativePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.PercentNegativePattern = RHS
End Property

''
' Returns the group separator for a percent formatted number.
'
' @return The group separator symbol.
' @remarks Generally a group separator is a comman ','. However,
' different cultures may use different symbols.
'
Public Property Get PercentGroupSeparator() As String
    VerifyLoaded
    PercentGroupSeparator = mProps.PercentGroupSeparator
End Property

''
' Sets the group separator for a percent formatted number.
'
' @param RHS The symbol used as a group separator defined by the group sizes.
'
Public Property Let PercentGroupSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentGroupSeparator = RHS
End Property

''
' Returns the percent symbol used in percent formatted numbers.
'
' @return The percent symbole, usually '%'.
'
Public Property Get PercentSymbol() As String
    VerifyLoaded
    PercentSymbol = mProps.PercentSymbol
End Property

''
' Set the percent symbol used in percent formatted numbers.
'
' @param RHS The symbol to be used in percent formatted numbers.
'
Public Property Let PercentSymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentSymbol = RHS
End Property

''
' Returns the pattern to be used in formatting positive percent numbers.
'
' @return The pattern used in formatting positive percent numbers.
'
Public Property Get PercentPositivePattern() As PercentPositivePatterns
    VerifyLoaded
    PercentPositivePattern = mProps.PercentPositivePattern
End Property

''
' Sets the pattern to be used in formatting positive percent numbers.
'
' @param RHS The pattern to be used.
'
Public Property Let PercentPositivePattern(ByVal RHS As PercentPositivePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.PercentPositivePattern = RHS
End Property

''
' Returns the decimal separator used in formatting numbers into currency.
'
' @return The decimal separator character.
'
Public Property Get CurrencyDecimalSeparator() As String
    VerifyLoaded
    CurrencyDecimalSeparator = mProps.CurrencyDecimalSeparator
End Property

''
' Sets the decimal separator used in formatting numbers into currency.
'
' @param RHS The decimal separator to be used in formatting.
'
Public Property Let CurrencyDecimalSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyDecimalSeparator = RHS
End Property

''
' Returns the default number of digits after the decimal separator.
'
' @return The default number of digits after the decimal separator.
'
Public Property Get CurrencyDecimalDigits() As Long
    VerifyLoaded
    CurrencyDecimalDigits = mProps.CurrencyDecimalDigits
End Property

''
' Sets the default number of digits to be shown after the decimal separator.
'
' @param RHS The default number of digits to be shown after the decimal separator.
'
Public Property Let CurrencyDecimalDigits(ByVal RHS As Long)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99))
    mProps.CurrencyDecimalDigits = RHS
End Property

''
' Returns an array of longs which represent a comma separator group size.
'
' @return An array of longs.
' @remarks <p>Each element represents the number of digits that will be grouped
' together before a comma separator is inserted. Each element is used once
' to specify the group size, then the next element is used for the next group
' of digits, until the last group size. The last group size is used for the
' remainder of digits to be formatted.</p>
' <p>If a 0 is specified for a group size, it can only be the last size in the array.
' A 0 as the last group size will prevent the remaining digits from being separated.</p>
'
Public Property Get CurrencyGroupSizes() As Long()
    VerifyLoaded
    CurrencyGroupSizes = mProps.CurrencyGroupSizes
End Property

''
' Sets an array of longs which represent a set of group sizes before being comma separated.
'
' @param RHS An array of longs that are the group sizes.
' @remarks <p>Each element represents the number of digits that will be grouped
' together before a comma separator is inserted. Each element is used once
' to specify the group size, then the next element is used for the next group
' of digits, until the last group size. The last group size is used for the
' remainder of digits to be formatted.</p>
' <p>If a 0 is specified for a group size, it can only be the last size in the array.
' A 0 as the last group size will prevent the remaining digits from being separated.</p>
'
Public Property Let CurrencyGroupSizes(ByRef RHS() As Long)
    VerifyWritable
    VerifyLoaded
    If CorArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException("CurrencyGroupSizes", Environment.GetResourceString(ArgumentNull_Array))
    
    VerifyAndSetGroupSizes RHS, mProps.CurrencyGroupSizes
End Property

''
' The numeric pattern to be used when formatting negative values.
'
' @return The negative pattern used.
'
Public Property Get CurrencyNegativePattern() As CurrencyNegativePatterns
    VerifyLoaded
    CurrencyNegativePattern = mProps.CurrencyNegativePattern
End Property

''
' Sets the numeric pattern to be used when formatting negative values.
'
' @param RHS The pattern to be used when formatting.
'
Public Property Let CurrencyNegativePattern(ByVal RHS As CurrencyNegativePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyNegativePattern = RHS
End Property

''
' Returns the string to be use as a group separator.
'
' @return The group separator string.
' @remarks Generally a comma is the group separator. However, in some
' parts of the world, a comma is not used.
'
Public Property Get CurrencyGroupSeparator() As String
    VerifyLoaded
    CurrencyGroupSeparator = mProps.CurrencyGroupSeparator
End Property

''
' Sets the group separator to be used when separating digits.
'
' @param RHS The separator to be used when separating digits.
' @remarks Generally a comma is the group separator. However, in some
' parts of the world, a comma is not used.
'
Public Property Let CurrencyGroupSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyGroupSeparator = RHS
End Property

''
' Returns the currency symbol.
'
' @return The currency symbol.
' @remarks In the US the symbol is the dollar sign '$', in other parts
' of the world, other symbols are used.
'
Public Property Get CurrencySymbol() As String
    VerifyLoaded
    CurrencySymbol = mProps.CurrencySymbol
End Property

''
' Sets the currency symbol to be used.
'
' @param RHS The currency symbol to be used when formatting currency.
' @remarks In the US the symbol is the dollar sign '$', in other parts
' of the world, other symbols are used.
'
Public Property Let CurrencySymbol(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencySymbol = RHS
End Property

''
' Returns the formatting pattern to be used for positive currency values.
'
' @return The formatting pattern for positive currency values.
'
Public Property Get CurrencyPositivePattern() As CurrencyPositivePatterns
    VerifyLoaded
    CurrencyPositivePattern = mProps.CurrencyPositivePattern
End Property

''
' Sets the formatting pattern to be used for positive currency values.
'
' @param RHS The formatting pattern to be used with positive currency values.
'
Public Property Let CurrencyPositivePattern(ByVal RHS As CurrencyPositivePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyPositivePattern = RHS
End Property

''
' Returns the separator character for groups of numbers.
'
' @return The character that is used as a separator between groups of numbers.
' This is typically a comma.
'
Public Property Get NumberGroupSeparator() As String
    VerifyLoaded
    NumberGroupSeparator = mProps.NumberGroupSeparator
End Property

''
' Sets the separator charactor for a group of numbers.
'
' @param RHS The character to be used when separating groups of numbers.
' This is typically a comma.
'
Public Property Let NumberGroupSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NumberGroupSeparator = RHS
End Property

''
' Returns the pattern to be used when formatting negative numbers.
'
' @return The pattern to be used.
'
Public Property Get NumberNegativePattern() As NumberNegativePatterns
    VerifyLoaded
    NumberNegativePattern = mProps.NumberNegativePattern
End Property

''
' Sets the pattern to be used when formatting negative numbers.
'
' @param RHS The new pattern to be used.
'
Public Property Let NumberNegativePattern(ByVal RHS As NumberNegativePatterns)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 4 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Index))
    mProps.NumberNegativePattern = RHS
End Property

''
' Returns the current group sizes array for formatted numbers.
'
' @return An array containing grouping sizes for a formatted number.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Get NumberGroupSizes() As Variant
    VerifyLoaded
    NumberGroupSizes = mProps.NumberGroupSizes
End Property

''
' Sets the group sizes for formatted numbers.
'
' @param RHS A vbLong array containing the grouping sizes.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Let NumberGroupSizes(ByRef RHS As Variant)
    VerifyWritable
    VerifyLoaded
    If CorArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException("NumberGroupSizes", Environment.GetResourceString(ArgumentNull_Array))
    
    VerifyGroupSizes RHS
    mProps.NumberGroupSizes = RHS
End Property

''
' The default number of digits to be appended after the decimal point if one is used.
'
' @return The number of digits to be appended after the decimal point if one is used.
'
Public Property Get NumberDecimalDigits() As Long
    VerifyLoaded
    NumberDecimalDigits = mProps.NumberDecimalDigits
End Property

''
' Sets the default number of digits to be appended after the decimal point.
'
' @param RHS The number of digits to be appended after the decimal point.
'
Public Property Let NumberDecimalDigits(ByVal RHS As Long)
    VerifyWritable
    VerifyLoaded
    If RHS < 0 Or RHS > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99))
    mProps.NumberDecimalDigits = RHS
End Property

''
' Returns the character used as a decimal point.
'
' @return The decimal point character.
'
Public Property Get NumberDecimalSeparator() As String
    VerifyLoaded
    NumberDecimalSeparator = mProps.NumberDecimalSeparator
End Property

''
' Sets the character to be used as the decimal point.
'
' @param RHS The new character to be used as the decimal point.
'
Public Property Let NumberDecimalSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NumberDecimalSeparator = RHS
End Property

''
' Returns the negative symbol for negative numbers.
'
' @return The current negative symbol.
'
Public Property Get NegativeSign() As String
    VerifyLoaded
    NegativeSign = mProps.NegativeSign
End Property

''
' Sets the negative symbol to be used for negative numbers.
'
' @param RHS The new negative symbol.
'
Public Property Let NegativeSign(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.NegativeSign = RHS
End Property

''
' Returns the current positive symbol.
'
' @return The current positive symbol.
'
Public Property Get PositiveSign() As String
    VerifyLoaded
    PositiveSign = mProps.PositiveSign
End Property

''
' Sets the positive symbol.
'
' @param RHS The new positive symbol.
'
Public Property Let PositiveSign(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PositiveSign = RHS
End Property

''
' Returns a format provider of the requested type.
'
' @param FormatType The name of the format provider.
' @return The requested format provider, or Nothing if no format provider was found.
'
Public Function GetFormat(ByRef FormatType As String) As Object
    If CorString.Equals(FormatType, vbNumberFormatInfo, OrdinalIgnoreCase) Then
        Set GetFormat = Me
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

Friend Function Format(ByRef Value As Variant, Optional ByRef FormatText As String) As String
    VerifyLoaded
    
    Dim vt As Long
    vt = VarType(Value)
    If vt = vbVariant Then
        vt = MemLong(MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)) And &HFF
    End If
    
    Select Case vt
        Case vbLong, vbInteger, vbByte
            Format = FormatInt32(Value, FormatText)
        Case vbDouble
            Format = FormatDouble(Value, FormatText)
        Case vbSingle
            Format = FormatSingle(Value, FormatText)
        Case Else
            Format = Value
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitPatterns
    
    With mNumber.DigitsSA
        .cbElements = 1
        .cDims = 1
        .cElements = 128
    End With
    
    mPtrNumberBuffer = CoTaskMemAlloc(40)
    If mPtrNumberBuffer = vbNullPtr Then _
        Throw New OutOfMemoryException
    
    SAPtr(mNumber.Digits) = VarPtr(mNumber.DigitsSA)
    
    InitChars mAppendStringBuffer
    InitChars mFormatSpecifier
End Sub

Private Sub Class_InitProperties()
    mProps.PerMilleSymbol = ChrW$(&H2030)
End Sub

Friend Sub InitCultureInfo(ByVal Info As CultureInfo)
    Load Info.LCID, Info.UseUserOverride
    If Info.IsReadOnly Then
        MarkReadOnly
    End If
End Sub

Friend Sub InitClone(ByRef Props As PropsType)
    mProps = Props
    mLoaded = True
    mProps.IsReadOnly = False
End Sub

Private Sub Class_Terminate()
    SAPtr(mNumber.Digits) = 0
    CoTaskMemFree mPtrNumberBuffer
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load(ByVal LCID As Long, ByVal UseUserOverride As Boolean)
    If Not CultureTable.IsLoaded Then _
        Throw Cor.NewNotSupportedException("Cannot load NumberFormatInfo without culture table.")
    
    LoadCommon LCID
    
    If UseUserOverride Then
        LoadUserOverride LCID
    Else
        LoadFromCultureTable LCID
    End If
    mLoaded = True
End Sub

Private Sub VerifyLoaded()
    If mLoaded Then Exit Sub
    Load INVARIANT_LCID, False
End Sub

Private Sub LoadUserOverride(ByVal LCID As Long)
    With mProps
        .NumberDecimalSeparator = GetLocaleString(LCID, LOCALE_SDECIMAL)
        .NumberDecimalDigits = GetLocaleLong(LCID, LOCALE_IDIGITS)
        .NumberGroupSeparator = GetLocaleString(LCID, LOCALE_STHOUSAND)
        .NumberGroupSizes = ParseGroupSizes(GetLocaleString(LCID, LOCALE_SGROUPING))
        .NumberNegativePattern = GetLocaleLong(LCID, LOCALE_INEGNUMBER)
        .PercentDecimalSeparator = .NumberDecimalSeparator
        .PercentDecimalDigits = .NumberDecimalDigits
        .PercentGroupSeparator = .NumberGroupSeparator
        .PercentGroupSizes = .NumberGroupSizes
        .CurrencyDecimalSeparator = GetLocaleString(LCID, LOCALE_SMONDECIMALSEP)
        .CurrencyDecimalDigits = GetLocaleLong(LCID, LOCALE_ICURRDIGITS)
        .CurrencyGroupSeparator = GetLocaleString(LCID, LOCALE_SMONTHOUSANDSEP)
        .CurrencyGroupSizes = ParseGroupSizes(GetLocaleString(LCID, LOCALE_SMONGROUPING))
        .CurrencyNegativePattern = GetLocaleLong(LCID, LOCALE_INEGCURR)
        .CurrencyPositivePattern = GetLocaleLong(LCID, LOCALE_ICURRENCY)
        .CurrencySymbol = GetLocaleString(LCID, LOCALE_SCURRENCY)
        .NegativeSign = GetLocaleString(LCID, LOCALE_SNEGATIVESIGN)
    End With
End Sub

Private Sub LoadFromCultureTable(ByVal LCID As Long)
    With mProps
        .NumberDecimalDigits = CultureTable.GetNumber(LCID, INUMBERDECIMALDIGITS)
        .NumberDecimalSeparator = CultureTable.GetString(LCID, SNUMBERDECIMALSEPARATOR)
        .NumberGroupSeparator = CultureTable.GetString(LCID, SNUMBERGROUPSEPARATOR)
        .NumberGroupSizes = CultureTable.GetNumberArray(LCID, SNUMBERGROUPSIZES)
        .NumberNegativePattern = CultureTable.GetNumber(LCID, INUMBERNEGATIVEPATTERN)
        .PercentDecimalDigits = CultureTable.GetNumber(LCID, IPERCENTDECIMALDIGITS)
        .PercentDecimalSeparator = CultureTable.GetString(LCID, SPERCENTDECIMALSEPARATOR)
        .PercentGroupSeparator = CultureTable.GetString(LCID, SPERCENTGROUPSEPARATOR)
        .PercentGroupSizes = CultureTable.GetNumberArray(LCID, SPERCENTGROUPSIZES)
        .PercentNegativePattern = CultureTable.GetNumber(LCID, IPERCENTNEGATIVEPATTERN)
        .CurrencyDecimalDigits = CultureTable.GetNumber(LCID, ICURRENCYDECIMALDIGITS)
        .CurrencyDecimalSeparator = CultureTable.GetString(LCID, SCURRENCYDECIMALSEPARATOR)
        .CurrencyGroupSeparator = CultureTable.GetString(LCID, SCURRENCYGROUPSEPARATOR)
        .CurrencyGroupSizes = CultureTable.GetNumberArray(LCID, SCURRENCYGROUPSIZES)
        .CurrencyNegativePattern = CultureTable.GetNumber(LCID, ICURRENCYNEGATIVEPATTERN)
        .CurrencyPositivePattern = CultureTable.GetNumber(LCID, ICURRENCYPOSITIVEPATTERN)
        .CurrencySymbol = CultureTable.GetString(LCID, SCURRENCYSYMBOL)
        .NegativeSign = CultureTable.GetString(LCID, SNEGATIVESIGN)
    End With
End Sub

Private Sub LoadCommon(ByVal LCID As Long)
    With mProps
        .NaNSymbol = CultureTable.GetString(LCID, SNANSYMBOL)
        .PositiveInfinitySymbol = CultureTable.GetString(LCID, SPOSITIVEINFINITYSYMBOL)
        .NegativeInfinitySymbol = CultureTable.GetString(LCID, SNEGATIVEINFINITYSYMBOL)
        .PercentPositivePattern = CultureTable.GetNumber(LCID, IPERCENTPOSITIVEPATTERN)
        .PercentNegativePattern = CultureTable.GetNumber(LCID, IPERCENTNEGATIVEPATTERN)
        .PercentSymbol = CultureTable.GetString(LCID, SPERCENTSYMBOL)
        .PositiveSign = CultureTable.GetString(LCID, SPOSITIVESIGN)
    End With
End Sub

Private Function ParseGroupSizes(ByRef Source As String) As Long()
    Dim SourceSizes() As String
    SourceSizes = Split(Source, ";")
    
    Dim Result() As Long
    ReDim Result(0 To UBound(SourceSizes) - 1) ' we don't use the last element.
    
    Dim i As Long
    For i = 0 To UBound(Result)
        Result(i) = CLng(SourceSizes(i))
    Next i
    
    ParseGroupSizes = Result
End Function

Private Function FormatSingle(ByVal Value As Single, ByRef Format As String) As String
    Dim FormatPrecision As Long
    Dim FormatType      As Long
    
    GetFormatSpecifierParts Format, FormatType, FormatPrecision

    If ((FormatType And &HFFDF) <> GeneralFormat) And (FormatType <> CustomFormat) Then
        DoubleToNumber Value, DefaultSinglePrecision
        If Not IsValidNumber(FormatSingle) Then Exit Function
        If Value = 0! Then mNumber.Precision = 0
    End If

    Select Case FormatType And &HFFDF
        Case FixedFormat:               FormatFixedDouble FormatPrecision
        Case ExponentFormat:            FormatScientific FormatPrecision, FormatType
        Case DecimalFormat, HexFormat:  Throw Cor.NewFormatException("Invalid format specifier.")
        Case NumberFormat:              FormatNumber FormatPrecision
        Case CurrencyFormat:            FormatCurrency FormatPrecision
        
        Case GeneralFormat
            If FormatPrecision < 0 Then
                FormatPrecision = DefaultSinglePrecision
            End If
            
            DoubleToNumber Value, FormatPrecision
            
            If Not IsValidNumber(FormatSingle) Then _
                Exit Function
            
            If mNumber.Digits(0) = vbZeroChar Then
                FormatSingle = "0"
                Exit Function
            End If
            
            If FormatPrecision > 7 Then
                FormatPrecision = 9
            End If
            
            RoundNumber FormatPrecision
            FormatGeneralDouble mNumber.Precision, FormatType - 2
        
        Case PercentFormat
            If FormatPrecision = UnknownPrecision Then
                FormatPrecision = mProps.PercentDecimalDigits
            End If
            
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            RoundNumber FormatPrecision + mNumber.DecimalPosition
            FormatPercent FormatPrecision

        Case Else
            If Value < 0 Then
                FormatSingle = mProps.NegativeSign & VBA.Format$(Abs(Value), Format)
            Else
                FormatSingle = VBA.Format$(Value, Format)
            End If
            Exit Function
    End Select
    If mOutputPos = 0 Then
        FormatSingle = "0"
    Else
        FormatSingle = SysAllocStringLen(VarPtr(mOutput(0)), mOutputPos)
    End If
End Function

Private Function FormatDouble(ByVal Value As Double, ByRef Format As String) As String
    Dim FormatPrecision As Long
    Dim FormatType      As Long
    
    GetFormatSpecifierParts Format, FormatType, FormatPrecision
    
    If ((FormatType And &HFFDF) <> GeneralFormat) And (FormatType <> CustomFormat) Then
        DoubleToNumber Value, DefaultDoublePrecision
        
        If Not IsValidNumber(FormatDouble) Then
            Exit Function
        End If
        
        If Value = 0# Then
            mNumber.Precision = 0
        End If
    End If
    
    Select Case FormatType And &HFFDF
        Case FixedFormat:               FormatFixedDouble FormatPrecision
        Case ExponentFormat:            FormatScientific FormatPrecision, FormatType
        Case DecimalFormat, HexFormat:  Throw Cor.NewFormatException("Invalid format specifier.")
        Case NumberFormat:              FormatNumber FormatPrecision
        Case CurrencyFormat:            FormatCurrency FormatPrecision
        
        Case GeneralFormat
            If FormatPrecision = UnknownPrecision Or FormatPrecision = 0 Then
                FormatPrecision = DefaultDoublePrecision
            End If
            
            DoubleToNumber Value, FormatPrecision
            
            If Not IsValidNumber(FormatDouble) Then Exit Function
            
            If mNumber.Digits(0) = vbZeroChar Then
                FormatDouble = "0"
                Exit Function
            End If
            
            RoundNumber FormatPrecision
            FormatGeneralDouble FormatPrecision, FormatType - 2
        
        Case PercentFormat
            If FormatPrecision = UnknownPrecision Then
                FormatPrecision = mProps.PercentDecimalDigits
            End If
            
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            RoundNumber FormatPrecision + mNumber.DecimalPosition
            FormatPercent FormatPrecision
        
        Case Else
            If Value < 0# Then
                FormatDouble = mProps.NegativeSign & VBA.Format$(Abs(Value), Format)
            Else
                FormatDouble = VBA.Format$(Value, Format)
            End If
            Exit Function
    End Select
    If mOutputPos = 0 Then
        FormatDouble = "0"
    Else
        FormatDouble = SysAllocStringLen(VarPtr(mOutput(0)), mOutputPos)
    End If
End Function

Private Function FormatInt32(ByRef Value As Variant, ByRef Format As String) As String
    Dim Specifiers As Specifiers
    
    Specifiers = ParseSpecifiers(Format)
    
    If Specifiers.Format <> CustomSpecifier And _
        Specifiers.Format <> HexUpperSpecifier And _
        Specifiers.Format <> HexLowerSpecifier Then
        Int32ToNumber Value
    End If
    
    Select Case Specifiers.Format
        Case FormatSpecifier.GeneralLowerSpecifier
            FormatGeneralInt32 Specifiers.Precision, vbLowerEChar
            
        Case FormatSpecifier.GeneralUpperSpecifier
            FormatGeneralInt32 Specifiers.Precision, vbUpperEChar
            
        Case FormatSpecifier.FixedUpperSpecifier, FormatSpecifier.FixedLowerSpecifier
            FormatFixedDouble Specifiers.Precision
            
        Case FormatSpecifier.DecimalUpperSpecifier, FormatSpecifier.DecimalLowerSpecifier
            FormatDecimal Specifiers.Precision
            
        Case FormatSpecifier.NumberUpperSpecifier, FormatSpecifier.NumberLowerSpecifier
            FormatNumber Specifiers.Precision
            
        Case FormatSpecifier.CurrencyUpperSpecifier, FormatSpecifier.CurrencyLowerSpecifier
            FormatCurrency Specifiers.Precision
            
        Case FormatSpecifier.ExponentLowerSpecifier
            FormatExponentInt32 Specifiers.Precision, vbLowerEChar
            
        Case FormatSpecifier.ExponentUpperSpecifier
            FormatExponentInt32 Specifiers.Precision, vbUpperEChar
        
        Case FormatSpecifier.PercentUpperSpecifier, FormatSpecifier.PercentLowerSpecifier
            If mNumber.Precision > 0 Then
                mNumber.DecimalPosition = mNumber.DecimalPosition + 2
            End If
            
            If Specifiers.Precision = UnknownPrecision Then
                FormatPercent mProps.PercentDecimalDigits
            Else
                FormatPercent Specifiers.Precision
            End If

        Case FormatSpecifier.HexLowerSpecifier, FormatSpecifier.HexUpperSpecifier
            FormatHex Value, Specifiers
            FormatInt32 = SysAllocStringLen(VarPtr(mOutput(mOutputPos)), BUFFER_SIZE - mOutputPos)
            Exit Function
        
        Case Else
            If Value < 0 Then
                FormatInt32 = mProps.NegativeSign & VBA.Format$(Abs(Value), Format)
            Else
                FormatInt32 = VBA.Format$(Value, Format)
            End If
            Exit Function
    End Select
    
    If mOutputPos = 0 Then
        FormatInt32 = "0"
    Else
        FormatInt32 = SysAllocStringLen(VarPtr(mOutput(0)), mOutputPos)
    End If
End Function

Private Sub FormatExponentInt32(ByVal FormatPrecision As Long, ByVal FormatType As Long)
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = 6
    End If
    
    Dim Places As Long
    Places = mNumber.DecimalPosition - 1
    mNumber.DecimalPosition = 1
    
    FormatFixedDouble FormatPrecision
    AppendExponent Places, FormatType, 3
End Sub

Private Sub FormatNumber(ByVal FormatPrecision As Long)
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = mProps.NumberDecimalDigits
    End If
    
    RoundNumber mNumber.DecimalPosition + FormatPrecision
    
    Dim Pattern As String
    If mNumber.IsNegative Then
        Pattern = mNumberNegativePatterns(mProps.NumberNegativePattern)
    Else
        Pattern = "#"
    End If
    
    InitChars mFormatSpecifier, Pattern
    
    Dim i   As Long
    Dim Ch  As Integer
    With mProps
        For i = 0 To Len(Pattern) - 1
            Ch = mFormatSpecifier.Chars(i)
            Select Case Ch
                Case vbPoundChar:    FormatGroups FormatPrecision, .NumberGroupSizes, .NumberGroupSeparator, .NumberDecimalSeparator
                Case vbMinusChar:    AppendString .NegativeSign
                Case Else
                    mOutput(mOutputPos) = Ch
                    mOutputPos = mOutputPos + 1
            End Select
        Next i
    End With
End Sub

Private Sub FormatCurrency(ByVal DecimalPlaces As Long)
    If DecimalPlaces = UnknownPrecision Then
        DecimalPlaces = mProps.CurrencyDecimalDigits
    End If
    
    RoundNumber mNumber.DecimalPosition + DecimalPlaces
    
    Dim Pattern As String
    If mNumber.IsNegative Then
        Pattern = mCurrencyNegativePatterns(mProps.CurrencyNegativePattern)
    Else
        Pattern = mCurrencyPositivePatterns(mProps.CurrencyPositivePattern)
    End If
    
    InitChars mFormatSpecifier, Pattern
    
    Dim i   As Long
    Dim Ch  As Integer
    With mProps
        For i = 0 To Len(Pattern) - 1
            Ch = mFormatSpecifier.Chars(i)
            Select Case Ch
                Case vbPoundChar:    FormatGroups DecimalPlaces, .CurrencyGroupSizes, .CurrencyGroupSeparator, .CurrencyDecimalSeparator
                Case vbMinusChar:    AppendString .NegativeSign
                Case vbDollarChar:   AppendString .CurrencySymbol
                Case Else
                    mOutput(mOutputPos) = Ch
                    mOutputPos = mOutputPos + 1
            End Select
        Next i
    End With
End Sub

Private Sub FormatPercent(ByVal DecimalPlaces As Long)
    Dim Pattern As String
    If mNumber.IsNegative Then
        Pattern = mPercentNegativePatterns(mProps.PercentNegativePattern)
    Else
        Pattern = mPercentPositivePatterns(mProps.PercentPositivePattern)
    End If
    
    InitChars mFormatSpecifier, Pattern
    
    Dim i   As Long
    Dim Ch  As Integer
    With mProps
        For i = 0 To Len(Pattern) - 1
            Ch = mFormatSpecifier.Chars(i)
            Select Case Ch
                Case vbPoundChar:    FormatGroups DecimalPlaces, .PercentGroupSizes, .PercentGroupSeparator, .PercentDecimalSeparator
                Case vbMinusChar:    AppendString .NegativeSign
                Case vbPercentChar:  AppendString .PercentSymbol
                Case Else
                    mOutput(mOutputPos) = Ch
                    mOutputPos = mOutputPos + 1
            End Select
        Next i
    End With
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
Private Sub FormatGroups(ByVal DecimalPlaces As Long, ByRef GroupSizes() As Long, ByRef GroupSeparator As String, ByRef DecimalSeparator As String)
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
            Next i

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
                If DigitCount = 0 Then DigitCount = GroupSize
            Else
                DigitCount = RepeatingGroupDigitCount
            End If

            DigitIndex = 0
            Ch = .Digits(DigitIndex)
            Do While RepeatingGroupDigitCount > 0
                If DigitCount = 0 And HasGroups Then
                    AppendString GroupSeparator
                    DigitCount = GroupSize
                End If
                If Ch = vbTerminator Then
                    mOutput(mOutputPos) = vbZeroChar
                Else
                    mOutput(mOutputPos) = Ch
                    Ch = .Digits(DigitIndex + 1)
                End If
                DigitIndex = DigitIndex + 1
                mOutputPos = mOutputPos + 1
                DigitCount = DigitCount - 1
                RepeatingGroupDigitCount = RepeatingGroupDigitCount - 1
            Loop
        Else
            mOutput(mOutputPos) = vbZeroChar
            mOutputPos = mOutputPos + 1
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
                        AppendString GroupSeparator
                    ElseIf DigitCount = GroupSize Then
                        AppendString GroupSeparator
                        GroupSizeIndex = GroupSizeIndex - 1
                        If GroupSizeIndex < 0 Then
                            GroupSizeIndex = 0
                        End If
                        GroupSize = GroupSize - GroupSizes(GroupSizeIndex)
                    End If
                End If
                If Ch = 0 Then
                    mOutput(mOutputPos) = vbZeroChar
                Else
                    mOutput(mOutputPos) = Ch
                    Ch = .Digits(DigitIndex + 1)
                End If
                DigitIndex = DigitIndex + 1
                mOutputPos = mOutputPos + 1
                DigitCount = DigitCount - 1
            Next
        End If
        
        ' Append any digits that need to appear after a decimal point
        If DecimalPlaces > 0 Then
            AppendString DecimalSeparator
            DigitIndex = .DecimalPosition
            
            ' negative scale represents numbers such as 0.000000123,
            ' so we need to insert the zeros before the digits.
            Do While DigitIndex < 0
                mOutput(mOutputPos) = vbZeroChar
                mOutputPos = mOutputPos + 1
                DigitIndex = DigitIndex + 1
                DecimalPlaces = DecimalPlaces - 1
            Loop
            If Ch <> vbTerminator Then Ch = .Digits(i)
            Do While DecimalPlaces > 0
                If Ch = 0 Then
                    mOutput(mOutputPos) = vbZeroChar
                Else
                    mOutput(mOutputPos) = Ch
                    DigitIndex = DigitIndex + 1
                    Ch = .Digits(DigitIndex)
                End If
                mOutputPos = mOutputPos + 1
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
Private Sub FormatHex(ByRef Value As Variant, ByRef Specifiers As Specifiers)
    Dim Digits      As Long
    Dim CaseOffset  As Long
    Dim Ch          As Integer
    Dim HexValue    As Long
    
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
    
    CaseOffset = Specifiers.Format - 33
    
    mOutputPos = BUFFER_SIZE - 1
    Do While HexValue <> 0
        Ch = HexValue And &HF
        If Ch > 9 Then
            mOutput(mOutputPos) = Ch + CaseOffset  ' add the offset to upper or lower 'a-f' or 'A-F'.
        Else
            mOutput(mOutputPos) = vbZeroChar + Ch
        End If
        mOutputPos = mOutputPos - 1
        If HexValue < 0 Then
            HexValue = Helper.ShiftRight(HexValue, 4)
        Else
            HexValue = HexValue \ &H10
        End If
    Loop
    
    ' Add leading zeros if needed.
    Do While mOutputPos >= BUFFER_SIZE - Digits
        mOutput(mOutputPos) = vbZeroChar
        mOutputPos = mOutputPos - 1
    Loop
    mOutputPos = mOutputPos + 1
End Sub

''
' Formats an integer value with enough leading zeros as necessary
' to be as long as the number of digits specified.
'
' @param digits The number of digits to be displayed, filling the
' number with leading zeros if necessary.
'
Private Sub FormatDecimal(ByVal TotalDigits As Long)
    With mNumber
        ' Add the negative sign if necessary.
        If .IsNegative Then
            AppendString mProps.NegativeSign
        End If
        
        ' Insert the leading zeros until we reach the first digit.
        Do While TotalDigits > .Precision
            mOutput(mOutputPos) = vbZeroChar
            mOutputPos = mOutputPos + 1
            TotalDigits = TotalDigits - 1
        Loop
        
        ' Fill the rest of the buffer with the digits.
        Dim i As Long
        For i = 0 To .Precision - 1
            mOutput(mOutputPos) = .Digits(i)
            mOutputPos = mOutputPos + 1
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
Private Sub FormatScientific(ByVal Digits As Long, ByVal ExpChar As Integer)
    If Digits = UnknownPrecision Then
        Digits = DefaultExponentPrecision
    End If
    
    RoundNumber Digits + 1
    
    With mNumber
        If .IsNegative Then
            AppendString mProps.NegativeSign
        End If
        
        ' We need to represent "digits" number of digits in the precision.
        ' After the first digit is ouptput we put the decimal separator, since
        ' all scientific notation has a leading digit-period-digits format.
        Dim i           As Long
        Dim Ch          As Integer
        Dim OutOfDigits As Boolean
        
        Do While Digits >= 0
            If i = 1 Then
                AppendString mProps.NumberDecimalSeparator
            End If
            
            If Not OutOfDigits Then
                Ch = .Digits(i)
            End If
            
            ' Simply copy the digit to the output buffer or place
            ' a 0 in the output buffer if we have run out of digits
            ' but required more digits to be displayed.
            If Ch = 0 Then
                mOutput(mOutputPos) = vbZeroChar
                OutOfDigits = True
            Else
                mOutput(mOutputPos) = Ch
            End If
            
            i = i + 1
            mOutputPos = mOutputPos + 1
            Digits = Digits - 1
        Loop
        AppendExponent .DecimalPosition - 1, ExpChar, 3
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
    If Digits > mNumber.Precision Then Exit Sub
    
    With mNumber
        ' are we rounding up or down?
        If .Digits(Digits) >= vbFiveChar Then
            ' We don't round 9's directly, since they would just
            ' carry over, so we keep skipping them until there
            ' are no more. Then we'll round up the first non-nine digit.
            ' This will also lower the precision, since we had
            ' to keep dealing with the carry.
            Do While Digits > 0
                If .Digits(Digits - 1) <> vbNineChar Then Exit Do
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
                If .Digits(Digits - 1) <> vbZeroChar Then Exit Do
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

''
' This will parse the format type and optional numerical value.
'
' @param Format The string format to be parsed (ie. "D99").
' @param FormatType The type of formatting parsed from the string.
' @param Precision The number of digits parsed out.
' @remarks A format of "D99" will return DecimalFormat in FormatType and
' 99 in Precision. We indicate no digits were supplied by using -1.
'
Private Sub GetFormatSpecifierParts(ByRef Format As String, ByRef FormatType As FormatType, ByRef FormatPrecision As Long)
    Dim FormatLength As Long
    FormatLength = Len(Format)
    
    ' A valid format specifier is 1 to 3 characters long.
    Select Case FormatLength
        Case 0
            FormatType = GeneralFormat
            FormatPrecision = UnknownPrecision
            
        Case Is > 3
            FormatType = CustomFormat
            FormatPrecision = UnknownPrecision
            
        Case Else
            InitChars mFormatSpecifier, Format
            FormatType = mFormatSpecifier.Chars(0)
            
            Select Case FormatType
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    ' We've got a format letter, so attempt to get any precision digits.
                    If FormatLength = 1 Then
                        FormatPrecision = UnknownPrecision
                        Exit Sub
                    End If
                    
                    FormatPrecision = 0   ' we need to reset this because it's a ByRef variable.
                    
                    Dim Ch  As Long
                    Dim i   As Long
                    For i = 1 To FormatLength - 1
                        Ch = mFormatSpecifier.Chars(i)
                        
                        Select Case Ch
                            Case vbZeroChar To vbNineChar
                                FormatPrecision = FormatPrecision * 10 + Ch - vbZeroChar
                            Case Else
                                FormatType = CustomFormat
                                FormatPrecision = UnknownPrecision
                                Exit Sub
                        End Select
                    Next i
                    
                Case Else
                    FormatType = CustomFormat
                    FormatPrecision = UnknownPrecision
                    Exit Sub
            End Select
    End Select
End Sub

Private Function ParseSpecifiers(ByRef Format As String) As Specifiers
    Dim FormatLength As Long
    FormatLength = Len(Format)
    
    Select Case FormatLength
        Case 0
            ParseSpecifiers.Format = GeneralUpperSpecifier
            ParseSpecifiers.Precision = UnknownPrecision
            
        Case Is > 3
            ParseSpecifiers.Format = CustomSpecifier
            ParseSpecifiers.Precision = UnknownPrecision
            
        Case Else
            InitChars mFormatSpecifier, Format
            
            Select Case mFormatSpecifier.Chars(0)
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    ParseSpecifiers.Format = mFormatSpecifier.Chars(0)
                    
                    If FormatLength = 1 Then
                        ParseSpecifiers.Precision = UnknownPrecision
                        Exit Function
                    End If
                    
                    Dim Precision   As Long
                    Dim Ch          As Long
                    Dim i           As Long
                    For i = 1 To FormatLength - 1
                        Ch = mFormatSpecifier.Chars(i)
                        
                        Select Case Ch
                            Case vbZeroChar To vbNineChar
                                Precision = Precision * 10 + Ch - vbZeroChar
                            Case Else
                                ParseSpecifiers.Format = CustomSpecifier
                                ParseSpecifiers.Precision = UnknownPrecision
                                Exit Function
                        End Select
                    Next i
                    
                    ParseSpecifiers.Precision = Precision
                Case Else
                    ParseSpecifiers.Format = CustomSpecifier
                    ParseSpecifiers.Precision = UnknownPrecision
            End Select
    End Select
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
    Dim IsMaxNeg As Boolean
    
    mOutputPos = 0
    With mNumber
        .DigitsSA.pvData = mPtrNumberBuffer
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
        Dim DigitIndex As Long
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
        Dim OneTenth As Long
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

Private Sub DoubleToNumber(ByVal Value As Double, ByVal FormatPrecision As Long)
    mOutputPos = 0
    With AsDLong(Value)
        If (.HiDWord And &H7FF00000) = &H7FF00000 Then
            If (.HiDWord And &HFFFFF) Or (.LoDWord) Then
                mNumber.DecimalPosition = SCALE_NAN
            Else
                mNumber.DecimalPosition = SCALE_INF
                mNumber.IsNegative = (Value < 0#)
            End If
        Else
            mNumber.DigitsSA.pvData = Helper.Ecvt(Value, FormatPrecision, mNumber.DecimalPosition, mNumber.IsNegative)
            mNumber.Precision = FormatPrecision
        End If
    End With
End Sub

Private Sub FormatGeneralInt32(ByVal FormatPrecision As Long, ByVal ExpChar As Long)
    If FormatPrecision = UnknownPrecision Or FormatPrecision = 0 Then
        FormatPrecision = DefaultIntegerPrecision
    End If
    
    With mNumber
        If .IsNegative Then
            AppendString mProps.NegativeSign
        End If
        
        ' If the requested precision is less than the
        ' actuall number of digits, then we have to
        ' go the scientific format route.
        Dim HasExponent As Boolean
        If FormatPrecision < .Precision Then
            .DecimalPosition = 1
            RoundNumber FormatPrecision
            HasExponent = True
        Else
            FormatPrecision = .Precision
        End If
        
        Dim i As Long
        For i = 0 To FormatPrecision - 1
            If i = .DecimalPosition Then
                AppendString mProps.NumberDecimalSeparator
            End If
            
            mOutput(mOutputPos) = .Digits(i)
            mOutputPos = mOutputPos + 1
        Next i
        
        If HasExponent Then
            AppendExponent .Precision - 1, ExpChar, 2
        End If
    End With
End Sub

Private Sub FormatGeneralDouble(ByVal FormatPrecision As Long, ByVal ExpChar As Integer)
    With mNumber
        If .IsNegative Then
            AppendString mProps.NegativeSign
        End If
        
        Dim DP As Long
        DP = .DecimalPosition
        
        Dim HasExponent As Boolean
        If FormatPrecision < .DecimalPosition Or .DecimalPosition < -3 Then
            HasExponent = True
            DP = 1
        ElseIf DP <= 0 Then
            mOutput(mOutputPos) = vbZeroChar
            mOutputPos = mOutputPos + 1
            DP = 0
        End If
        
        Dim Ch As Integer
        Dim i As Long
        Ch = .Digits(i)
        
        Do While Ch <> 0
            If i = DP Then
                AppendString mProps.NumberDecimalSeparator
                
                If Not HasExponent Then
                    For DP = DP To .DecimalPosition + 1 Step -1
                        mOutput(mOutputPos) = vbZeroChar
                        mOutputPos = mOutputPos + 1
                    Next DP
                End If
            End If
            
            mOutput(mOutputPos) = Ch
            mOutputPos = mOutputPos + 1
            i = i + 1
            Ch = .Digits(i)
        Loop
        If HasExponent Then
            AppendExponent .DecimalPosition - 1, ExpChar, 2
        Else
            For i = i To .DecimalPosition - 1
                mOutput(mOutputPos) = vbZeroChar
                mOutputPos = mOutputPos + 1
            Next i
        End If
    End With
End Sub

Private Sub FormatFixedInt32(ByVal FormatPrecision As Long)
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = mProps.NumberDecimalDigits
    End If
    
    With mNumber
        If .IsNegative Then
            AppendString mProps.NegativeSign
        End If
        
        Dim i As Long
        If .Precision = 0 Then
            mOutput(mOutputPos) = vbZeroChar
            mOutputPos = mOutputPos + 1
        Else
            For i = 0 To .Precision - 1
                mOutput(mOutputPos) = .Digits(i)
                mOutputPos = mOutputPos + 1
            Next i
        End If
        
        If FormatPrecision > 0 Then
            AppendString mProps.NumberDecimalSeparator
            
            For i = 1 To FormatPrecision
                mOutput(mOutputPos) = vbZeroChar
                mOutputPos = mOutputPos + 1
            Next i
        End If
    End With
End Sub

Private Sub FormatFixedDouble(ByVal FormatPrecision As Long)
    If FormatPrecision = UnknownPrecision Then
        FormatPrecision = mProps.NumberDecimalDigits
    End If
    
    RoundNumber FormatPrecision + mNumber.DecimalPosition
    
    Dim i   As Long
    Dim Ch  As Integer
    With mNumber
        If .IsNegative Then
            AppendString mProps.NegativeSign
        End If
        
        If .Precision > 0 And .DecimalPosition > 0 Then
            For i = 0 To .DecimalPosition - 1
                mOutput(mOutputPos) = .Digits(i)
                mOutputPos = mOutputPos + 1
            Next i
        Else
            mOutput(mOutputPos) = vbZeroChar
            mOutputPos = mOutputPos + 1
            i = .DecimalPosition
        End If
        
        If FormatPrecision > 0 Then
            AppendString mProps.NumberDecimalSeparator
            
            Do While FormatPrecision > 0
                If i >= 0 Then
                    Ch = .Digits(i)
                End If
                
                If .Precision = 0 Or Ch = 0 Then
                    mOutput(mOutputPos) = vbZeroChar
                    If i < 0 Then
                        i = i + 1
                    End If
                Else
                    mOutput(mOutputPos) = Ch
                    i = i + 1
                End If
                mOutputPos = mOutputPos + 1
                FormatPrecision = FormatPrecision - 1
            Loop
        End If
    End With
End Sub

''
' Appends a string to the buffer, using the most effecient method.
'
Private Sub AppendString(ByRef StringToAppend As String)
    Dim StringToAppendLength As Long
    StringToAppendLength = Len(StringToAppend)
    
    If StringToAppendLength < 6 Then
        InitChars mAppendStringBuffer, StringToAppend
        Dim i As Long
        For i = 0 To StringToAppendLength - 1
            mOutput(mOutputPos + i) = mAppendStringBuffer.Chars(i)
        Next i
    Else
        CopyMemory mOutput(mOutputPos), ByVal StrPtr(StringToAppend), StringToAppendLength * vbSizeOfChar
    End If
    
    mOutputPos = mOutputPos + StringToAppendLength
End Sub

Private Function IsValidNumber(ByRef RetVal As String) As Boolean
    Select Case mNumber.DecimalPosition
        Case SCALE_NAN
            RetVal = mProps.NaNSymbol
        Case SCALE_INF
            If mNumber.IsNegative Then
                RetVal = mProps.NegativeInfinitySymbol
            Else
                RetVal = mProps.PositiveInfinitySymbol
            End If
        Case Else
            IsValidNumber = True
    End Select
End Function

''
' This adds the 'e+000' portion of a scientific formatted output.
'
Private Sub AppendExponent(ByVal Places As Long, ByVal ExpChar As Integer, ByVal MinDigits As Long)
    mOutput(mOutputPos) = ExpChar
    mOutputPos = mOutputPos + 1
    
    If mNumber.Precision = 0 Then Places = 0
    
    If Places < 0 Then
        Places = -Places
        AppendString mProps.NegativeSign
    Else
        AppendString mProps.PositiveSign
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
    Dim i As Long
    For i = Digits To MinDigits - 1
        mOutput(mOutputPos) = vbZeroChar
        mOutputPos = mOutputPos + 1
    Next i
    
    ' Move our buffer pointer out to the end.
    mOutputPos = mOutputPos + Digits
    
    ' We add the right-most character to the buffer. We add
    ' them in reverse order, starting from the rightside and
    ' moving back towards the left. This is the easiest way
    ' mathematically without having to reverse the characters.
    Dim pos As Long
    pos = mOutputPos - 1
    Do While Places > 0
        Dim OneTenth As Long
        OneTenth = Places \ 10
        mOutput(pos) = (Places - (OneTenth * 10)) + vbZeroChar
        Places = OneTenth
        pos = pos - 1
    Loop
End Sub

Private Sub VerifyWritable()
    If mProps.IsReadOnly Then
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_ReadOnly))
    End If
End Sub

Private Sub VerifyAndSetGroupSizes(ByRef Source() As Long, ByRef Target() As Long)
    VerifyGroupSizes Source
    
    Dim Length As Long
    Length = CorArray.Length(Source)
    
    If Length > 0 Then
        ReDim Target(0 To Length - 1)
    Else
        Target = Cor.NewLongs()
    End If
    
    CorArray.Copy Source, Target, Length
End Sub

Private Sub VerifyGroupSizes(ByRef Arr As Variant)
    Dim i As Long
    Dim ub As Long
    
    If VarType(Arr) <> (vbLong Or vbArray) Then _
        Throw Cor.NewArgumentException("Group sizes array must by of type Long", "Array")
        
    ub = UBound(Arr)
    For i = 0 To ub
        Select Case Arr(i)
            Case Is < 0, Is > 9
                Throw Cor.NewArgumentException("Group sizes must be between 0 and 9.")
            Case 0
                If i < ub Then _
                    Throw Cor.NewArgumentException("Only last group size can be 0.")
        End Select
    Next i
End Sub

Private Sub ReadLongArray(ByRef a() As Long, ByVal Name As String, ByVal Bag As PropertyBag)
    With Bag
        Dim Count As Long
        Count = .ReadProperty(Name & "_Count", 0)
        If Count = 0 Then
            a = CorArray.CreateInstance(vbLong)
            Exit Sub
        End If
        
        ReDim a(0 To Count - 1)
        
        Dim i As Long
        For i = 0 To Count - 1
            a(i) = .ReadProperty(Name & "_" & i)
        Next i
    End With
End Sub

Private Sub WriteLongArray(ByRef a() As Long, ByVal Name As String, ByVal Bag As PropertyBag)
    With Bag
        .WriteProperty Name & "_Count", UBound(a) + 1
        
        Dim i As Long
        For i = 0 To UBound(a)
            .WriteProperty Name & "_" & i, a(i)
        Next i
    End With
End Sub

Private Sub InitPatterns()
    mNumberNegativePatterns = Cor.NewStrings("(#)", "-#", "- #", "#-", "# -")
    mCurrencyNegativePatterns = Cor.NewStrings("($#)", "-$#", "$-#", "$#-", "(#$)", "-#$", "#-$", "#$-", "-# $", "-$ #", "# $-", "$ #-", "$ -#", "#- $", "($ #)", "(# $)")
    mCurrencyPositivePatterns = Cor.NewStrings("$#", "#$", "$ #", "# $")
    mPercentNegativePatterns = Cor.NewStrings("-# %", "-#%", "-%#", "%-#", "%#-", "#-%", "#%-", "-% #", "# %-", "% #-", "% -#", "#- %")
    mPercentPositivePatterns = Cor.NewStrings("# %", "#%", "%#", "% n")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With mProps
        .CurrencyDecimalDigits = PropBag.ReadProperty(PROP_CURRENCYDECIMALDIGITS)
        .CurrencyDecimalSeparator = PropBag.ReadProperty(PROP_CURRENCYDECIMALSEPARATOR)
        .CurrencyGroupSeparator = PropBag.ReadProperty(PROP_CURRENcYGROUPSEPARATOR)
        ReadLongArray .CurrencyGroupSizes, PROP_CURRENCYGROUPSIZES, PropBag
        .CurrencyNegativePattern = PropBag.ReadProperty(PROP_CURRENCYNEGATIVEPATTERN)
        .CurrencyPositivePattern = PropBag.ReadProperty(PROP_CURRENCYPOSITIVEPATTERN)
        .CurrencySymbol = PropBag.ReadProperty(PROP_CURRENCYSYMBOL)
        .NaNSymbol = PropBag.ReadProperty(PROP_NANSYMBOL)
        .NegativeInfinitySymbol = PropBag.ReadProperty(PROP_NEGATIVEINFINITYSYMBOL)
        .NegativeSign = PropBag.ReadProperty(PROP_NEGATIVESIGN)
        .NumberDecimalDigits = PropBag.ReadProperty(PROP_NUMBERDECIMALDIGITS)
        .NumberDecimalSeparator = PropBag.ReadProperty(PROP_NUMBERDECIMALSEPARATOR)
        .NumberGroupSeparator = PropBag.ReadProperty(PROP_NUMBERGROUPSEPARATOR)
        ReadLongArray .NumberGroupSizes, PROP_NUMBERGROUPSIZES, PropBag
        .NumberNegativePattern = PropBag.ReadProperty(PROP_NUMBERNEGATIVEPATTERN)
        .PercentDecimalDigits = PropBag.ReadProperty(PROP_PERCENTDECIMALDIGITS)
        .PercentDecimalSeparator = PropBag.ReadProperty(PROP_PERCENTDECIMALSEPARATOR)
        .PercentGroupSeparator = PropBag.ReadProperty(PROP_PERCENTGROUPSEPARATOR)
        ReadLongArray .PercentGroupSizes, PROP_PERCENTGROUPSIZES, PropBag
        .PercentNegativePattern = PropBag.ReadProperty(PROP_PERCENTNEGATIVEPATTERN)
        .PercentPositivePattern = PropBag.ReadProperty(PROP_PERCENTPOSITIVEPATTERN)
        .PercentSymbol = PropBag.ReadProperty(PROP_PERCENTSYMBOL)
        .PerMilleSymbol = PropBag.ReadProperty(PROP_PERMILLESYMBOL)
        .PositiveInfinitySymbol = PropBag.ReadProperty(PROP_POSITIVEINFINITYSYMBOL)
        .PositiveSign = PropBag.ReadProperty(PROP_POSITIVESIGN)
        .IsReadOnly = PropBag.ReadProperty(PROP_ISREADONLY)
    End With
    mLoaded = True
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With mProps
        PropBag.WriteProperty PROP_CURRENCYDECIMALDIGITS, .CurrencyDecimalDigits
        PropBag.WriteProperty PROP_CURRENCYDECIMALSEPARATOR, .CurrencyDecimalSeparator
        PropBag.WriteProperty PROP_CURRENcYGROUPSEPARATOR, .CurrencyGroupSeparator
        WriteLongArray .CurrencyGroupSizes, PROP_CURRENCYGROUPSIZES, PropBag
        PropBag.WriteProperty PROP_CURRENCYNEGATIVEPATTERN, .CurrencyNegativePattern
        PropBag.WriteProperty PROP_CURRENCYPOSITIVEPATTERN, .CurrencyPositivePattern
        PropBag.WriteProperty PROP_CURRENCYSYMBOL, .CurrencySymbol
        PropBag.WriteProperty PROP_NANSYMBOL, .NaNSymbol
        PropBag.WriteProperty PROP_NEGATIVEINFINITYSYMBOL, .NegativeInfinitySymbol
        PropBag.WriteProperty PROP_NEGATIVESIGN, .NegativeSign
        PropBag.WriteProperty PROP_NUMBERDECIMALDIGITS, .NumberDecimalDigits
        PropBag.WriteProperty PROP_NUMBERDECIMALSEPARATOR, .NumberDecimalSeparator
        PropBag.WriteProperty PROP_NUMBERGROUPSEPARATOR, .NumberGroupSeparator
        WriteLongArray .NumberGroupSizes, PROP_NUMBERGROUPSIZES, PropBag
        PropBag.WriteProperty PROP_NUMBERNEGATIVEPATTERN, .NumberNegativePattern
        PropBag.WriteProperty PROP_PERCENTDECIMALDIGITS, .PercentDecimalDigits
        PropBag.WriteProperty PROP_PERCENTDECIMALSEPARATOR, .PercentDecimalSeparator
        PropBag.WriteProperty PROP_PERCENTGROUPSEPARATOR, .PercentGroupSeparator
        WriteLongArray .PercentGroupSizes, PROP_PERCENTGROUPSIZES, PropBag
        PropBag.WriteProperty PROP_PERCENTNEGATIVEPATTERN, .PercentNegativePattern
        PropBag.WriteProperty PROP_PERCENTPOSITIVEPATTERN, .PercentPositivePattern
        PropBag.WriteProperty PROP_PERCENTSYMBOL, .PercentSymbol
        PropBag.WriteProperty PROP_PERMILLESYMBOL, .PerMilleSymbol
        PropBag.WriteProperty PROP_POSITIVEINFINITYSYMBOL, .PositiveInfinitySymbol
        PropBag.WriteProperty PROP_POSITIVESIGN, .PositiveSign
        PropBag.WriteProperty PROP_ISREADONLY, .IsReadOnly
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormatProvider
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormatProvider_GetFormat(ByRef FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function
