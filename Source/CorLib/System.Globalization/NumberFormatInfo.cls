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
    LCID                        As Long
End Type

Private mLoaded                 As Boolean
Private mProps                  As PropsType



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
' @remarks A clone of a read-only instance does not become read-only.
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
' @param Value The new string that represents negative infinity.
' @remarks Negative infinity is achieved by dividing a negative
' number by zero. The default is "-Infinity".
Public Property Let NegativeInfinitySymbol(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.NegativeInfinitySymbol = Value
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
' @param Value The new string representation of positive infinity.
' @remarks Positive infinity is achieved by dividing a positive
' number by zero. The default is "Infinity".
Public Property Let PositiveInfinitySymbol(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PositiveInfinitySymbol = Value
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
' @param Value The new permille string representation.
' @remarks From From <a href="http://mathworld.wolfram.com/Permil.html">MathWorld</a><br>
' The use of permille (a.k.a. parts per thousand) is a way of
' expressing ratios in terms of whole numbers. Given a ratio or fraction,
' it is converted to a permil-age by multiplying by 1000 and appending a "‰" (Unicode &H2030).
' For example, if an investment grows from a number P=13.00 to a number A=22.50,
' then A is 22.50/13.00=1.7308 times as much as P, or 1730.8‰.
Public Property Let PerMilleSymbol(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PerMilleSymbol = Value
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
' @param Value The new string NaN representation.
' @remarks Not-A-Number is achieved by dividing zero by zero. The default
' NaN symbol is "NaN".
Public Property Let NaNSymbol(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.NaNSymbol = Value
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
' @param Value The new decimal symbol.
Public Property Let PercentDecimalSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentDecimalSeparator = Value
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
' @param Value The number of digits to display after the decimal point.
' @remarks <p>This is used when a specifier of "p" is used with no
' associated digits (ie "p5").</p>
' <p>If the number is less than the precision
' of the value, then the value is rounded to fit within the number of
' digits specified. If omitted the value is rounded to the number of
' digits returned by this property.</p>
' <p>If the specified number of digits is larger than the precision of
' value, then the formatted value will be padded with zeros.</p>
' <p>The default number of digits is 2.</p>
Public Property Let PercentDecimalDigits(ByVal Value As Long)
    VerifyWritable
    VerifyLoaded
    If Value < 0 Or Value > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99))
        
    mProps.PercentDecimalDigits = Value
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
Public Property Get PercentGroupSizes() As Long()
    VerifyLoaded
    PercentGroupSizes = mProps.PercentGroupSizes
End Property

''
' Sets the group sizes for percent formatted numbers.
'
' @param Value A vbLong array containing the grouping sizes.
' @remarks A group size is the number of digits that are to be groupsed together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Let PercentGroupSizes(ByRef Value() As Long)
    VerifyWritable
    VerifyLoaded
    
    If CorArray.IsNull(Value) Then _
        Throw Cor.NewArgumentNullException("PercentGroupSizes", Environment.GetResourceString(ArgumentNull_Array))
    
    VerifyGroupSizes Value
    mProps.PercentGroupSizes = Value
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
' @param Value The pattern to be used.
'
Public Property Let PercentNegativePattern(ByVal Value As PercentNegativePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.PercentNegativePattern = Value
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
' @param Value The symbol used as a group separator defined by the group sizes.
'
Public Property Let PercentGroupSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentGroupSeparator = Value
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
' @param Value The symbol to be used in percent formatted numbers.
'
Public Property Let PercentSymbol(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PercentSymbol = Value
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
' @param Value The pattern to be used.
'
Public Property Let PercentPositivePattern(ByVal Value As PercentPositivePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.PercentPositivePattern = Value
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
' @param Value The decimal separator to be used in formatting.
'
Public Property Let CurrencyDecimalSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyDecimalSeparator = Value
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
' @param Value The default number of digits to be shown after the decimal separator.
'
Public Property Let CurrencyDecimalDigits(ByVal Value As Long)
    VerifyWritable
    VerifyLoaded
    If Value < 0 Or Value > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99))
    mProps.CurrencyDecimalDigits = Value
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
' @param Value An array of longs that are the group sizes.
' @remarks <p>Each element represents the number of digits that will be grouped
' together before a comma separator is inserted. Each element is used once
' to specify the group size, then the next element is used for the next group
' of digits, until the last group size. The last group size is used for the
' remainder of digits to be formatted.</p>
' <p>If a 0 is specified for a group size, it can only be the last size in the array.
' A 0 as the last group size will prevent the remaining digits from being separated.</p>
'
Public Property Let CurrencyGroupSizes(ByRef Value() As Long)
    VerifyWritable
    VerifyLoaded
    If CorArray.IsNull(Value) Then _
        Throw Cor.NewArgumentNullException("CurrencyGroupSizes", Environment.GetResourceString(ArgumentNull_Array))
    
    VerifyAndSetGroupSizes Value, mProps.CurrencyGroupSizes
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
' @param Value The pattern to be used when formatting.
'
Public Property Let CurrencyNegativePattern(ByVal Value As CurrencyNegativePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyNegativePattern = Value
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
' @param Value The separator to be used when separating digits.
' @remarks Generally a comma is the group separator. However, in some
' parts of the world, a comma is not used.
'
Public Property Let CurrencyGroupSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyGroupSeparator = Value
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
' @param Value The currency symbol to be used when formatting currency.
' @remarks In the US the symbol is the dollar sign '$', in other parts
' of the world, other symbols are used.
'
Public Property Let CurrencySymbol(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencySymbol = Value
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
' @param Value The formatting pattern to be used with positive currency values.
'
Public Property Let CurrencyPositivePattern(ByVal Value As CurrencyPositivePatterns)
    VerifyWritable
    VerifyLoaded
    mProps.CurrencyPositivePattern = Value
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
' @param Value The character to be used when separating groups of numbers.
' This is typically a comma.
'
Public Property Let NumberGroupSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.NumberGroupSeparator = Value
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
' @param Value The new pattern to be used.
'
Public Property Let NumberNegativePattern(ByVal Value As NumberNegativePatterns)
    VerifyWritable
    VerifyLoaded
    If Value < 0 Or Value > 4 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Index))
    mProps.NumberNegativePattern = Value
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
Public Property Get NumberGroupSizes() As Long()
    VerifyLoaded
    NumberGroupSizes = mProps.NumberGroupSizes
End Property

''
' Sets the group sizes for formatted numbers.
'
' @param Value A vbLong array containing the grouping sizes.
' @remarks A group size is the number of digits that are to be grouped together
' before being separated by the group separator symbol. Generally the group sizes
' is an array of one element having a value of 3.
' <p>The sizes can be changed to
' different sizes, such as {3,2,1}. This particular set of group sizes would
' cause a number 1234567890 to be formatted as 1,2,3,4,5,67,890. Once the last
' group size has been reached, it is repeatedly used in grouping the remaining digits.</p>
'
Public Property Let NumberGroupSizes(ByRef Value() As Long)
    VerifyWritable
    VerifyLoaded
    
    If CorArray.IsNull(Value) Then _
        Throw Cor.NewArgumentNullException("NumberGroupSizes", Environment.GetResourceString(ArgumentNull_Array))
    
    VerifyGroupSizes Value
    mProps.NumberGroupSizes = Value
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
' @param Value The number of digits to be appended after the decimal point.
'
Public Property Let NumberDecimalDigits(ByVal Value As Long)
    VerifyWritable
    VerifyLoaded
    If Value < 0 Or Value > 99 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 99))
    mProps.NumberDecimalDigits = Value
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
' @param Value The new character to be used as the decimal point.
'
Public Property Let NumberDecimalSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.NumberDecimalSeparator = Value
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
' @param Value The new negative symbol.
'
Public Property Let NegativeSign(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.NegativeSign = Value
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
' @param Value The new positive symbol.
'
Public Property Let PositiveSign(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PositiveSign = Value
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

Friend Property Get LCID() As Long
    LCID = mProps.LCID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    If Not mLoaded Then
        Load INVARIANT_LCID, False
    End If
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
        .LCID = LCID
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mProps.PerMilleSymbol = ChrW$(&H2030)
End Sub

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
