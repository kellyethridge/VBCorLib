VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberFormatInfoStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: NumberFormatInfoStatic
'

''
' Provides methods to retrieve the various types of NumberFormatInfo objects.
'
' @remarks This class cannot be directly instantiated. To access these methods
' use the variable name directly.
' <pre>
' Dim Info As NumberFormatInfo
' Set Info = NumberFormatInfo.CurrentInfo
' </pre>
'
' @see NumberFormatInfo
'
Option Explicit

''
' Determines the styles permitted in numeric string arguments that are passed to the <b>Parse</b> and <b>TryParse</b> methods of the integral and floating-point numeric types.
'
' @param AllowCurrencySymbol Indicates that the numeric string can contain a currency symbol. Valid currency symbols are determined by the <a href="NumberFormatInfo.Get.CurrencySymbol.html">NumberFormatInfo.CurrencySymbol</a> property.
' @param AllowDecimalPoint Indicates that the numeric string can have a decimal point. If the <b>NumberStyles</b> value includes the <b>AllowCurrencySymbol</b> flag and the parsed string includes a currency symbol, the decimal separator character is determined by the <a href="NumberFormatInfo.Get.CurrencyDecimalSeparator.html">NumberFormatInfo.CurrencyDecimalSeparator</a> property. Otherwise, the decimal separator character is determined by the <a href="NumberFormatInfo.Get.NumberDecimalSeparator.html">NumberFormatInfo.NumberDecimalSeparator</a> property.
' @param AllowExponent Indicates that the numeric string can be in exponential notation. The <b>AllowExponent</b> flag allows the parsed string to contain an exponent that begins with the "E" or "e" character and that is followed by an optional positive or negative sign and an integer. In other words, it successfully parses strings in the form nnnExx, nnnE+xx, and nnnE-xx. It does not allow a decimal separator or sign in the significand or mantissa; to allow these elements in the string to be parsed, use the <b>AllowDecimalPoint</b> and <b>AllowLeadingSign</b> flags, or use a composite style that includes these individual flags.
' @param AllowHexSpecifier Indicates that the numeric string represents a hexadecimal value. Valid hexadecimal values include the numeric digits 0-9 and the hexadecimal digits A-F and a-f. Strings that are parsed using this style cannot be prefixed with "0x" or "&h". A string that is parsed with the <b>AllowHexSpecifier</b> style will always be interpreted as a hexadecimal value. The only flags that can be combined with <b>AllowHexSpecifier</b> are <b>AllowLeadingWhite</b> and <b>AllowTrailingWhite</b>. The <b>NumberStyles</b> enumeration includes a composite style, <b>HexNumber</b>, that consists of these three flags.
' @param AllowLeadingSign Indicates that the numeric string can have a leading sign. Valid leading sign characters are determined by the <a href="NumberFormatInfo.Get.PositiveSign.html">NumberFormatInfo.PositiveSign</a> and <a href="NumberFormatInfo.Get.NegativeSign.html">NumberFormatInfo.NegativeSign</a> properties.
' @param AllowLeadingWhite Indicates that leading white-space characters can be present in the parsed string. Valid white-space characters have the Unicode values U+0009, U+000A, U+000B, U+000C, U+000D, and U+0020. Note that this is a subset of the characters for which the <a href="Char.IsWhiteSpace.html">Char.IsWhiteSpace</a> method returns true.
' @param AllowParentheses Indicates that the numeric string can have one pair of parentheses enclosing the number. The parentheses indicate that the string to be parsed represents a negative number.
' @param AllowThousands Indicates that the numeric string can have group separators, such as symbols that separate hundreds from thousands. If the <b>NumberStyles</b> value includes the <b>AllowCurrencySymbol</b> flag and the string to be parsed includes a currency symbol, the valid group separator character is determined by the <a href="NumberFormatInfo.Get.CurrencyGroupSeparator.html">NumberFormatInfo.CurrencyGroupSeparator</a> property, and the number of digits in each group is determined by the <a href="NumberFormatInfo.Get.CurrencyGroupSizes.html">NumberFormatInfo.CurrencyGroupSizes</a> property. Otherwise, the valid group separator character is determined by the <a href="NumberFormatInfo.Get.NumberGroupSeparator.html">NumberFormatInfo.NumberGroupSeparator</a> property, and the number of digits in each group is determined by the <a href="NumberFormatInfo.Get.NumberGroupSizes.html">NumberFormatInfo.NumberGroupSizes</a> property.
' @param AllowTrailingSign Indicates that the numeric string can have a trailing sign. Valid trailing sign characters are determined by the <a href="NumberFormatInfo.Get.PositiveSign.html">NumberFormatInfo.PositiveSign</a> and <a href="NumberFormatInfo.Get.NegativeSign.html">NumberFormatInfo.NegativeSign</a> properties.
' @param AllowTrailingWhite Indicates that trailing white-space characters can be present in the parsed string. Valid white-space characters have the Unicode values U+0009, U+000A, U+000B, U+000C, U+000D, and U+0020. Note that this is a subset of the characters for which the <a href="Char.IsWhiteSpace.html">Char.IsWhiteSpace</a> method returns true.
' @param [Any] Indicates that all styles except <b>AllowHexSpecifier</b> are used. This is a composite number style.
' @param [Currency] Indicates that all styles except <b>AllowExponent</b> and <b>AllowHexSpecifier</b> are used. This is a composite number style.
' @param Float Indicates that the <b>AllowLeadingWhite</b>, <b>AllowTrailingWhite</b>, <b>AllowLeadingSign</b>, <b>AllowDecimalPoint</b>, and <b>AllowExponent</b> styles are used. This is a composite number style.
' @param HexNumber Indicates that the <b>AllowLeadingWhite</b>, <b>AllowTrailingWhite</b>, and <b>AllowHexSpecifier</b> styles are used. This is a composite number style.
' @param [Integer] Indicates that the <b>AllowLeadingWhite</b>, <b>AllowTrailingWhite</b>, and <b>AllowLeadingSign</b> styles are used. This is a composite number style.
' @param None Indicates that no style elements, such as leading or trailing white space, thousands separators, or a decimal separator, can be present in the parsed string. The string to be parsed must consist of integral decimal digits only.
' @param Number Indicates that the <b>AllowLeadingWhite</b>, <b>AllowTrailingWhite</b>, <b>AllowLeadingSign</b>, <b>AllowTrailingSign</b>, <b>AllowDecimalPoint</b>, and <b>AllowThousands</b> styles are used. This is a composite number style.
Public Enum NumberStyles
    None = 0
    AllowLeadingWhite = 1
    AllowTrailingWhite = 2
    AllowLeadingSign = 4
    [Integer] = 7
    AllowTrailingSign = 8
    AllowParentheses = 16
    AllowDecimalPoint = 32
    AllowThousands = 64
    Number = 111
    AllowExponent = 128
    Float = 167
    AllowCurrencySymbol = 256
    [Currency] = 383
    [Any] = 511
    AllowHexSpecifier = 512
    HexNumber = 515
End Enum

Private mCurrencyPositivePatterns() As String
Private mCurrencyNegativePatterns() As String
Private mNumberNegativePatterns()   As String
Private mPercentPositivePatterns()  As String
Private mPercentNegativePatterns()  As String


''
' Returns the NumberFormatInfo object of the threads current culture.
'
Public Property Get CurrentInfo() As NumberFormatInfo
    Set CurrentInfo = Thread.CurrentThread.CurrentCulture.NumberFormat
End Property

''
' Returns the NumberFormatInfo object of the variant culture type. This
' is the same no matter what culture is current.
'
Public Property Get InvariantInfo() As NumberFormatInfo
    Set InvariantInfo = CultureInfo.InvariantCulture.NumberFormat
End Property

''
' Returns the NumberFormatInfo object from a specific provider.
'
' @param FormatProvider The provider of the NumberFormatInfo object.
' @return A NumberFormatInfo object from the provider.
' @remarks If the provider does not provide a NumberFormatInfo object, then
' the NumberFormatInfo object for the current culture is returned.
'
Public Function GetInstance(ByVal FormatProvider As IFormatProvider) As NumberFormatInfo
    If FormatProvider Is Nothing Then
        Set GetInstance = CurrentInfo
    Else
        Set GetInstance = FormatProvider.GetFormat(vbNumberFormatInfo)
        
        If GetInstance Is Nothing Then
            Set GetInstance = CurrentInfo
        End If
    End If
End Function

''
' Returns a read-only version of the NumberFormatInfo object.
'
' @param Info The NumberFormatInfo object to create a read-only version of.
' @return A read-only version of the NumberFormatInfo object.
'
Public Function ReadOnly(ByVal Info As NumberFormatInfo) As NumberFormatInfo
    If Info Is Nothing Then _
        Throw New ArgumentNullException
    
    If Info.IsReadOnly Then
        Set ReadOnly = Info
    Else
        Set ReadOnly = Info.Clone
        ReadOnly.MarkReadOnly
    End If
End Function

Friend Property Get NumberNegativePatterns(ByVal Index As Long) As String
    NumberNegativePatterns = mNumberNegativePatterns(Index)
End Property

Friend Property Get CurrencyPositivePatterns(ByVal Index As Long) As String
    CurrencyPositivePatterns = mCurrencyPositivePatterns(Index)
End Property

Friend Property Get CurrencyNegativePatterns(ByVal Index As Long) As String
    CurrencyNegativePatterns = mCurrencyNegativePatterns(Index)
End Property

Friend Property Get PercentPositivePatterns(ByVal Index As Long) As String
    PercentPositivePatterns = mPercentPositivePatterns(Index)
End Property

Friend Property Get PercentNegativePatterns(ByVal Index As Long) As String
    PercentNegativePatterns = mPercentNegativePatterns(Index)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitPatterns
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitPatterns()
    mNumberNegativePatterns = Cor.NewStrings("(#)", "-#", "- #", "#-", "# -")
    mCurrencyNegativePatterns = Cor.NewStrings("($#)", "-$#", "$-#", "$#-", "(#$)", "-#$", "#-$", "#$-", "-# $", "-$ #", "# $-", "$ #-", "$ -#", "#- $", "($ #)", "(# $)")
    mCurrencyPositivePatterns = Cor.NewStrings("$#", "#$", "$ #", "# $")
    mPercentNegativePatterns = Cor.NewStrings("-# %", "-#%", "-%#", "%-#", "%#-", "#-%", "#%-", "-% #", "# %-", "% #-", "% -#", "#- %")
    mPercentPositivePatterns = Cor.NewStrings("# %", "#%", "%#", "% n")
End Sub
