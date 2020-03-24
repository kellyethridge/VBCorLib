VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Char"
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
' Module: Char
'

''
' Provides methods for inspecting and manipulating Unicode characters.
'
' @remarks A character in <i>VisualBasic</i> is represented as a UTF-16 code unit.
'
Option Explicit

''
' Returns minimum character value.
'
' @remarks This return the value &H0000.
'
Public Property Get MinValue() As Integer
    MinValue = 0
End Property

''
' Returns the maximum character value.
'
' @remarks This returns the value &HFFFF.
' <p>Because characters are represented using the <b>Integer</b> datatype the max
' value will be a negative value. To correctly compare character values use the
' <a href="Char.Equals.html">Char.Equals</a> and <a href="Char.Compare.html">Char.Compare</a> methods.
'
Public Property Get MaxValue() As Integer
    MaxValue = &HFFFF
End Property

''
' Compares two character code units.
'
' @param a A character to compare with.
' @param b A character to compare with.
' @return A signed number indicating the position of <i>a</i> parameter in the sort order in relation to the <i>b</i> parameter.
' <table style="width: auto">
'   <tr>
'       <th>Return Value</th>
'       <th>Description</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>a</i> precedes <i>b</i>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>a</i> has the same position in the sort order as <i>b</i></td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>a</i> follows <i>b</i> in the sort order.</td>
'    </tr>
' </table>
' @remarks This method handles the discrepancy between integers and longs particularly when
' dealing with negative integer values that should map to positive long values.
Public Function Compare(ByVal a As Long, ByVal b As Long) As Long
    If (Not IsValidChar(a)) Or (Not IsValidChar(b)) Then _
        Throw Cor.NewArgumentOutOfRangeException(IIf(IsValidChar(a), "b", "a"), Environment.GetResourceString(ArgumentOutOfRange_Range))
            
    a = a And &HFFFF&
    b = b And &HFFFF&
    
    If a < b Then
        Compare = -1
    ElseIf a > b Then
        Compare = 1
    End If
End Function

''
' Evaluates two character code units for equality.
'
' @param a A character to compare with.
' @param b A character to compare with.
' @return <b>true</b> if parameter <i>a</i> is equal to parameter <i>b</i>; otherwise <b>false</b>.
' @remarks This method handles the discrepancy between integers and longs particularly when
' dealing with negative integer values that should map to positive long values.
Public Function Equals(ByVal a As Long, ByVal b As Long) As Boolean
    If Not IsValidChar(a) Or Not IsValidChar(b) Then _
        ThrowInvalidChar IIf(IsValidChar(a), "b", "a")
        
    Equals = (a And &HFFFF&) = (b And &HFFFF&)
End Function

''
' Indicates if the specified character in a string is white space.
'
' @param s A string.
' @param Index The position of the character to evaluate in <i>s</i>.
' @return <b>true</b> if the character is considered white space; otherwise <b>false</b>.
' @remarks <p>The <i>Index</i> parameter is zero-based where as <i>VisualBasic</i> is one-based.
' <p>White space characters are the following Unicode characters:
' <ul>
'   <li>Members of the SpaceSeparator category, which includes the characters SPACE (U+0020), NO-BREAK SPACE (U+00A0), OGHAM SPACE MARK (U+1680), EN QUAD (U+2000), EM QUAD (U+2001), EN SPACE (U+2002), EM SPACE (U+2003), THREE-PER-EM SPACE (U+2004), FOUR-PER-EM SPACE (U+2005), SIX-PER-EM SPACE (U+2006), FIGURE SPACE (U+2007), PUNCTUATION SPACE (U+2008), THIN SPACE (U+2009), HAIR SPACE (U+200A), NARROW NO-BREAK SPACE (U+202F), MEDIUM MATHEMATICAL SPACE (U+205F), and IDEOGRAPHIC SPACE (U+3000).</li>
'   <li>Members of the LineSeparator category, which consists solely of the LINE SEPARATOR character (U+2028).</li>
'   <li>Members of the ParagraphSeparator category, which consists solely of the PARAGRAPH SEPARATOR character (U+2029).</li>
'   <li>The characters CHARACTER TABULATION (U+0009), LINE FEED (U+000A), LINE TABULATION (U+000B), FORM FEED (U+000C), CARRIAGE RETURN (U+000D), and NEXT LINE (U+0085).</li>
' </ul>
' </p>
Public Function IsWhiteSpaceStr(ByRef s As String, ByVal Index As Long) As Boolean
    Dim Ptr As Long
    
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
        
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsWhiteSpaceStr = CharMethods.IsWhiteSpace(MemWord(Ptr))
End Function

''
' Indicates if a character is white space.
'
' @param c An integer to check if is white space.
' @return <b>true</b> if the character is considered white space; otherwise <b>false</b>.
' @remarks White space characters are the following Unicode characters:
' <ul>
'   <li>Members of the SpaceSeparator category, which includes the characters SPACE (U+0020), NO-BREAK SPACE (U+00A0), OGHAM SPACE MARK (U+1680), EN QUAD (U+2000), EM QUAD (U+2001), EN SPACE (U+2002), EM SPACE (U+2003), THREE-PER-EM SPACE (U+2004), FOUR-PER-EM SPACE (U+2005), SIX-PER-EM SPACE (U+2006), FIGURE SPACE (U+2007), PUNCTUATION SPACE (U+2008), THIN SPACE (U+2009), HAIR SPACE (U+200A), NARROW NO-BREAK SPACE (U+202F), MEDIUM MATHEMATICAL SPACE (U+205F), and IDEOGRAPHIC SPACE (U+3000).</li>
'   <li>Members of the LineSeparator category, which consists solely of the LINE SEPARATOR character (U+2028).</li>
'   <li>Members of the ParagraphSeparator category, which consists solely of the PARAGRAPH SEPARATOR character (U+2029).</li>
'   <li>The characters CHARACTER TABULATION (U+0009), LINE FEED (U+000A), LINE TABULATION (U+000B), FORM FEED (U+000C), CARRIAGE RETURN (U+000D), and NEXT LINE (U+0085).</li>
' </ul>
' <p>This method handles the discrepancy between integers and longs particularly when
' dealing with negative integer values that should map to positive long values.</p>
Public Function IsWhiteSpace(ByVal c As Long) As Boolean
    If Not IsValidChar(c) Then _
        ThrowInvalidChar "c"
    
    IsWhiteSpace = CharMethods.IsWhiteSpace(c)
End Function

''
' Indicates where the character in the string at the specified index is a high surrogate code unit.
'
' @param s A string.
' @param Index The position of the character to evaluate in <i>s</i>.
' @return <b>true</b> if the value of <i>c</i> parameter ranges from U+D800 to U+DBFF; otherwise, <b>false</b>.
' @exceptions
' <table style="width:auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>index</i> is not a position within <i>s</i>.</td>
'   </tr>
' </table>
Public Function IsHighSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    Dim Ptr As Long

    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsHighSurrogateStr = CharMethods.IsHighSurrogate(MemWord(Ptr))
End Function

''
' Indicates whether the specified character has a high surrogate code unit.
'
' @param c The Unicode character to evaluate.
' @return <b>true</b> if the value of <i>c</i> parameter ranges from U+D800 to U+DBFF; otherwise, <b>false</b>.
' @remarks <p>This method handles the discrepancy between integers and longs particularly when
' dealing with negative integer values that should map to positive long values.</p>
Public Function IsHighSurrogate(ByVal c As Long) As Boolean
    If Not IsValidChar(c) Then _
        ThrowInvalidChar "c"
    
    IsHighSurrogate = CharMethods.IsHighSurrogate(c)
End Function

''
' Indicates where the character in the string at the specified index is a low surrogate code unit.
'
' @param s A string.
' @param Index The position of the character to evaluate in <i>s</i>.
' @return <b>True</b> if the specified character in the <i>s</i> parameter ranges from U+DC00 to U+DFFF; otherwise, <b>False</b>.
' @exceptions
' <table style="width:auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>Index</i> is not a position within <i>s</i>.</td>
'   </tr>
' </table>
Public Function IsLowSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    Dim Ptr As Long
    
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsLowSurrogateStr = CharMethods.IsLowSurrogate(MemWord(Ptr))
End Function

''
' Indicates whether the specified character has a low surrogate code unit.
'
' @param c The Unicode character to evaluate.
' @return <b>true</b> if the value of <i>c</i> parameter ranges from U+DC00 to U+DFFF; otherwise, <b>false</b>.
' @remarks <p>This method handles the discrepancy between integers and longs particularly when
' dealing with negative integer values that should map to positive long values.</p>
'
Public Function IsLowSurrogate(ByVal c As Long) As Boolean
    If Not IsValidChar(c) Then _
        ThrowInvalidChar "c"

    IsLowSurrogate = CharMethods.IsLowSurrogate(c)
End Function

''
' Indicates whether the character at the specified position in a specified string has a surrogate code unit.
'
' @param s A string.
' @param Index The position of the character to evaluate in <i>s</i>.
' @return <b>true</b> if <i>c</i> is either a high surrogate or a low surrogate; otherwise, <b>false</b>.
' @remarks <p>The <i>Index</i> parameter is zero-based where as <i>VisualBasic</i> is one-based.
' <p>A surrogate is a Char object with a UTF-16 code unit in the range from U+D800 to U+DFFF. The individual surrogate code unit has no interpretation of its own, but has meaning only when used as part of a surrogate pair.</p>
' <p>For more information about surrogate pairs, see the Unicode Standard at the <a href="http://www.unicode.org">OnlineUnicode home page</a>.</p>
' @see IsSurrogate
' @see IsHighSurrogate
' @see IsHighSurrogateStr
' @see IsLowSurrogate
' @see IsLowSurrogateStr
Public Function IsSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    Dim Ptr As Long
    
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsSurrogateStr = CharMethods.IsSurrogate(MemWord(Ptr))
End Function

''
' Indicates whether the specified character has a surrogate code unit.
'
' @param c The Unicode character to evaluate.
' @return <b>true</b> if <i>c</i> is either a high surrogate or a low surrogate; otherwise, <b>false</b>.
' @remarks <p>A surrogate is a Char with a UTF-16 code unit in the range from U+D800 to U+DFFF. The individual surrogate code unit has no interpretation of its own, but has meaning only when used as part of a surrogate pair.</p>
' <p>For more information about surrogate pairs, see the Unicode Standard at the <a href="http://www.unicode.org">OnlineUnicode home page</a>.</p>
' <p>This method handles the discrepancy between integers and longs particularly when
' dealing with negative integer values that should map to positive long values.</p>
' @see IsSurrogateStr
' @see IsHighSurrogate
' @see IsHighSurrogateStr
' @see IsLowSurrogate
' @see IsLowSurrogateStr
Public Function IsSurrogate(ByVal c As Long) As Boolean
    If Not IsValidChar(c) Then _
        ThrowInvalidChar "c"
    
    IsSurrogate = CharMethods.IsSurrogate(c)
End Function

''
' Converts the value of a UTF-16 encoded surrogate pair into a Unicode code point.
'
' @param HighSurrogate A high surrogate code unit (that is, a code unit ranging from U+D800 through U+DBFF).
' @param LowSurrogate A low surrogate code unit (that is, a code unit ranging from U+DC00 through U+DFFF).
' @return The 21-bit Unicode code point represented by the highSurrogate and lowSurrogate parameters.
' @exceptions
' <table style="width:auto">
'   <tr><th>Exception</th><th>Condition</th></tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>HighSurrogate</i> is not in the range U+D800 through U+DBFF, or <i>LowSurrogate</i> is not in the range U+DC00 through U+DFFF.</td>
'   </tr>
' </table>
Public Function ConvertToUtf32(ByVal HighSurrogate As Long, ByVal LowSurrogate As Long) As Long
    If Not IsValidChar(HighSurrogate) Or Not IsValidChar(LowSurrogate) Then _
        ThrowInvalidChar IIf(IsValidChar(HighSurrogate), "LowSurrogate", "HighSurrogate")
    If Not CharMethods.IsHighSurrogate(HighSurrogate) Then _
        Error.ArgumentOutOfRange "HighSurrogate", ArgumentOutOfRange_InvalidHighSurrogate
    If Not CharMethods.IsLowSurrogate(LowSurrogate) Then _
        Error.ArgumentOutOfRange "LowSurrogate", ArgumentOutOfRange_InvalidLowSurrogate

    ConvertToUtf32 = CharMethods.ConvertToUtf32(HighSurrogate, LowSurrogate)
End Function

''
' Converts the value of a UTF-16 encoded character or surrogate pair at a specified position in a string into a Unicode code point.
'
' @param s A string that contains a character or surrogate pair.
' @param Index The index position of the character or surrogate pair in s.
' @return The 21-bit Unicode code point represented by the character or surrogate pair at the position in the <i>s</i> parameter specified by the <i>index</i> parameter.
' @exceptions
' <table style="width:auto">
'   <tr><th>Exception</th><th>Condition</th></tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>index</i> is not a position within <i>s</i>.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td>
'       <td>The specified index position contains a surrogate pair, and either the first character in the pair is not a valid high surrogate or the second character in the pair is not a valid low surrogate.</td>
'   </tr>
' </table>
Public Function ConvertToUtf32Str(ByRef s As String, ByVal Index As Long) As Long
    Dim Char1   As Long
    Dim Char2   As Long
    Dim Ptr     As Long
    
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    Char1 = MemWord(Ptr) And &HFFFF&
    
    If IsLowSurrogate(Char1) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidLowSurrogate, Index), "s")
    
    If IsHighSurrogate(Char1) Then
        Char2 = MemWord(Ptr + vbSizeOfChar) And &HFFFF&
                
        If Not IsLowSurrogate(Char2) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidHighSurrogate, Index), "s")
        
        ConvertToUtf32Str = CharMethods.ConvertToUtf32(Char1, Char2)
    Else
        ConvertToUtf32Str = Char1
    End If
End Function

''
' Converts the specified Unicode code point into a UTF-16 encoded string.
'
' @param Utf32 A 21-bit Unicode code point.
' @return A string consisting of one character or a surrogate pair of characters equivalent to the code point specified by the <i>Utf32</i> parameter.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td><i>Utf32</i> is not a valid 21-bit Unicode code point ranging from U+0 through U+10FFFF, excluding the surrogate pair range from U+D800 through U+DFFF.</td>
'   </tr>
' </table>
Public Function ConvertFromUtf32(ByVal UTF32 As Long) As String
    Dim High As Long
    Dim Low As Long
    
    If (UTF32 < 0 Or UTF32 > UnicodePlane16End) Or (UTF32 >= SurrogateStart And UTF32 <= SurrogateEnd) Then _
        Error.ArgumentOutOfRange "Utf32", ArgumentOutOfRange_InvalidUTF32
    
    If UTF32 < CharMethods.UnicodePlane1Start Then
        ConvertFromUtf32 = ChrW$(UTF32)
    Else
        UTF32 = UTF32 - UnicodePlane1Start
        High = (UTF32 And &HFFC00) \ vbShift10Bits + HighSurrogateStart
        Low = (UTF32 And &H3FF&) + LowSurrogateStart
        ConvertFromUtf32 = ChrW$(High) & ChrW$(Low)
    End If
End Function

''
' Converts the character value to a 32-bit signed integer.
'
' @param c The character to be converted.
' @return A 32-bit signed integer representing the numeric value of the character.
' @remarks In VBCorLib characters are represented with 16-bit signed integers. In order
' to represent all characters, the numeric value may be represented with a negative value.
' <p>If the numeric value is negative, converted positive value will have the same 16-bit binary representation.</p>
' <p>A character value of -1(&hFFFF) is converted to 65535(&hFFFF&). Both values have the same 16-bit binary representation.</p>
Public Function ConvertToInt32(ByVal c As Integer) As Long
    ConvertToInt32 = c And &HFFFF&
End Function

''
' Converts a 32-bit signed integer value to a 16-bit signed integer as a character.
'
' @param Value The 32-bit signed integer value to be converted to a corrisponding 16-bit integer character.
' @return A 16-bit signed integer character.
' @remarks The 16-bit result will be a binary equivalent of the first 16 bits of the <i>Value</i> parameter.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td><i>Value</i> is less than -32768 or greater than 65535.</td>
'   </tr>
' </table>
Public Function ConvertFromInt32(ByVal Value As Long) As Integer
    If Not IsValidChar(Value) Then _
        ThrowInvalidChar "Value"

    ConvertFromInt32 = Value And &H7FFF&
    
    If Value And &H8000& Then
        ConvertFromInt32 = ConvertFromInt32 Or &H8000
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ThrowInvalidChar(ByVal ParamName As String)
    Throw Cor.NewArgumentOutOfRangeException(ParamName, Environment.GetResourceString(ArgumentOutOfRange_Range, MinCharValue, MaxCharValue))
End Sub
