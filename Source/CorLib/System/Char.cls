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
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
        
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsWhiteSpaceStr = IsWhiteSpace(MemWord(Ptr))
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
'
Public Function IsWhiteSpace(ByVal c As Long) As Boolean
    Select Case c
        Case &H20, &HD, &H9, &HA, &HB, &HC, &H85, &HA0, &H1680, &H180E, _
             &H2000 To &H200A, _
             &H2028, &H2029, &H202F, &H205F, _
             &H3000
            IsWhiteSpace = True
    End Select
End Function

''
' Indicates where the character in the string at the specified index is a high surrogate code unit.
'
' @param s A string.
' @param Index The position of the character to evaluate in <i>s</i>.
' @return <b>true</b> if the value of <i>c</i> parameter ranges from U+D800 to U+DBFF; otherwise, <b>false</b>.
' @remarks
' <h4>Exceptions</h4>
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
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsHighSurrogateStr = IsHighSurrogate(MemWord(Ptr))
End Function

''
' Indicates whether the specified character has a high surrogate code unit.
'
' @param c The Unicode character to evaluate.
' @return <b>true</b> if the value of <i>c</i> parameter ranges from U+D800 to U+DBFF; otherwise, <b>false</b>.
' @remarks This method will handle the <b>vbInteger</b> datatype correctly if a negative value is supplied.
'
Public Function IsHighSurrogate(ByVal c As Long) As Boolean
    Select Case c
        Case &HD800& To &HDBFF&, &HFFFFD800 To &HFFFFDBFF
            IsHighSurrogate = True
    End Select
End Function

''
' Indicates where the character in the string at the specified index is a low surrogate code unit.
'
' @param s A string.
' @param Index The position of the character to evaluate in <i>s</i>.
' @return <b>true</b> if the specified character in the <i>s</i> parameter ranges from U+DC00 to U+DFFF; otherwise, <b>false</b>.
' @remarks
' <h4>Exceptions</h4>
' <table style="width:auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>index</i> is not a position within <i>s</i>.</td>
'   </tr>
' </table>
Public Function IsLowSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    If Index < 0 Or Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index"
    
    Dim Ptr As Long
    Ptr = StrPtr(s) + Index * vbSizeOfChar
    IsLowSurrogateStr = IsLowSurrogate(MemWord(Ptr))
End Function

''
' Indicates whether the specified character has a low surrogate code unit.
'
' @param c The Unicode character to evaluate.
' @return <b>true</b> if the value of <i>c</i> parameter ranges from U+DC00 to U+DFFF; otherwise, <b>false</b>.
' @remarks This method will handle the <b>vbInteger</b> datatype correctly if a negative value is supplied.
'
Public Function IsLowSurrogate(ByVal c As Long) As Boolean
    Select Case c
        Case &HDC00& To &HDFFF&, &HFFFFDC00 To &HFFFFDFFF
            IsLowSurrogate = True
    End Select
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
'
Public Function IsSurrogateStr(ByRef s As String, ByVal Index As Long) As Boolean
    If IsHighSurrogateStr(s, Index) Then
        IsSurrogateStr = True
    ElseIf IsLowSurrogateStr(s, Index) Then
        IsSurrogateStr = True
    End If
End Function

''
' Indicates whether the specified character has a surrogate code unit.
'
' @param c The Unicode character to evaluate.
' @return <b>true</b> if <i>c</i> is either a high surrogate or a low surrogate; otherwise, <b>false</b>.
' @remarks <p>A surrogate is a Char with a UTF-16 code unit in the range from U+D800 to U+DFFF. The individual surrogate code unit has no interpretation of its own, but has meaning only when used as part of a surrogate pair.</p>
' <p>For more information about surrogate pairs, see the Unicode Standard at the <a href="http://www.unicode.org">OnlineUnicode home page</a>.</p>
' @see IsSurrogateStr
' @see IsHighSurrogate
' @see IsHighSurrogateStr
' @see IsLowSurrogate
' @see IsLowSurrogateStr
'
Public Function IsSurrogate(ByVal c As Long) As Boolean
    If IsHighSurrogate(c) Then
        IsSurrogate = True
    ElseIf IsLowSurrogate(c) Then
        IsSurrogate = True
    End If
End Function
