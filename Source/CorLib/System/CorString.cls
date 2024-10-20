VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: CorString
'

''
' Provides static functions to format and manipulate string values.
'
' @remarks This class cannot be directly instantiated. The methods are accessed through the class name itself.
' @include "..\DocIncludes\CorString.html"
Option Explicit

''
' Specifies the culture, case, and sort rules to be used with certain implementations of the
' <a href="CorString.Compare.html">CorString.Compare</a> and <a href="CorString.Equals.html">CorString.Equals</a> methods.
'
' @param CurrentCulture Compare strings using culture-sensitive sort rules and the current culture.
' @param CurrentCultureIgnoreCase Compare strings using culture-sensitive sort rules, the current culture, and ignoring the case of the strings being compared.
' @param InvariantCulture Compare strings using culture-sensitive sort rules and the invariant culture.
' @param InvariantCultureIgnoreCase Compare strings using culture-sensitive sort rules, the invariant culture, and ignoring the case of the strings being compared.
' @param Ordinal Specifies a case-sensitive string comparison.
' @param OrdinalIgnoreCase Specifies a case-insensitive string comparison.
' @param BinaryCompare Uses the Visual Basic <b>StrComp</b> function with <b>vbBinaryCompare</b> option.
' @param TextCompare Uses the Visual Basic <b>StrComp</b> function with <b>vbTextCompare</b> option.
Public Enum StringComparison
    CurrentCulture = 0
    CurrentCultureIgnoreCase = 1
    InvariantCulture = 2
    InvariantCultureIgnoreCase = 3
    Ordinal = 4
    OrdinalIgnoreCase = 5
    BinaryCompare = 100
    TextCompare = 101
End Enum

''
' Defines the string comparison options to use with certain implementations of the
' <a href="CorString.Compare.html">CorString.Compare</a> and <a href="CorString.Equals.html">CorString.Equals</a> methods.
' <p>This enumeration allows a bitwise combination of its member values.</p>
'
' @param None Indicates the default option settings for string comparisons.
' @param IgnoreCase Indicates that the string comparison must ignore case.
' @param IgnoreNonSpace Indicates that the string comparison must ignore nonspacing combining characters, such as diacritics.
' The <a href="https://home.unicode.org/">Online Unicode Standard</a> defines combining characters as characters that are
' combined with base characters to produce a new character. Nonspacing combining characters do not occupy a spacing position by themselves when rendered.
' @param IgnoreSymbols Indicates that the string comparison must ignore symbols, such as white-space characters, punctuation, currency symbols, the percent sign, mathematical symbols, the ampersand, and so on.
' @param IgnoreKanaType Indicates that the string comparison must ignore the Kana type. Kana type refers to Japanese hiragana and katakana characters,
' which represent phonetic sounds in the Japanese language. Hiragana is used for native Japanese expressions and words, while katakana is used for
' words borrowed from other languages, such as "computer" or "Internet". A phonetic sound can be expressed in both hiragana and katakana. If this value
' is selected, the hiragana character for one sound is considered equal to the katakana character for the same sound.
' @param IgnoreWidth Indicates that the string comparison must ignore the character width. For example, Japanese katakana
' characters can be written as full-width or half-width. If this value is selected, the katakana characters written as full-width
' are considered equal to the same characters written as half-width
' @param OrdinalIgnoreCaseOption String comparison must ignore case, then perform an ordinal comparison. This technique is
' equivalent to converting the string to uppercase using the invariant culture and then performing an ordinal comparison on the result.
' @param StringSort Indicates that the string comparison must use the string sort algorithm. In a string sort, the hyphen and the apostrophe, as well as other nonalphanumeric symbols, come before alphanumeric characters.
' @param OrdinalOption Indicates that the string comparison must use successive Unicode UTF-16 encoded values of the string
' (code unit by code unit comparison), leading to a fast comparison but one that is culture-insensitive. A string starting with a
' code unit \uXXXX comes before a string starting with \uYYYY, if \uXXXX is less than \uYYYY.
' <p>This value cannot be combined with other CompareOptions values and must be used alone.</p>
Public Enum CompareOptions
    None = 0
    IgnoreCase = &H1
    IgnoreNonSpace = &H2
    IgnoreSymbols = &H4
    IgnoreKanaType = &H8
    IgnoreWidth = &H10
    OrdinalIgnoreCaseOption = &H10000000
    StringSort = &H20000000
    OrdinalOption = &H40000000
End Enum


''
' Specifies whether <a href="CorString.Split.html">CorString.Split</a> method include or omit empty substrings from the return value.
'
' @param None The return value includes array elements that contain an empty string.
' @param RemoveEmptyEntries The return value does not include array elements that contain an empty string.
Public Enum StringSplitOptions
    None
    RemoveEmptyEntries
End Enum

Private mCharCheckerDelegate    As Delegate
Private mLowerUpper()           As Integer
Private mString1()              As Integer
Private mString1SA              As SafeArray1d
Private mString2()              As Integer
Private mString2SA              As SafeArray1d
Private mString1Lng()           As Long
Private mString1LngSA           As SafeArray1d
Private mString2Lng()           As Long
Private mString2LngSA           As SafeArray1d


''
' Returns if the string is null.
'
' @param s The string to be tested for null.
' @return <b>True</b> if the string pointer is null, <b>False</b> otherwise.
' @include "..\DocIncludes\CorString.IsNull.html"
Public Function IsNull(ByRef S As String) As Boolean
    IsNull = (StrPtr(S) = vbNullPtr)
End Function

''
' Returns if the string is null, empty or contains only white-space characters.
'
' @param s The string to be tested for white-space.
' @return <b>True</b> if the string has a length of zero or contains only white-space, <b>False</b> otherwise.
' @include "..\DocIncludes\CorString.IsNullOrWhiteSpace.html"
Public Function IsNullOrWhiteSpace(ByRef S As String) As Boolean
    Dim Chars() As Integer
    Dim I As Long
    
    If LenB(S) = 0 Then
        IsNullOrWhiteSpace = True
    Else
        Chars = AllocChars(S)
        
        For I = 0 To Len(S) - 1
            If Not IsWhiteSpace(Chars(I)) Then
                GoTo Finally
            End If
        Next

        IsNullOrWhiteSpace = True
    End If
    
Finally:
    FreeChars Chars
End Function

''
' Replaces the format item in a specified string with the string representation of a corresponding value in a specified array.
'
' @param FormatText A composite format string (see Remarks).
' @param args An argument array that contains zero or more values to format.
' @return A copy of <i>FormatText</i> in which the format items have been replaced by the string representation of the corresponding values in <i>Args</i>.
' @see #FormatArray
' @see #FormatEx
' @see #FormatArrayEx
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\DocIncludes\CorString.Format.html"
Public Function Format(ByRef FormatText As String, ParamArray Args() As Variant) As String
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    Format = FormatCore(Nothing, FormatText, SwappedArgs)
End Function

''
' Replaces the format item in a specified string with the string representation of a corresponding object in a specified array. A specified parameter supplies culture-specific formatting information.
'
' @param Provider An object that supplies culture-specific formatting information.
' @param FormatText A composite format string (see Remarks).
' @param Args An argument array that contains zero or more values to format.
' @return A copy of <i>FormatText</i> in which the format items have been replaced by the string representation of the corresponding values in <i>Args</i>.
' @see #Format
' @see #FormatArray
' @see #FormatArrayEx
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\DocIncludes\CorString.FormatEx.html"
Public Function FormatEx(ByVal Provider As IFormatProvider, ByRef FormatText As String, ParamArray Args() As Variant) As String
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(20)
    FormatEx = FormatCore(Provider, FormatText, SwappedArgs)
End Function

''
' Replaces the format item in a specified string with the string representation of a corresponding value in a specified array.
'
' @param FormatText A composite format string (see Remarks).
' @param args An argument array that contains zero or more values to format.
' @return A copy of <i>FormatText</i> in which the format items have been replaced by the string representation of the corresponding values in <i>Args</i>.
' @see #Format
' @see #FormatEx
' @see #FormatArrayEx
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\DocIncludes\CorString.Format.html"
Public Function FormatArray(ByRef FormatText As String, ByRef Args() As Variant) As String
    FormatArray = FormatCore(Nothing, FormatText, Args)
End Function

''
' Replaces the format item in a specified string with the string representation of a corresponding object in a specified array. A specified parameter supplies culture-specific formatting information.
'
' @param Provider An object that supplies culture-specific formatting information.
' @param FormatText A composite format string (see Remarks).
' @param Args An argument array that contains zero or more values to format.
' @return A copy of <i>FormatText</i> in which the format items have been replaced by the string representation of the corresponding values in <i>Args</i>.
' @see #Format
' @see #FormatArray
' @see #FormatEx
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\DocIncludes\CorString.FormatEx.html"
Public Function FormatArrayEx(ByVal Provider As IFormatProvider, ByRef FormatText As String, ByRef Args() As Variant) As String
    FormatArrayEx = FormatCore(Provider, FormatText, Args)
End Function

''
' Creates and array of chars (Integers) from the specified string.
'
' @param s The string to create the chars from.
' @param StartIndex The start index in the string to begin converting to a char array. This is zero-based.
' @param Length The number of characters to convert to chars.
' @return An array containing the converted characters from the string.
' @remarks The <i>StartIndex</i> parameter is zero-based. That is, the index of the first character in the string <i>s</i> is zero.
' <p>If <i>Length</i> is zero, the returned array is empty and has a zero length. If <i>s</i> is an empty string (""),
' the returned array is empty and has a zero length.</p>
' @include "..\DocIncludes\CorString.ToCharArray.html"
Public Function ToCharArray(ByRef S As String, Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As Integer()
    Dim Range   As ListRange
    Dim Chars() As Integer
    
    Range = MakeDefaultRange(StartIndex, 0, Length, Len(S), NameOfStartIndex, NameOfLength)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > Len(S) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_IndexCount
    
    If Range.Count > 0 Then
        ReDim Chars(0 To Range.Count - 1)
        CopyMemory Chars(0), ByVal StrPtr(S) + Range.Index * vbSizeOfChar, Range.Count * vbSizeOfChar
    Else
        Chars = Cor.NewIntegers()
    End If
    
    ToCharArray = Chars
End Function

''
' Returns a new string that right-aligns the characters in this instance by padding them on the left with a specified Unicode character, for a specified total length.
'
' @param s The string to be padded.
' @param TotalWidth The number of characters in the resulting string, equal to the number of original characters plus any additional padding characters.
' @param PaddingChar A Unicode padding character. <b>AscW</b> and <b>Chrw$</b> values are accepted.
' @return A new string that is equivalent to <i>s</i>, but right-aligned and padded on the left with as many <i>PaddingChar</i> characters as needed to create
' a length of <i>TotalWidth</i>. However, if <i>TotalWidth</i> is less than the length of <i>s</i>, the method returns <i>s</i>.
' If <i>TotalWidth</i> is equal to the length of <i>s</i>, the method returns a new string that is identical to <i>s</i>.
' @include "..\DocIncludes\CorString.PadLeft.html"
Public Function PadLeft(ByRef S As String, ByVal TotalWidth As Long, Optional ByRef PaddingChar As Variant = 32) As String
    Dim Length As Long
    Length = Len(S)
    If TotalWidth > Length Then
        PadLeft = String$(TotalWidth - Length, PaddingChar) & S
    Else
        PadLeft = S
    End If
End Function

''
' Returns a new string that left-aligns the characters in this string by padding them on the right with a specified Unicode character, for a specified total length.
'
' @param s The string to be padded.
' @param TotalWidth The number of characters in the resulting string, equal to the number of original characters plus any additional padding characters.
' @param PaddingChar A Unicode padding character. <b>AscW</b> and <b>Chrw$</b> values are accepted.
' @return A new string that is equivalent to <i>s</i>, but left-aligned and padded on the right with as many <i>PaddingChar</i> characters as needed to create
' a length of <i>TotalWidth</i>. However, if <i>TotalWidth</i> is less than the length of <i>s</i>, the method returns <i>s</i>.
' If <i>TotalWidth</i> is equal to the length of <i>s</i>, the method returns a new string that is identical to <i>s</i>.
' @include "..\DocIncludes\CorString.PadRight.html"
Public Function PadRight(ByRef S As String, ByVal TotalWidth As Long, Optional ByRef PaddingChar As Variant = 32) As String
    Dim Length As Long
    Length = Len(S)
    If TotalWidth > Length Then
        PadRight = S & String$(TotalWidth - Length, PaddingChar)
    Else
        PadRight = S
    End If
End Function

''
' Removes all leading occurrences of a set of characters specified in an array from the <i>s</i> parameter.
'
' @param s The string to be trimmed.
' @param TrimChars An array of Unicode characters to remove. <b>String</b> and <b>Integer()</b> are accepted.
' @return The string that remains after all occurrences of characters in the <i>TrimChars</i> parameter are removed from the start of <i>s</i>. If <i>TrimChars</i> is missing, white-space characters are removed instead.
' @see #Trim
' @see #TrimEnd
' @see #ToCharArray
' @include "..\DocIncludes\CorString.TrimStart.html"
Public Function TrimStart(ByRef S As String, Optional ByRef TrimChars As Variant) As String
    TrimStart = TrimCore(S, TrimChars, True, False)
End Function

''
' Removes all trailing occurrences of a set of characters specified in an array from the <i>s</i> parameter.
'
' @param s The string to be trimmed.
' @param TrimChars An array of Unicode characters to remove. <b>String</b> and <b>Integer()</b> are accepted.
' @return The string that remains after all occurrences of characters in the <i>TrimChars</i> parameter are removed from the end of <i>s</i>. If <i>TrimChars</i> is missing, white-space characters are removed instead.
' @see #Trim
' @see #TrimStart
' @see #ToCharArray
' @include "..\DocIncludes\CorString.TrimEnd.html"
Public Function TrimEnd(ByRef S As String, Optional ByRef TrimChars As Variant) As String
    TrimEnd = TrimCore(S, TrimChars, False, True)
End Function

''
' Removes all leading and trailing occurrences of a set of characters specified in an array from the <i>s</i> parameter.
'
' @param s The string to be trimmed.
' @param TrimChars An array of Unicode characters to remove. <b>String</b> and <b>Integer()</b> are accepted.
' @return The string that remains after all occurrences of characters in the <i>TrimChars</i> parameter are removed from the start and end of <i>s</i>. If <i>TrimChars</i> is missing, white-space characters are removed instead.
' @see #TrimStart
' @see #TrimEnd
' @see #ToCharArray
' @include "..\DocIncludes\CorString.Trim.html"
Public Function Trim(ByRef S As String, Optional ByRef TrimChars As Variant) As String
    Trim = TrimCore(S, TrimChars, True, True)
End Function

''
' Concatenates the specified elements of an array or collection, using the specified separator between each element.
'
' @param Separator The string to use as a saparator.
' @param Value An array or collection that contains the elements to concatenate.
' @param StartIndex The first element in <i>Value</i> to use.
' @param Count The number of elements of <i>Value</i> to use.
' @return A string that consists of the strings in value delimited by the separator string.<br/>
' -or-<br/>
' "" if <i>Count</i> is zero, <i>Value</i> has no elements, or <i>Separator</i> and all the elements of <i>Value</i> are "".
' @see #Split
' @include "..\DocIncludes\CorString.Join.html"
Public Function Join(ByRef Separator As String, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As String
    Dim Range As ListRange
    
    If IsObject(Value) Then
        Join = JoinCollection(Separator, Value)
        Exit Function
    End If
    
    ValidateArray Value, Parameter_Value
    Range = MakeDefaultRange(StartIndex, LBound(Value), Count, Len1D(Value), Parameter_StartIndex)
    
    If Range.Index < LBound(Value) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Value) Then _
        Error.ArgumentOutOfRange "", Argument_InvalidOffLen
    
    Select Case VarType(Value)
        Case vbStringArray
            Join = JoinStrings(Separator, Value, Range)
        Case vbVariantArray
            Join = JoinVariants(Separator, Value, Range)
        Case Else
            Join = JoinArray(Separator, Value, Range)
    End Select
End Function

''
' Copies a specified number of characters from a specified position in a string to a specified position in an array of Unicode characters.
'
' @param Source The string to copy the characters from.
' @param SourceIndex The zero-based index of the first character to copy.
' @param Destination An array of Unicode characters to which characters in the string are copied.
' @param DestinationIndex The index in <i>Destination</i> as which the copy operation begins.
' @param Count The number of characters in the string to copy to <i>Destination</i>.
' @include "..\DocIncludes\CorString.CopyTo.html"
Public Sub CopyTo(ByRef Source As String, ByVal SourceIndex As Long, ByRef Destination() As Integer, ByVal DestinationIndex As Long, ByVal Count As Long)
    If CorArray.IsNull(Destination) Then _
        Error.ArgumentNull "Destination", ArgumentNull_Array
    If SourceIndex < 0 Then _
        Error.ArgumentOutOfRange "SourceIndex", ArgumentOutOfRange_Index
    If DestinationIndex < LBound(Destination) Then _
        Error.ArgumentOutOfRange "DestinationIndex", ArgumentOutOfRange_LBound
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NegativeCount
    If SourceIndex + Count > Len(Source) Then _
        Error.ArgumentOutOfRange "SourceIndex", ArgumentOutOfRange_IndexCount
    If DestinationIndex + Count - 1 > UBound(Destination) Then _
        Error.ArgumentOutOfRange "DestinationIndex", ArgumentOutOfRange_UBound
    
    CopyMemory Destination(DestinationIndex), ByVal StrPtr(Source) + SourceIndex * vbSizeOfChar, Count * vbSizeOfChar
End Sub

''
' Returns a new string in which a specified number of characters beginning at a specified position have been deleted.
'
' @param s The string to remove characters from.
' @param StartIndex The zero-based position to begin deleting characters.
' @param Count The number of characters to delete.
' @return A new string with the substring deleted.
' @see #Insert
' @include "..\DocIncludes\CorString.Remove.html"
Public Function Remove(ByRef S As String, ByVal StartIndex As Long, Optional ByRef Count As Variant) As String
    Dim CharCount       As Long
    Dim RemovalCount    As Long
    Dim SPtr As Long
    
    CharCount = Len(S)
    RemovalCount = CLngOrDefault(Count, CharCount - StartIndex)
    
    If StartIndex < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If RemovalCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If StartIndex + RemovalCount > Len(S) Then _
        Error.ArgumentOutOfRange "", Argument_InvalidOffLen
    
    SPtr = StrPtr(S)
    Remove = SysAllocStringLen(SPtr, CharCount - RemovalCount)
    CopyMemory ByVal StrPtr(Remove) + StartIndex * vbSizeOfChar, ByVal SPtr + (StartIndex + RemovalCount) * vbSizeOfChar, (CharCount - (StartIndex + RemovalCount)) * vbSizeOfChar
End Function

''
' Reports the zero-based index of the first occurrence in <i>s</i> of any character in a specified array of Unicode characters.
' The search starts at a specified character position and examines a specified number of character positions.
'
' @param s The string to be searched.
' @param OfAny A Unicode character array containing one or more characters to seek. This accepts <b>String</b> or <b>Integer()</b>.
' @param StartIndex The zero-based starting index of <i>s</i> to begin searching.
' @param Count The number of characters in <i>s</i> to search.
' @return The zero-based index position of the first occurrence in <i>s</i> where any character in <i>AnyOf</i> was found; -1 if no character in <i>AnyOf</i> was found.
' @see #LastIndexOfAny
' @include "..\DocIncludes\CorString.IndexOfAny.html"
Public Function IndexOfAny(ByRef S As String, ByRef OfAny As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    Range = MakeDefaultStepRange(StartIndex, 0, Count, Len(S), NameOfStartIndex)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index > Len(S) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > Len(S) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Count
    
    If Len(S) > 0 Then
        IndexOfAny = IndexOfAnyCore(S, OfAny, Range.Index, Range.Index + Range.Count - 1, 1)
    Else
        IndexOfAny = -1
    End If
End Function

''
' Reports the zero-based index of the first occurrence in <i>s</i> of any character in a specified array of Unicode characters.
' The search starts at a specified character position and examines a specified number of character positions.
'
' @param s The string to be searched.
' @param OfAny A Unicode character array containing one or more characters to seek. This accepts <b>String</b> or <b>Integer()</b>.
' @param StartIndex The search starting position. The search proceeds from <i>StartIndex</i> toward the beginning of <i>s</i>.
' @param Count The number of characters in <i>s</i> to examine.
' @return The index position of the last occurrence in <i>s</i> where any character in <i>AnyOf</i> was found; -1 if no character in <i>AnyOf</i> was found or if <i>s</i> is empty.
' @see #IndexOfAny
' @include "..\DocIncludes\CorString.LastIndexOfAny.html"
Public Function LastIndexOfAny(ByRef S As String, ByRef OfAny As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    If Len(S) > 0 Then
        Range = MakeStringRange(S, StartIndex, Count)
        
        If Range.Index < 0 Then _
            Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
        If Range.Count < 0 Then _
            Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
        If Range.Index - Range.Count + 1 < 0 Then _
            Error.ArgumentOutOfRange "", ArgumentOutOfRange_Count
        
        LastIndexOfAny = IndexOfAnyCore(S, OfAny, Range.Index, Range.Index - Range.Count + 1, -1)
    Else
        LastIndexOfAny = -1
    End If
End Function

''
' Returns a new string in which a specified string is inserted at a specified index position in <i>s</i>.
'
' @param s The string to insert to.
' @param StartIndex The zero-base index position of the insertion.
' @param Value The string to insert.
' @return A new string that is equivalent to this instance, but with <i>Value</i> inserted at position <i>StartIndex</i>.
' @see #Remove
' @include "..\DocIncludes\CorString.Insert.html"
Public Function Insert(ByRef S As String, ByVal StartIndex As Long, ByRef Value As String) As String
    Dim ValueLength As Long
    Dim SourcePtr As Long
    Dim TargetPtr As Long
       
    If StartIndex < 0 Or StartIndex > Len(S) Then _
        Error.ArgumentOutOfRange "StartIndex"
    
    ValueLength = Len(Value)
    
    If ValueLength = 0 Then
        Insert = S
        Exit Function
    End If
    
    SourcePtr = StrPtr(S)
    Insert = SysAllocStringLen(ByVal 0&, ValueLength + Len(S))
    TargetPtr = StrPtr(Insert)
    ValueLength = ValueLength * vbSizeOfChar
    StartIndex = StartIndex * vbSizeOfChar
    
    CopyMemory ByVal TargetPtr, ByVal SourcePtr, StartIndex
    CopyMemory ByVal TargetPtr + StartIndex, ByVal StrPtr(Value), ValueLength
    CopyMemory ByVal TargetPtr + StartIndex + ValueLength, ByVal SourcePtr + StartIndex, LenB(Insert) - (StartIndex + ValueLength)
End Function

''
' Determines whether the beginning of a string matches the specified string when compared using the specified comparison option.
'
' @param s The string to check.
' @param Value The string to compare to the substring at the beginning of <i>s</i>.
' @param ComparisonType One of the enumeration values that determines how <i>s</i> and <i>Value</i> are compared.
' @return <b>True</b> if the <i>s</i> parameter begins with <i>Value</i>; otherwise <b>False</b>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>ComparisonType</i> is not a <a href="CorString.StringComparison.html">StringComparison</a> value.</td>
'   </tr>
' </table>
Public Function StartsWith(ByRef S As String, ByRef Value As String, Optional ByVal ComparisonType As StringComparison = StringComparison.Ordinal) As Boolean
    Dim LenValue As Long
    
    LenValue = Len(Value)
    
    If LenValue <= Len(S) Then
        StartsWith = CompareHelper(StrPtr(S), LenValue, StrPtr(Value), LenValue, ComparisonType) = 0
    End If
End Function

''
' Determines whether the beginning of this string instance matches the specified string when compared using the specified culture.
'
' @param s The string to check.
' @param Value The string to compare to the substring at the beginning of <i>s</i>.
' @param IgnoreCase <b>True</b> to ignore case during comparison; otherwise <b>False</b>.
' @param Culture Cultural information that determines how <i>s</i> and <i>Value</i> are compared. If <i>Culture</i> is <b>Nothing</b>, the current culture is used.
' @return <b>True</b> if the <i>s</i> parameter begins with <i>Value</i>; otherwise <b>False</b>.
Public Function StartsWithEx(ByRef S As String, ByRef Value As String, ByVal IgnoreCase As Boolean, ByVal Culture As CultureInfo) As Boolean
    Dim LenValue As Long
    
    If Culture Is Nothing Then
        Set Culture = CultureInfo.CurrentCulture
    End If
    
    LenValue = Len(Value)
    
    If LenValue <= Len(S) Then
        StartsWithEx = CompareCulturalHelper(StrPtr(S), LenValue, StrPtr(Value), LenValue, Culture, IIfLong(IgnoreCase, CompareOptions.IgnoreCase, CompareOptions.None)) = 0
    End If
End Function

''
' Determines whether the end of a string matches the specified string when compared using the specified comparison option.
'
' @param s The string to check.
' @param Value The string to compare to the substring at the end of <i>s</i>.
' @param ComparisonType One of the enumeration values that determines how <i>s</i> and <i>Value</i> are compared.
' @return <b>True</b> if the <i>Value</i> parameter matches the end of <i>s</i>; otherwise <b>False</b>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>ComparisonType</i> is not a <a href="CorString.StringComparison.html">StringComparison</a> value.</td>
'   </tr>
' </table>
Public Function EndsWith(ByRef S As String, ByRef Value As String, Optional ByVal ComparisonType As StringComparison = StringComparison.Ordinal) As Boolean
    Dim StartIndex  As Long
    Dim LenValue    As Long
    
    LenValue = Len(Value)
    
    If LenValue <= Len(S) Then
        StartIndex = (Len(S) - LenValue) * vbSizeOfChar
        EndsWith = CompareHelper(StrPtr(S) + StartIndex, LenValue, StrPtr(Value), LenValue, ComparisonType) = 0
    End If
End Function

''
' Determines whether the end of this string instance matches the specified string when compared using the specified culture.
'
' @param s The string to check.
' @param Value The string to compare to the substring at the end of the <i>s</i> parameter.
' @param IgnoreCase <b>True</b> to ignore case during comparison; otherwise <b>False</b>.
' @param Culture Cultural information that determines how <i>s</i> and <i>Value</i> are compared. If <i>Culture</i> is <b>Nothing</b>, the current culture is used.
' @return <b>True</b> if the <i>Value</i> parameter matches the end of <i>s</i>; otherwise <b>False</b>.
Public Function EndsWithEx(ByRef S As String, ByRef Value As String, ByVal IgnoreCase As Boolean, ByVal Culture As CultureInfo) As Boolean
    Dim StartIndex  As Long
    Dim LenValue    As Long
    Dim Options     As CompareOptions
    
    If Culture Is Nothing Then
        Set Culture = CultureInfo.CurrentCulture
    End If
    
    LenValue = Len(Value)
    Options = IIfLong(IgnoreCase, CompareOptions.IgnoreCase, CompareOptions.None)
    
    If LenValue <= Len(S) Then
        StartIndex = (Len(S) - LenValue) * vbSizeOfChar
        EndsWithEx = CompareCulturalHelper(StrPtr(S) + StartIndex, LenValue, StrPtr(Value), LenValue, Culture, Options) = 0
    End If
End Function

''
' Splits a string into a maximum number of substrings based on the strings or characters in an array. You can specify whether the substrings include empty array elements.
'
' @param s The string to split.
' @param Separator A String or Integer array that delimits the substrings in the <i>s</i> parameter, an empty array that contains no delimiters, or null.
' @param Count The maximum number of substrings to return.
' @param Options <a href="CorString.StringSplitOptions.html">StringSplitOptions.RemoveEmptyEntries</a> to omit empty array elements from the array returned; or <a href="CorString.StringSplitOptions.html">StringSplitOptions.None</a> to include empty array elements in the array returned.
' @return An array whose elements contain the substrings in this string that are delimited by one or more strings in <i>Separator</i>. For more information, see the Remarks section.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Count</i> is negative.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Options</i> is not one of the <a href="CorString.StringSplitOptions.html">StringSplitOptions</a> values.</td>
'    </tr>
' </table>
' @remarks <p>Delimiter strings are not included in the elements of the returned array.</p>
' <p>If the <i>s</i> parameter does not contain any of the strings or characters in <i>Separator</i>, or the <i>Count</i> parameter is 1, the returned array consists of a single element
' that contains the <i>s</i> parameter. If the <i>Separator</i> parameter is <b>null</b> or contains no elements, white-space characters are assumed to be the delimiters.
' White-space characters are defined by the Unicode standard and return <b>True</b> if they are passed to the <a href="Char.IsWhiteSpace.html">Char.IsWhiteSpace</a> method.</p>
' <p>If the <i>Count</i> parameter is zero, or the <i>Options</i> parameter is <a href="CorString.StringSplitOptions.html">RemoveEmptyEntries</a> and the length of the <i>s</i> parameter is zero, an empty array is returned.</p>
Public Function Split(ByRef S As String, Optional ByRef Separator As Variant, Optional ByVal Count As Long = &H7FFFFFFF, Optional ByVal Options As StringSplitOptions = StringSplitOptions.None) As String()
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    If Count = 0 Or (Len(S) = 0 And Options = StringSplitOptions.RemoveEmptyEntries) Then
        Split = Cor.NewStrings()
    Else
        Dim SeparatorIndexes()  As Long
        Dim SeparatorLengths()  As Long
        Dim MatchCount          As Long
        Dim MaxIndexes          As Long
        
        MaxIndexes = Len(S) \ 2 + 1
        ReDim SeparatorIndexes(0 To MaxIndexes)
        ReDim SeparatorLengths(0 To MaxIndexes)
        
        If IsMissing(Separator) Then
            MatchCount = FindWhitespaceSeparatorIndexes(S, SeparatorIndexes, SeparatorLengths)
        ElseIf VarType(Separator) = vbString Then
            If Len(Separator) = 0 Then
                MatchCount = FindWhitespaceSeparatorIndexes(S, SeparatorIndexes, SeparatorLengths)
            Else
                MatchCount = FindSeparatorIndexes(S, Separator, SeparatorIndexes, SeparatorLengths)
            End If
        ElseIf CorArray.IsNullOrEmpty(Separator) Then
            MatchCount = FindWhitespaceSeparatorIndexes(S, SeparatorIndexes, SeparatorLengths)
        Else
            MatchCount = FindSeparatorIndexes(S, Separator, SeparatorIndexes, SeparatorLengths)
        End If
        
        If MatchCount > 0 Then
            Select Case Options
                Case StringSplitOptions.None
                    Split = SplitWithEmpties(S, MatchCount, SeparatorIndexes, SeparatorLengths, Count)
                Case StringSplitOptions.RemoveEmptyEntries
                    Split = SplitWithoutEmpties(S, MatchCount, SeparatorIndexes, SeparatorLengths, Count)
                Case Else
                    Error.Argument ArgumentOutOfRange_Enum
            End Select
        Else
            Split = Cor.NewStrings(S)
        End If
    End If
End Function

Private Function FindSeparatorIndexes(ByRef S As String, ByRef Separator As Variant, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim StringSeparator()   As String
    Dim CharSeparator()     As Integer
    
    On Error GoTo Catch
    
    Select Case VarType(Separator)
        Case vbString
            FindSeparatorIndexes = FindStringSeparatorIndexes(S, Separator, SeparatorIndexes, SeparatorLengths)
        Case vbStringArray
            SAPtr(StringSeparator) = SAPtrV(Separator)
            
            If Len1D(Separator) = 1 Then
                FindSeparatorIndexes = FindStringSeparatorIndexes(S, StringSeparator(LBound(StringSeparator)), SeparatorIndexes, SeparatorLengths)
            Else
                FindSeparatorIndexes = FindStringSeparatorsIndexes(S, StringSeparator, SeparatorIndexes, SeparatorLengths)
            End If
        Case vbIntegerArray
            SAPtr(CharSeparator) = SAPtrV(Separator)
            FindSeparatorIndexes = FindCharSeparatorsIndexes(S, CharSeparator, SeparatorIndexes, SeparatorLengths)
        Case Else
            Error.Argument Argument_CharArrayRequired, "Separator"
    End Select

    Exit Function
    
Catch:
    SAPtr(StringSeparator) = vbNullPtr
    SAPtr(CharSeparator) = vbNullPtr
    ThrowOrErr Err
End Function

Private Function FindStringSeparatorIndexes(ByRef S As String, ByRef Separator As Variant, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim MatchCount As Long
    Dim SepLength  As Long
    Dim MatchIndex As Long
    
    SepLength = Len(Separator)
        
    Do
        MatchIndex = InStr(MatchIndex + 1, S, Separator) - 1
        
        If MatchIndex >= 0 Then
            SeparatorIndexes(MatchCount) = MatchIndex
            SeparatorLengths(MatchCount) = SepLength
            MatchCount = MatchCount + 1
            MatchIndex = MatchIndex + SepLength
        Else
            Exit Do
        End If
    Loop
    
    FindStringSeparatorIndexes = MatchCount
End Function

Private Function FindStringSeparatorsIndexes(ByRef S As String, ByRef Separator() As String, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim CharIndex   As Long
    Dim SLength     As Long
    Dim SeparatorUB As Long
    Dim SeparatorLB As Long
    Dim MatchCount  As Long
    Dim Sep         As String
    Dim SepLength   As Long
        
    SeparatorUB = UBound(Separator)
    SeparatorLB = LBound(Separator)
    SLength = Len(S)
    
    Do While CharIndex < SLength
        Dim I As Long
        
        For I = SeparatorLB To SeparatorUB
            Sep = Separator(I)
            SepLength = Len(Sep)
            
            If SepLength > SLength - CharIndex Then
                Exit For
            End If
            
            If CompareEx(S, CharIndex, Sep, 0, SepLength, StringComparison.Ordinal) = 0 Then
                SeparatorIndexes(MatchCount) = CharIndex
                SeparatorLengths(MatchCount) = SepLength
                MatchCount = MatchCount + 1
                CharIndex = CharIndex + SepLength - 1
                Exit For
            End If
        Next

        CharIndex = CharIndex + 1
    Loop
    
    FindStringSeparatorsIndexes = MatchCount
End Function

Private Function FindCharSeparatorsIndexes(ByRef S As String, ByRef Separator() As Integer, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim Chars()     As Integer
    Dim CharIndex   As Long
    Dim SLength     As Long
    Dim MatchCount  As Long
    Dim SeparatorUB As Long
    Dim SeparatorLB As Long
    
    SeparatorUB = UBound(Separator)
    SeparatorLB = LBound(Separator)
    SLength = Len(S)
    Chars = AllocChars(S)
    
    On Error GoTo Catch
    
    Do While CharIndex < SLength
        Dim I As Long
        
        For I = SeparatorLB To SeparatorUB
            If Chars(CharIndex) = Separator(I) Then
                SeparatorIndexes(MatchCount) = CharIndex
                SeparatorLengths(MatchCount) = 1
                MatchCount = MatchCount + 1
                Exit For
            End If
        Next
        
        CharIndex = CharIndex + 1
    Loop
    
    FreeChars Chars
    FindCharSeparatorsIndexes = MatchCount
    Exit Function
    
Catch:
    FreeChars Chars
    ThrowOrErr Err
End Function

Private Function FindWhitespaceSeparatorIndexes(ByRef S As String, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim Chars()     As Integer
    Dim CharIndex   As Long
    Dim SLength     As Long
    Dim MatchCount  As Long
    
    SLength = Len(S)
    Chars = AllocChars(S)
    
    On Error GoTo Catch
    
    Do While CharIndex < SLength
        If IsWhiteSpace(Chars(CharIndex)) Then
            SeparatorIndexes(MatchCount) = CharIndex
            SeparatorLengths(MatchCount) = 1
            MatchCount = MatchCount + 1
        End If
        
        CharIndex = CharIndex + 1
    Loop
    
    FreeChars Chars
    FindWhitespaceSeparatorIndexes = MatchCount
    Exit Function
    
Catch:
    FreeChars Chars
    ThrowOrErr Err
End Function

Private Function SplitWithEmpties(ByRef S As String, ByVal MatchCount As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long, ByVal Count As Long) As String()
    Dim Result()            As String
    Dim MaxSubstrings       As Long
    Dim SubstringStartIndex As Long
    
    MaxSubstrings = IIf(MatchCount + 1 < Count, MatchCount, Count - 1)
    ReDim Result(0 To MaxSubstrings)
    
    Dim I As Long
    For I = 0 To MaxSubstrings - 1
        Result(I) = Mid$(S, SubstringStartIndex + 1, SeparatorIndexes(I) - SubstringStartIndex)
        SubstringStartIndex = SeparatorIndexes(I) + SeparatorLengths(I)
    Next
    
    If SubstringStartIndex < Len(S) Then
        Result(MaxSubstrings) = Mid$(S, SubstringStartIndex + 1, Len(S) - SubstringStartIndex)
    End If

    SplitWithEmpties = Result
End Function

Private Function SplitWithoutEmpties(ByRef S As String, ByVal MatchCount As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long, ByVal Count As Long) As String()
    Dim MaxSubstrings As Long
    
    MaxSubstrings = CountSubstrings(S, MatchCount, SeparatorIndexes, SeparatorLengths)
    MaxSubstrings = IIf(MaxSubstrings < Count, MaxSubstrings, Count)
    SplitWithoutEmpties = CreateSubstrings(S, MatchCount, MaxSubstrings, SeparatorIndexes, SeparatorLengths)
End Function

Private Function CountSubstrings(ByRef S As String, ByVal MatchCount As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim SubstringStartIndex As Long
    
    Dim I As Long
    For I = 0 To MatchCount - 1
        If SeparatorIndexes(I) - SubstringStartIndex > 0 Then
            CountSubstrings = CountSubstrings + 1
        End If
        
        SubstringStartIndex = SeparatorIndexes(I) + SeparatorLengths(I)
    Next
    
    If SubstringStartIndex < Len(S) And (Len(S) - SubstringStartIndex) > 0 Then
        CountSubstrings = CountSubstrings + 1
    End If
End Function

Private Function CreateSubstrings(ByRef S As String, ByVal MatchCount As Long, ByVal MaxSubstrings As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As String()
    Dim Result()            As String
    Dim SubstringStartIndex As Long
    Dim ResultIndex         As Long
    
    ReDim Result(0 To MaxSubstrings - 1)
    
    Dim I As Long
    For I = 0 To MatchCount - 1
        If SeparatorIndexes(I) - SubstringStartIndex > 0 Then
            Result(ResultIndex) = Mid$(S, SubstringStartIndex + 1, SeparatorIndexes(I) - SubstringStartIndex)
            ResultIndex = ResultIndex + 1
        End If
        
        SubstringStartIndex = SeparatorIndexes(I) + SeparatorLengths(I)
    Next
    
    If SubstringStartIndex < Len(S) And (Len(S) - SubstringStartIndex) > 0 Then
        Result(MaxSubstrings - 1) = Mid$(S, SubstringStartIndex + 1, Len(S) - SubstringStartIndex)
    End If

    CreateSubstrings = Result
End Function

''
' Determines whether two specified Strings have the same value.
'
' @param a The first string to compare.
' @param b The second string to compare.
' @param ComparisonType One of the enumeration values that specifies the rules for the comparison.
' @return <b>True</b> if the value of the <i>a</i> parameter is equal to the value of the <i>b</i> parameter; otherwise, <b>False</b>.
Public Function Equals(ByRef a As String, ByRef b As String, Optional ByVal ComparisonType As StringComparison = StringComparison.Ordinal) As Boolean
    Select Case ComparisonType
        Case StringComparison.Ordinal
            If Len(a) = Len(b) Then
                Equals = EqualsOrdinal(StrPtr(a), StrPtr(b), Len(a))
            End If
        Case StringComparison.BinaryCompare
            Equals = StrComp(a, b, vbBinaryCompare) = 0
        Case StringComparison.TextCompare
            Equals = StrComp(a, b, vbTextCompare) = 0
        Case Else
            Equals = CompareHelper(StrPtr(a), Len(a), StrPtr(b), Len(b), ComparisonType) = 0
    End Select
End Function

' there are places where strings will be held in variants and we don't want to
' have to convert them to a string variable causing additional string allocations.
Friend Function EqualsV(ByRef a As Variant, ByRef b As Variant) As Boolean
    If Len(a) = Len(b) Then
        EqualsV = EqualsOrdinal(StrPtr(a), StrPtr(b), Len(a))
    End If
End Function

''
' Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.
'
' @param StrA The first string to compare.
' @param StrB The second string to compare.
' @param ComparisonType One of the enumeration values that specifies the rules to use in the comparison.
' @return A 32-bit signed integer that indicates the lexical relationship between the two comparands, as shown in the following table
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>StrA</i> is less than <i>StrB</>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>StrA</i> equals <i>StrB</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>StrA</i> is greater than <i>StrB</i>.</td>
'    </tr>
' </table>
Public Function Compare(ByRef StrA As String, ByRef StrB As String, Optional ByVal ComparisonType As StringComparison = StringComparison.CurrentCulture) As Long
    Select Case ComparisonType
        Case StringComparison.BinaryCompare
            Compare = StrComp(StrA, StrB, vbBinaryCompare)
        Case StringComparison.TextCompare
            Compare = StrComp(StrA, StrB, vbTextCompare)
        Case Else
            Compare = CompareHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), ComparisonType)
    End Select
End Function

' there are places where strings will be held in variants and we don't want to
' have to convert them to a string variable causing additional string allocations.
Friend Function CompareV(ByRef StrA As Variant, ByRef StrB As Variant, Optional ByVal ComparisonType As StringComparison = StringComparison.CurrentCulture) As Long
    Select Case ComparisonType
        Case StringComparison.BinaryCompare
            CompareV = StrComp(StrA, StrB, vbBinaryCompare)
        Case StringComparison.TextCompare
            CompareV = StrComp(StrA, StrB, vbTextCompare)
        Case Else
            CompareV = CompareHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), ComparisonType)
    End Select
End Function

''
' Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.
'
' @param StrA The first string to use in the comparison.
' @param IndexA The position of the substring within <i>StrA</i>.
' @param StrB The second string to use in the comparison.
' @param IndexB The position of the substring within <i>StrB</i>.
' @param Length The maximum number of characters in the substrings to compare.
' @param ComparisonType One of the enumeration values that specifies the rules to use in the comparison.
' @return A 32-bit signed integer that indicates the lexical relationship between <i>StrA</i> and <i>StrB</i>, as shown in the following table
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td>The substring in the <i>StrA</i> parameter is less than the substring in the <i>StrB</i> parameter.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td>The substrings are equal, or the <i>Length</i> parameter is zero.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td>The substring in <i>StrA</i> is greater than the substring in <i>StrB</i>.</td>
'    </tr>
' </table>
' @remarks The substrings to compare start in <i>StrA</i> at <i>IndexA</i> and in <i>StrB</i> at <i>IndexB</i>. Both <i>IndexA</i> and <i>IndexB</i> are zero-based;
' that is, the first character in <i>StrA</i> and <i>StrB</i> is at position zero, not position one.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>IndexA</i> is greater than <i>StrA</i> length.<br/>
'           -or-<br/>
'           <i>IndexB</i> is greater than <i>StrB</i> length.<br/>
'           -or-<br/>
'           <i>IndexA</i>, <i>IndexB</i>, or <i>Length</i> is negative.<br/>
'           -or-<br/>
'           Either <i>StrA</i> or <i>StrB</i> is empty, and <i>Length</i> is greater than zero.
'        </td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>ComparisonType</i> is not a <a href="CorString.StringComparison.html">StringComparison</a> value.</td>
'    </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td><a href="CorString.StringComparison.html">StringComparison</a> is not supported.</td>
'    </tr>
' </table>
Public Function CompareEx(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByVal Length As Long, Optional ByVal ComparisonType As StringComparison = StringComparison.CurrentCulture) As Long
    Dim PtrA As Long
    Dim PtrB As Long
    Dim LengthA As Long
    Dim LengthB As Long
    
    If ComparisonType = StringComparison.TextCompare Or _
        ComparisonType = StringComparison.BinaryCompare Then _
        Error.NotSupported NotSupported_StringComparison

    ValidateAndGetLengths StrA, IndexA, StrB, IndexB, Length, LengthA, LengthB
    GetOffsetPointers StrA, IndexA, StrB, IndexB, PtrA, PtrB
    CompareEx = CompareHelper(PtrA, LengthA, PtrB, LengthB, ComparisonType)
End Function

''
' Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.
'
' @param StrA The first string to compare.
' @param StrB The second string to compare.
' @param Culture The culture that supplies culture-specific comparison information.
' @param Options Options to use when performing the comparison (such as ignoring case or symbols).
' @return A 32-bit signed integer that indicates the lexical relationship between <i>StrA</i> and <i>StrB</i>, as shown in the following table
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>StrA</i> is less than <i>StrB</>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>StrA</i> equals <i>StrB</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>StrA</i> is greater than <i>StrB</i>.</td>
'    </tr>
' </table>
Public Function CompareCultural(ByRef StrA As String, ByRef StrB As String, ByVal Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    CompareCultural = CompareCulturalHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), Culture, Options)
End Function

' there are places where strings will be held in variants and we don't want to
' have to convert them to a string variable causing additional string allocations.
Friend Function CompareCulturalV(ByRef StrA As Variant, ByRef StrB As Variant, ByRef Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    CompareCulturalV = CompareCulturalHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), Culture, Options)
End Function

''
' Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.
'
' @param StrA The first string to use in the comparison.
' @param IndexA The position of the substring within <i>StrA</i>.
' @param StrB The second string to use in the comparison.
' @param IndexB The position of the substring within <i>StrB</i>.
' @param Length The maximum number of characters in the substrings to compare.
' @param Culture The culture that supplies culture-specific comparison information.
' @param Options Options to use when performing the comparison (such as ignoring case or symbols).
' @return A 32-bit signed integer that indicates the lexical relationship between <i>StrA</i> and <i>StrB</i>, as shown in the following table
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>StrA</i> is less than <i>StrB</>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>StrA</i> equals <i>StrB</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>StrA</i> is greater than <i>StrB</i>.</td>
'    </tr>
' </table>
Public Function CompareCulturalEx(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByVal Length As Long, ByVal Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    Dim PtrA As Long
    Dim PtrB As Long
    Dim LengthA As Long
    Dim LengthB As Long
    
    ValidateAndGetLengths StrA, IndexA, StrB, IndexB, Length, LengthA, LengthB
    GetOffsetPointers StrA, IndexA, StrB, IndexB, PtrA, PtrB
    CompareCulturalEx = CompareCulturalHelper(PtrA, LengthA, PtrB, LengthB, Culture, Options)
End Function

''
' Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.
'
' @param StrA The first string to compare.
' @param StrB The second string to compare.
' @return A 32-bit signed integer that indicates the lexical relationship between <i>StrA</i> and <i>StrB</i>, as shown in the following table
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>StrA</i> is less than <i>StrB</>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>StrA</i> equals <i>StrB</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>StrA</i> is greater than <i>StrB</i>.</td>
'    </tr>
' </table>
Public Function CompareOrdinal(ByRef StrA As String, ByRef StrB As String) As Long
    CompareOrdinal = CompareStringOrdinal(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB))
End Function

Friend Function CompareOrdinalV(ByRef StrA As Variant, ByRef StrB As Variant) As Long
    CompareOrdinalV = CompareStringOrdinal(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB))
End Function

''
' Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.
'
' @param StrA The first string to use in the comparison.
' @param IndexA The position of the substring within <i>StrA</i>.
' @param StrB The second string to use in the comparison.
' @param IndexB The position of the substring within <i>StrB</i>.
' @param Length The maximum number of characters in the substrings to compare.
' @return A 32-bit signed integer that indicates the lexical relationship between <i>StrA</i> and <i>StrB</i>, as shown in the following table
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td>The substring in the <i>StrA</i> parameter is less than the substring in the <i>StrB</i> parameter.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td>The substrings are equal, or the <i>Length</i> parameter is zero.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td>The substring in <i>StrA</i> is greater than the substring in <i>StrB</i>.</td>
'    </tr>
' </table>
' @remarks The <i>IndexA</i>, <i>IndexB</i>, and <i>Length</i> parameters must be nonnegative.
' <p>The number of characters compared is the lesser of the length of <i>StrA</i> less <i>IndexA</i>, the length of <i>StrB</i> less <i>IndexB</i>, and <i>Length</i>.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>IndexA</i> is greater than <i>StrA</i> length.<br/>
'           -or-<br/>
'           <i>IndexB</i> is greater than <i>StrB</i> length.<br/>
'           -or-<br/>
'           <i>IndexA</i>, <i>IndexB</i>, or <i>Length</i> is negative.<br/>
'           -or-<br/>
'           Either <i>StrA</i> or <i>StrB</i> is empty, and <i>Length</i> is greater than zero.
'        </td>
'   </tr>
' </table>
Public Function CompareOrdinalEx(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByVal Length As Long) As Long
    CompareOrdinalEx = CompareEx(StrA, IndexA, StrB, IndexB, Length, StringComparison.Ordinal)
End Function

''
' Returns character value from a string at the specified index.
'
' @param s The string to return the character from.
' @param Index The position of the character in the string.
' @return A 16-bit integer representing the character.
' @remarks <i>Index</i> is zero-based.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>
'           <i>Index</i> is less than zero.<br/>
'           -or-<br/>
'           <i>Index</i> is equal to or greater than the length of <i>s</i>.
'       </td>
'   </tr>
' </table>
Public Function CharAt(ByRef S As String, ByVal Index As Long) As Integer
    If Index < 0 Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_NeedNonNegNum
    If Index >= Len(S) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    CharAt = MemWord(StrPtr(S) + Index * vbSizeOfChar)
End Function

Friend Function LastChar(ByRef S As String) As Integer
    Dim Length As Long
    Dim LastCharIndex As Long
    
    Length = Len(S)
    
    If Length > 0 Then
        LastCharIndex = Length - 1
        LastChar = MemWord(StrPtr(S) + LastCharIndex * vbSizeOfChar)
    End If
End Function

Friend Function AllChars(ByRef S As String, ByVal Predicate As Long) As Boolean
    AllChars = CheckChars(S, False, Predicate)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitCompareStructures
    LoadLowerUpperLookupTable
End Sub

'Private Sub Class_Terminate()
'    DisposeCompareStructures
'End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitCompareStructures()
    With mString1SA
        .cbElements = 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    
    mString2SA = mString1SA
    
    With mString1LngSA
        .cbElements = 4
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    
    mString2LngSA = mString1LngSA
End Sub

Private Sub HookupStringProxies(ByVal PtrA As Long, ByVal PtrB As Long)
    mString1SA.PVData = PtrA
    mString2SA.PVData = PtrB
    mString1LngSA.PVData = PtrA
    mString2LngSA.PVData = PtrB
    SAPtr(mString1) = VarPtr(mString1SA)
    SAPtr(mString2) = VarPtr(mString2SA)
    SAPtr(mString1Lng) = VarPtr(mString1LngSA)
    SAPtr(mString2Lng) = VarPtr(mString2LngSA)
End Sub

Private Sub UnhookStringProxies()
    SAPtr(mString1) = vbNullPtr
    SAPtr(mString2) = vbNullPtr
    SAPtr(mString1Lng) = vbNullPtr
    SAPtr(mString2Lng) = vbNullPtr
End Sub

Private Sub LoadLowerUpperLookupTable()
    Dim Data()  As Byte
    Dim I       As Long
    Dim LowCh   As Long
    
    Data = LoadResData(101, "LOWERUPPERLOOKUP")
    ReDim mLowerUpper(0 To 65535)
    
    ' Fill in table with one-to-one lookup
    ' for when we just fallback outside ASCII values.
    For I = 0 To 65535
        mLowerUpper(I) = AsWord(I)
    Next
    
    ' Replace one-to-one values with lower-to-upper
    ' values when looking up characters outside ASCII.
    ' Table structure:
    ' lower case: 16bits
    ' upper case: 16bits
    For I = 0 To UBound(Data) Step 4
        LowCh = AsWord(Data(I)) And &HFFFF&
        mLowerUpper(LowCh) = AsWord(Data(I + 2))
    Next
End Sub

Private Function MakeStringRange(ByRef S As String, ByRef StartIndex As Variant, ByRef Count As Variant) As ListRange
    If IsMissing(StartIndex) Then
        If Not IsMissing(Count) Then _
            Error.Argument Argument_ParamRequired, "StartIndex"
            
        MakeStringRange.Index = Len(S) - 1
        MakeStringRange.Count = Len(S)
    Else
        MakeStringRange.Index = StartIndex
        MakeStringRange.Count = CLngOrDefault(Count, MakeStringRange.Index + 1)
    End If
End Function

Private Function FormatCore(ByRef Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant) As String
    If SAPtr(Args) = vbNullPtr Then _
        Error.ArgumentNull "Args", ArgumentNull_Array
    
    Dim Output As StringBuilder
    Set Output = StringBuilderCache.Acquire(Len(Format) + Len1D(Args) * 8)
    Output.AppendFormatArrayEx Provider, Format, Args
    FormatCore = StringBuilderCache.GetStringAndRelease(Output)
End Function

Private Function IndexOfAnyCore(ByRef S As String, ByRef OfAny As Variant, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal Step As Long) As Long
    Dim OfAnyChars()    As Integer
    Dim Chars()         As Integer
    Dim OfAnyLB         As Long
    Dim OfAnyUB         As Long
        
    On Error GoTo Catch
    OfAnyChars = AsChars(OfAny)
    
    If SAPtr(OfAnyChars) = vbNullPtr Then _
        Error.ArgumentNull "OfAny", ArgumentNull_Array
        
    Chars = AllocChars(S)
    OfAnyLB = LBound(OfAnyChars)
    OfAnyUB = UBound(OfAnyChars)
    
    Dim I   As Long
    Dim j   As Long
    Dim Ch  As Integer
    For I = StartIndex To EndIndex Step Step
        Ch = Chars(I)
        
        For j = OfAnyLB To OfAnyUB
            If Ch = OfAnyChars(j) Then
                IndexOfAnyCore = I
                GoTo CleanExit
            End If
        Next
    Next
    
    IndexOfAnyCore = -1
    GoTo CleanExit
    
Catch:
    FreeChars OfAnyChars
    FreeChars Chars
    ThrowOrErr Err
    
CleanExit:
    FreeChars OfAnyChars
    FreeChars Chars
End Function

Private Function TrimCore(ByRef S As String, ByRef TrimChars As Variant, ByVal TrimHead As Boolean, ByVal TrimTail As Boolean) As String
    Select Case VarType(TrimChars)
        Case vbString, vbIntegerArray
            Dim Chars() As Integer
            Chars = AsChars(TrimChars)
            TrimCore = TrimGeneral(S, Chars, TrimHead, TrimTail)
            FreeChars Chars
            
        Case vbMissing
            TrimCore = TrimWhiteSpace(S, TrimHead, TrimTail)
            
        Case Else
            Error.Argument Argument_CharArrayRequired, "TrimChars"
    End Select
End Function

Private Function TrimGeneral(ByRef S As String, ByRef TrimChars() As Integer, ByVal TrimHead As Boolean, ByVal TrimTail As Boolean) As String
    Dim I           As Long
    Dim TrimCharsLB As Long
    Dim TrimCharsUB As Long
    Dim Ch          As Integer
    Dim Length      As Long
    Dim StartIndex  As Long
    Dim EndIndex    As Long
    Dim Chars()     As Integer
    
    Chars = AllocChars(S)
    TrimCharsLB = LBound(TrimChars)
    TrimCharsUB = UBound(TrimChars)
    Length = Len(S)
    
    If TrimHead Then
        Do While StartIndex < Length
            Ch = Chars(StartIndex)
            For I = TrimCharsLB To TrimCharsUB
                If Ch = TrimChars(I) Then Exit For
            Next I
            If I > TrimCharsUB Then Exit Do
            StartIndex = StartIndex + 1
        Loop
        
        If StartIndex = Length Then
            GoTo Finally
        End If
    End If
    
    EndIndex = Length - 1
    If TrimTail Then
        Do While EndIndex >= StartIndex
            Ch = Chars(EndIndex)
            For I = TrimCharsLB To TrimCharsUB
                If Ch = TrimChars(I) Then Exit For
            Next I
            If I > TrimCharsUB Then Exit Do
            EndIndex = EndIndex - 1
        Loop
        
        If EndIndex < 0 Then
            GoTo Finally
        End If
    End If
        
    TrimGeneral = Mid$(S, StartIndex + 1, EndIndex - StartIndex + 1)
    
Finally:
    FreeChars Chars
End Function

Private Function TrimWhiteSpace(ByRef S As String, ByVal TrimHead As Boolean, ByVal TrimTail As Boolean) As String
    Dim StartIndex  As Long
    Dim EndIndex    As Long
    Dim I           As Long
    Dim Chars()     As Integer
    
    Chars = AllocChars(S)
    If TrimHead Then
        For I = 0 To UBound(Chars)
            If Not IsWhiteSpace(Chars(I)) Then
                StartIndex = I
                Exit For
            End If
        Next
        
        If I > UBound(Chars) Then
            GoTo Finally
        End If
    End If
    
    EndIndex = Len(S) - 1
    If TrimTail Then
        For I = UBound(Chars) To StartIndex Step -1
            If Not IsWhiteSpace(Chars(I)) Then
                EndIndex = I
                Exit For
            End If
        Next
        
        If I < 0 Then
            GoTo Finally
        End If
    End If
        
    TrimWhiteSpace = Mid$(S, StartIndex + 1, EndIndex - StartIndex + 1)
    
Finally:
    FreeChars Chars
End Function

Private Function CharCount(ByRef S As String, ByVal Char As Integer) As Long
    If LenB(S) > 0 Then
        Dim Chars() As Integer
        Chars = AllocChars(S)

        Dim I As Long
        For I = 0 To UBound(Chars)
            If Chars(I) = Char Then
                CharCount = CharCount + 1
            End If
        Next

        FreeChars Chars
    End If
End Function

Private Function JoinStrings(ByRef Separator As String, ByRef Value As Variant, ByRef Range As ListRange) As String
    Dim SubStrings()    As String
    Dim ArrayStruct     As SafeArray1d
    
    With ArrayStruct
        .cbElements = vbSizeOfPtr
        .cDims = 1
        .cElements = Range.Count
        .fFeatures = &H100
        .PVData = MemLong(SAPtrV(Value) + PVDATA_OFFSET) + (Range.Index - LBound(Value)) * vbSizeOfPtr
    End With
    
    SAPtr(SubStrings) = VarPtr(ArrayStruct)
    JoinStrings = VBA.Join(SubStrings, Separator)
    SAPtr(SubStrings) = vbNullPtr
End Function

Private Function JoinVariants(ByRef Separator As String, ByRef Value As Variant, ByRef Range As ListRange) As String
    Dim SubStrings() As String
    Dim Variants() As Variant
    Dim I As Long
    
    ReDim SubStrings(0 To Range.Count - 1)
    SAPtr(Variants) = SAPtrV(Value)
    
    For I = 0 To Range.Count - 1
        SubStrings(I) = Object.ToString(Variants(Range.Index + I))
    Next
    
    SAPtr(Variants) = vbNullPtr
    JoinVariants = VBA.Join(SubStrings, Separator)
End Function

Private Function JoinArray(ByRef Separator As String, ByRef Value As Variant, ByRef Range As ListRange) As String
    Dim Strings() As String
    ReDim Strings(0 To Range.Count - 1)
    
    Dim I As Long
    For I = 0 To Range.Count - 1
        Strings(I) = Object.ToString(Value(Range.Index + I))
    Next
    
    JoinArray = VBA.Join(Strings, Separator)
End Function

Private Function JoinCollection(ByRef Separator As String, ByVal Value As Object) As String
    Dim AppendSeparator As Boolean
    Dim Item            As Variant
    Dim ResultBuilder   As StringBuilder
    
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    On Error GoTo NonEnumerable
    Set ResultBuilder = StringBuilderCache.Acquire
    
    For Each Item In Value
        If AppendSeparator Then
            ResultBuilder.AppendString Separator
        Else
            AppendSeparator = True
        End If
        
        ResultBuilder.AppendString Object.ToString(Item)
    Next
    
    JoinCollection = StringBuilderCache.GetStringAndRelease(ResultBuilder)
    Exit Function
    
NonEnumerable:
    StringBuilderCache.Release ResultBuilder
    Error.Argument Argument_NotEnumerable, "Value"
End Function

Private Function CheckChars(ByRef S As String, ByVal SuccessfulState As Boolean, ByVal Predicate As Long) As Boolean
    Dim Checker As Func_T_Boolean
    Dim Chars() As Integer
    Dim I       As Long

    Set Checker = InitDelegate(mCharCheckerDelegate, Predicate)
    Chars = AllocChars(S)
    CheckChars = Not SuccessfulState
    
    For I = 0 To UBound(Chars)
        If Checker.Invoke(Chars(I)) = SuccessfulState Then
            CheckChars = SuccessfulState
            Exit For
        End If
    Next
    
    FreeChars Chars
End Function

Private Sub ValidateAndGetLengths(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByVal Length As Long, ByRef LengthA As Long, ByRef LengthB As Long)
    LengthA = Len(StrA)
    LengthB = Len(StrB)

    If IndexA < 0 Or IndexB < 0 Then _
        Error.ArgumentOutOfRange IIf(IndexA < 0, "IndexA", "IndexB"), ArgumentOutOfRange_NeedNonNegNum
    If IndexA > LengthA Or IndexB > LengthB Then _
        Error.ArgumentOutOfRange IIf(IndexA > LengthA, "StrA", "StrB"), ArgumentOutOfRange_OffsetLength
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum

    If IndexA + Length > LengthA Then
        LengthA = LengthA - IndexA
    Else
        LengthA = Length
    End If

    If IndexB + Length > LengthB Then
        LengthB = LengthB - IndexB
    Else
        LengthB = Length
    End If
End Sub

Private Sub GetOffsetPointers(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByRef PtrA As Long, ByRef PtrB As Long)
    PtrA = StrPtr(StrA) + IndexA * vbSizeOfChar
    PtrB = StrPtr(StrB) + IndexB * vbSizeOfChar
End Sub

Private Function CompareHelper(ByVal PtrA As Long, ByVal LengthA As Long, ByVal PtrB As Long, ByVal LengthB As Long, ByVal ComparisonType As StringComparison) As Long
    ' CompareString returns a Windows defined value so we need to adjust for typical comparison results.
    ' CSTR_LESS_THAN = 1
    ' CSTR_EQUAL = 2
    ' CSTR_GREATER_THAN = 3
    Const Adjustment = -2
    
    Dim Result As Long
    
    ' Do a quick check first
    If LengthA = 0 And LengthB = 0 Then
        Exit Function
    End If
        
    Select Case ComparisonType
        Case StringComparison.Ordinal
            CompareHelper = CompareStringOrdinal(PtrA, LengthA, PtrB, LengthB)
            Exit Function
        Case StringComparison.OrdinalIgnoreCase
            CompareHelper = CompareStringOrdinalIgnoreCase(PtrA, LengthA, PtrB, LengthB)
            Exit Function
        Case StringComparison.CurrentCulture
            Result = CompareString(CultureInfo.CurrentCulture.LCID, 0, PtrA, LengthA, PtrB, LengthB)
        Case StringComparison.CurrentCultureIgnoreCase
            Result = CompareString(CultureInfo.CurrentCulture.LCID, NORM_IGNORECASE, PtrA, LengthA, PtrB, LengthB)
        Case StringComparison.InvariantCulture
            Result = CompareString(CultureInfo.InvariantCulture.LCID, 0, PtrA, LengthA, PtrB, LengthB)
        Case StringComparison.InvariantCultureIgnoreCase
            Result = CompareString(CultureInfo.InvariantCulture.LCID, NORM_IGNORECASE, PtrA, LengthA, PtrB, LengthB)
        Case Else
            Error.Argument NotSupported_StringComparison, "ComparisonType"
    End Select
        
    If Result = 0 Then _
        Error.Win32Error Err.LastDllError
        
    CompareHelper = Result + Adjustment
End Function

Private Function CompareCulturalHelper(ByVal PtrA As Long, ByVal LengthA As Long, ByVal PtrB As Long, ByVal LengthB As Long, ByRef Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    ' CompareString returns a Windows defined value so we need to adjust for typical comparison results.
    ' CSTR_LESS_THAN = 1
    ' CSTR_EQUAL = 2
    ' CSTR_GREATER_THAN = 3
    Const Adjustment = -2
    Dim Flags   As Long
    Dim Result  As Long
    
    If Culture Is Nothing Then _
        Error.ArgumentNull "Culture"
        
    If LengthA = 0 And LengthB = 0 Then
        Exit Function
    End If
    
    Select Case Options
        Case CompareOptions.OrdinalOption
            Result = CompareStringOrdinal(PtrA, LengthA, PtrB, LengthB)
        Case CompareOptions.OrdinalIgnoreCaseOption
            Result = CompareStringOrdinalIgnoreCase(PtrA, LengthA, PtrB, LengthB)
        Case Else
            Flags = GetCompareFlags(Options)
            Result = CompareString(Culture.LCID, Flags, PtrA, LengthA, PtrB, LengthB)
            
            If Result = 0 Then _
                Error.Win32Error Err.LastDllError
                
            Result = Result + Adjustment
    End Select

    CompareCulturalHelper = Result
End Function

Private Function GetCompareFlags(ByVal Options As CompareOptions) As Long
    Dim Flags As Long
    
    Debug.Assert (Options And CompareOptions.OrdinalOption) = 0
    Debug.Assert (Options And CompareOptions.OrdinalIgnoreCaseOption) = 0
    
    If Options And CompareOptions.IgnoreCase Then
        Flags = Flags Or NORM_IGNORECASE
    End If
    
    If Options And CompareOptions.IgnoreKanaType Then
        Flags = Flags Or NORM_IGNOREKANATYPE
    End If
    
    If Options And CompareOptions.IgnoreNonSpace Then
        Flags = Flags Or NORM_IGNORENONSPACE
    End If
    
    If Options And CompareOptions.IgnoreSymbols Then
        Flags = Flags Or NORM_IGNORESYMBOLS
    End If
    
    If Options And CompareOptions.IgnoreWidth Then
        Flags = Flags Or NORM_IGNOREWIDTH
    End If
    
    If Options And CompareOptions.StringSort Then
        Flags = Flags Or SORT_STRINGSORT
    End If
    
    GetCompareFlags = Flags
End Function

' This method is exposed to allow for quicker array sorting.
Friend Function CompareStringOrdinal(ByVal PtrA As Long, ByVal LengthA As Long, ByVal PtrB As Long, ByVal LengthB As Long) As Long
    Dim I       As Long
    Dim Length  As Long
    Dim Result  As Long
    
    HookupStringProxies PtrA, PtrB
'    mString1SA.PVData = PtrA
'    mString2SA.PVData = PtrB
    
    If LengthA <= LengthB Then
        Length = LengthA
    Else
        Length = LengthB
    End If
    
    ' We're guessing most strings will be 10 or less characters.
    ' Unroll 10 character comparisons without looping.
    Do While Length > 10
        Result = (mString1(I) And &HFFFF&) - (mString2(I) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 1) And &HFFFF&) - (mString2(I + 1) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 2) And &HFFFF&) - (mString2(I + 2) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 3) And &HFFFF&) - (mString2(I + 3) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 4) And &HFFFF&) - (mString2(I + 4) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 5) And &HFFFF&) - (mString2(I + 5) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 6) And &HFFFF&) - (mString2(I + 6) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 7) And &HFFFF&) - (mString2(I + 7) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 8) And &HFFFF&) - (mString2(I + 8) And &HFFFF&): If Result Then GoTo Finally
        Result = (mString1(I + 9) And &HFFFF&) - (mString2(I + 9) And &HFFFF&): If Result Then GoTo Finally
        
        I = I + 10
        Length = Length - 10
    Loop
    
    ' Compare the remaining characters one at a time.
    Do While Length > 0
        Result = (mString1(I) And &HFFFF&) - (mString2(I) And &HFFFF&)
        
        If Result Then
            GoTo Finally
        End If
        
        I = I + 1
        Length = Length - 1
    Loop
    
    Result = LengthA - LengthB
        
Finally:
    UnhookStringProxies
    CompareStringOrdinal = Result
End Function

' This method is exposed to allow for quicker array sorting with strings.
Friend Function CompareStringOrdinalIgnoreCase(ByVal PtrA As Long, ByVal LengthA As Long, ByVal PtrB As Long, ByVal LengthB As Long) As Long
    Dim ChA As Long
    Dim ChB As Long
    Dim Max As Long
    Dim I   As Long
    Dim Result As Long
    
    HookupStringProxies PtrA, PtrB
    
    If LengthA <= LengthB Then
        Max = LengthA - 1
    Else
        Max = LengthB - 1
    End If

    For I = 0 To Max
        ChA = mString1(I) And &HFFFF&
        ChB = mString2(I) And &HFFFF&
        
        ' We do quick range checks in the ASCII range
        ' to convert to upper case to avoid looking
        ' up in the table. Most characters will probably be ASCII.
        Select Case ChA
            Case Is < 97
            Case Is <= 122, 224 To 246, 248, 249 To 254
                ChA = ChA - 32
            Case Is > 254
                ChA = mLowerUpper(ChA) And &HFFFF&
        End Select
        
        Select Case ChB
            Case Is < 97
            Case Is <= 122, 224 To 246, 248, 249 To 254
                ChB = ChB - 32
            Case Is > 254
                ChB = mLowerUpper(ChB) And &HFFFF&
        End Select
        
        Result = ChA - ChB
        
        If Result Then
            GoTo Finally
        End If
    Next
    
    Result = LengthA - LengthB
    
Finally:
    UnhookStringProxies
    CompareStringOrdinalIgnoreCase = Result
End Function

' This function expects both strings to be the same length, so it
' is expected the calling method did a length check first.
Private Function EqualsOrdinal(ByVal PtrA As Long, ByVal PtrB As Long, ByVal Length As Long) As Boolean
    Dim ChA As Long
    Dim ChB As Long
    Dim Index As Long
    
    HookupStringProxies PtrA, PtrB
    
    ' We check 32 bits at a time since we are only looking for equality.
    ' The loop is unrolled compare 10 characters without looping. We are
    ' guessing that most strings won't be longer than that.
    Do While Length > 10
        If mString1Lng(Index) <> mString2Lng(Index) Then GoTo Finally
        If mString1Lng(Index + 1) <> mString2Lng(Index + 1) Then GoTo Finally
        If mString1Lng(Index + 2) <> mString2Lng(Index + 2) Then GoTo Finally
        If mString1Lng(Index + 3) <> mString2Lng(Index + 3) Then GoTo Finally
        If mString1Lng(Index + 4) <> mString2Lng(Index + 4) Then GoTo Finally
        Index = Index + 5
        Length = Length - 10
    Loop
    
    ' Finish the remaining characters 2 at a time. Even if we check
    ' passed the end of a string, we will safely access the null terminator instead.
    Do While Length > 0
        If mString1Lng(Index) <> mString2Lng(Index) Then
            Exit Do
        End If
    
        Index = Index + 1
        Length = Length - 2
    Loop
    
    ' We ran out of characters so the strings must be equal.
    EqualsOrdinal = Length <= 0
    
Finally:
    UnhookStringProxies
End Function

