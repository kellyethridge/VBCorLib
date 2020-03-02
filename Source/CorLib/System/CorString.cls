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
' @remarks This class cannot be directly instantiated. The methods are accessed
' through the class name itself.
' <pre>
' s = CorString.Trim(s, "\")
' </pre>
'
Option Explicit

''
' Enum flag to indicate how strings should be compared.
'
' @remarks The values assigned to each enum is equal to the .NET version of the enum.
'
    
''
' String comparisons
'
' @param Ordinal Specifies a case-sensitive string comparison.
' @param OrdinalIgnoreCase Specifies a case-insensitive string comparison.
Public Enum StringComparison
    Ordinal = 4
    OrdinalIgnoreCase = 5
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

Private mComparisonString1()    As Long
Private mComparisonString2()    As Long
Private mComparisonHeader1      As SafeArray1d
Private mComparisonHeader2      As SafeArray1d
Private mCharCheckerDelegate    As Delegate
Private mCharChecker            As Func_T_Boolean


''
' Returns if the string is null.
'
' @param s The string to be tested for null.
' @return <b>True</b> if the string pointer is null, <b>False</b> otherwise.
' @include "..\DocIncludes\CorString.IsNull.html"
Public Function IsNull(ByRef s As String) As Boolean
    IsNull = (StrPtr(s) = vbNullPtr)
End Function

''
' Returns if the string is null, empty or contains only white-space characters.
'
' @param s The string to be tested for white-space.
' @return <b>True</b> if the string has a length of zero or contains only white-space, <b>False</b> otherwise.
' @include "..\DocIncludes\CorString.IsNullOrWhiteSpace.html"
Public Function IsNullOrWhiteSpace(ByRef s As String) As Boolean
    Dim Chars() As Integer
    Dim i As Long
    
    If LenB(s) = 0 Then
        IsNullOrWhiteSpace = True
    Else
        Chars = AllocChars(s)
        
        For i = 0 To Len(s) - 1
            If Not IsWhiteSpace(Chars(i)) Then
                GoTo Finally
            End If
        Next

        IsNullOrWhiteSpace = True
    End If
    
Finally:
    FreeChars Chars
End Function

''
' Replaces the format item in a specified string with the string representation of a corresponding value in ParamArray arguments.
'
' @param FormatText The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\DocIncludes\CorString.Format.html"
Public Function Format(ByRef FormatText As String, ParamArray Args() As Variant) As String
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    Format = FormatCore(Nothing, FormatText, SwappedArgs)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param provider A custom format provider that is used to format the argument instead of the default formatting.
' @param FormatText The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
Public Function FormatEx(ByVal Provider As IFormatProvider, ByRef FormatText As String, ParamArray Args() As Variant) As String
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(20)
    FormatEx = FormatCore(Provider, FormatText, SwappedArgs)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param FormatText The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
Public Function FormatArray(ByRef FormatText As String, ByRef Args() As Variant) As String
    FormatArray = FormatCore(Nothing, FormatText, Args)
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param Provider A custom format provider that is used to format the argument instead of the default formatting.
' @param FormatText The string containing formatting information.
' @param Args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
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
Public Function ToCharArray(ByRef s As String, Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As Integer()
    Dim Range   As ListRange
    Dim Chars() As Integer
    
    Range = MakeDefaultRange(StartIndex, 0, Length, Len(s), NameOfStartIndex, NameOfLength)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_IndexCount
    
    If Range.Count > 0 Then
        ReDim Chars(0 To Range.Count - 1)
        CopyMemory Chars(0), ByVal StrPtr(s) + Range.Index * vbSizeOfChar, Range.Count * vbSizeOfChar
    Else
        Chars = Cor.NewIntegers()
    End If
    
    ToCharArray = Chars
End Function

''
' Pads a string's left side.
'
' @param s The string to be padded.
' @param totalWidth The total length of the final string.
' @param paddingChar The character to pad the left of the string with.
' @return The padded string.
' @remarks <i>PaddingChar</i> can accept either a number or Chr$ value.
Public Function PadLeft(ByRef s As String, ByVal TotalWidth As Long, Optional ByRef PaddingChar As Variant = 32) As String
    Dim Length As Long
    Length = Len(s)
    If TotalWidth > Length Then
        PadLeft = String$(TotalWidth - Length, PaddingChar) & s
    Else
        PadLeft = s
    End If
End Function

''
' Pads a string's right side.
'
' @param s The string to be padded.
' @param totalWidth The total length of the final string.
' @param paddingChar The character to pad the right of the string with.
' @return The padded string.
' @remarks <i>PaddingChar</i> can accept either a number or Chr$ value.
Public Function PadRight(ByRef s As String, ByVal TotalWidth As Long, Optional ByRef PaddingChar As Variant = 32) As String
    Dim Length As Long
    Length = Len(s)
    If TotalWidth > Length Then
        PadRight = s & String$(TotalWidth - Length, PaddingChar)
    Else
        PadRight = s
    End If
End Function

''
' Trims the beginning of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's LTrim function, this can trim multiple character types, not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from the beginning.</p>
Public Function TrimStart(ByRef s As String, Optional ByRef TrimChars As Variant) As String
    TrimStart = TrimCore(s, TrimChars, True, False)
End Function

''
' Trims the end of the string.
'
' @param s The string to be trimmed.
' @param TrimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's RTrim function, this can trim multiple character types, not just spaces.
' <p>The <i>TrimChars</i> can be either a string of the characters or an <b>Integer</b> array
' of characters that will be trimmed from the end.</p>
Public Function TrimEnd(ByRef s As String, Optional ByRef TrimChars As Variant) As String
    TrimEnd = TrimCore(s, TrimChars, False, True)
End Function

''
' Trims both the beginning and end of the string.
'
' @param s The string to be trimmed.
' @param TrimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's Trim function, this can trim multiple character types, not just spaces.
' <p><i>TrimChars</i> can be either a string of the characters or an <i>Integer</i> array
' of characters that will be trimmed from both ends.</p>
Public Function Trim(ByRef s As String, Optional ByRef TrimChars As Variant) As String
    Trim = TrimCore(s, TrimChars, True, True)
End Function

''
' Joins an array or subarray of strings into a single string, separated by the <i>Delimiter</i>.
'
' @param Separator A string inserted between each element when joining into a string.
' @param Value An array of strings or variants to join together.
' @param StartIndex The starting element in <i>Arr</i> to begin joining from.
' @param Count The number of string elements to join together.
' @return A string containing elements joined from the array separated by the delimiter.
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
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Destination</i> is <b>null</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>SourceIndex</i> or <i>Count</i> is negative.<br/>
'           -or-<br/>
'           <i>DestinationIndex</i> is less than the lower-bound of <i>Destination</i>.<br/>
'           -or-<br/>
'           <i>Count</i> is greater than the length of the substring from <i>StartIndex</i> to the end of the string.<br/>
'           -or-<br/>
'           <i>Count</i> is greater than the length of the subarray from <i>DestinationIndex</i> to the upper-bound of the <i>Destination</i> array.
'        </td>
'    </tr>
' </table>
' @remarks This method copies <i>Count</i> characters from the <i>SourceIndex</i> position of <i>Source</i> to the <i>DestinationIndex</i> position of destination character array.
' This method does not resize the destination character array; it must have a sufficient number of elements to accommodate the copied characters or the method throws an <a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a>.
' <p><i>SourceIndex</i> is zero-based, and <i>DestinationIndex</i> is lower-bound based.</p>
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
' Removes a substring from the string, returning a new string with the gap filled with the remaining characters.
'
' @param s The string to remove characters from.
' @param StartIndex The index of the first character to begin removing. This index is zero-based.
' @param Count The number of characters to be removed.
' @return A new string with the substring removed.
Public Function Remove(ByRef s As String, ByVal StartIndex As Long, Optional ByRef Count As Variant) As String
    Dim CharCount       As Long
    Dim RemovalCount    As Long
    Dim SPtr As Long
    
    CharCount = Len(s)
    RemovalCount = CLngOrDefault(Count, CharCount - StartIndex)
    
    If StartIndex < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If RemovalCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If StartIndex + RemovalCount > Len(s) Then _
        Error.ArgumentOutOfRange "", Argument_InvalidOffLen
    
    SPtr = StrPtr(s)
    Remove = SysAllocStringLen(SPtr, CharCount - RemovalCount)
    CopyMemory ByVal StrPtr(Remove) + StartIndex * vbSizeOfChar, ByVal SPtr + (StartIndex + RemovalCount) * vbSizeOfChar, (CharCount - (StartIndex + RemovalCount)) * vbSizeOfChar
End Function

''
' Returns the index of the first character found in the <i>OfAny<i> set of characters.
'
' @param s The string to be searched.
' @param OfAny The set of characters to search for (String or Integer() Array).
' @param StartIndex The starting index of <i>s</i> to begin searching.
' @param Count The number of characters in <i>s</i> to search.
' @return The index of a character from the set, or -1 if none are found.
' @remarks The indexing of <i>s</i> is zero-based.
Public Function IndexOfAny(ByRef s As String, ByRef OfAny As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    Range = MakeDefaultStepRange(StartIndex, 0, Count, Len(s), NameOfStartIndex)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > Len(s) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Count
    
    If Len(s) > 0 Then
        IndexOfAny = IndexOfAnyCore(s, OfAny, Range.Index, Range.Index + Range.Count - 1, 1)
    Else
        IndexOfAny = -1
    End If
End Function

''
' Returns the index of the first character found in the <i>OfAny<i> set of characters searching from the end of the string.
'
' @param s The string to be searched.
' @param OfAny The set of characters to search for (String or Integer() Array).
' @param StartIndex The starting index of <i>s</i> to begin searching.
' @param Count The number of characters in <i>s</i> to search.
' @return The index of a character from the set, or -1 if none are found.
' @remarks The indexing of <i>s</i> is zero-based.
Public Function LastIndexOfAny(ByRef s As String, ByRef OfAny As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    
    Range = MakeStringRange(s, StartIndex, Count)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_NeedNonNegNum
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index - Range.Count + 1 < 0 Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Count
    
    If Len(s) > 0 Then
        LastIndexOfAny = IndexOfAnyCore(s, OfAny, Range.Index, Range.Index - Range.Count + 1, -1)
    Else
        LastIndexOfAny = -1
    End If
End Function

''
' Inserts a string into an existing string, returning the result.
'
' @param s The string to insert to.
' @param StartIndex The starting position to insert the string.
' @param Value The string to be inserted.
' @return A string containing the new string created from the insertion.
' @remarks Index is zero-based.
Public Function Insert(ByRef s As String, ByVal StartIndex As Long, ByRef Value As String) As String
    Dim ValueLength As Long
    Dim SourcePtr As Long
    Dim TargetPtr As Long
       
    If StartIndex < 0 Or StartIndex > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex"
    
    ValueLength = Len(Value)
    
    If ValueLength = 0 Then
        Insert = s
        Exit Function
    End If
    
    SourcePtr = StrPtr(s)
    Insert = SysAllocStringLen(ByVal 0&, ValueLength + Len(s))
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
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>ComparisonType</i> is not a <a href="CorString.StringComparison.html">StringComparison</a> value.</td>
'   </tr>
' </table>
Public Function StartsWith(ByRef s As String, ByRef Value As String, Optional ByVal ComparisonType As StringComparison = Ordinal) As Boolean
    Dim SourceLength    As Long
    Dim ValueLength     As Long
    Dim SubstringPtr    As Long
    
    SourceLength = Len(s)
    ValueLength = Len(Value)
    
    If ValueLength <= SourceLength Then
        Select Case ComparisonType
            Case StringComparison.Ordinal
                StartsWith = OrdinalEquals(StrPtr(s), ValueLength, StrPtr(Value), ValueLength)
            Case StringComparison.OrdinalIgnoreCase
                StartsWith = OrdinalEqualsIgnoreCase(StrPtr(s), ValueLength, StrPtr(Value), ValueLength)
            Case Else
                Error.Argument ArgumentOutOfRange_Enum
        End Select
    End If
End Function

''
' Determines whether the end of a string matches the specified string when compared using the specified comparison option.
'
' @param s The string to check.
' @param Value The string to compare to the substring at the end of <i>s</i>.
' @param ComparisonType One of the enumeration values that determines how <i>s</i> and <i>Value</i> are compared.
' @return <b>True</b> if the <i>Value</i> parameter matches the end of <i>s</i>; otherwise <b>False</b>.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>ComparisonType</i> is not a <a href="CorString.StringComparison.html">StringComparison</a> value.</td>
'   </tr>
' </table>
Public Function EndsWith(ByRef s As String, ByRef Value As String, Optional ByVal ComparisonType As StringComparison = Ordinal) As Boolean
    Dim SourceLength    As Long
    Dim ValueLength     As Long
    Dim SubstringPtr    As Long
    
    SourceLength = Len(s)
    ValueLength = Len(Value)
    
    If ValueLength <= SourceLength Then
        Select Case ComparisonType
            Case StringComparison.Ordinal
                EndsWith = OrdinalEquals(StrPtr(s) + (SourceLength - ValueLength) * vbSizeOfChar, ValueLength, StrPtr(Value), ValueLength)
            Case StringComparison.OrdinalIgnoreCase
                EndsWith = OrdinalEqualsIgnoreCase(StrPtr(s) + (SourceLength - ValueLength) * vbSizeOfChar, ValueLength, StrPtr(Value), ValueLength)
            Case Else
                Error.Argument ArgumentOutOfRange_Enum
        End Select
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
' <h4>Exceptions</h4>
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
Public Function Split(ByRef s As String, ByRef Separator As Variant, Optional ByVal Count As Long = &H7FFFFFFF, Optional ByVal Options As StringSplitOptions = StringSplitOptions.None) As String()
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    If Count = 0 Or (Len(s) = 0 And Options = StringSplitOptions.RemoveEmptyEntries) Then
        Split = Cor.NewStrings()
    Else
        Dim SeparatorIndexes()  As Long
        Dim SeparatorLengths()  As Long
        Dim MatchCount          As Long
    
        If CorArray.IsNullOrEmpty(Separator) Then
            MatchCount = FindWhitespaceSeparatorIndexes(s, SeparatorIndexes, SeparatorLengths)
        Else
            MatchCount = FindSeparatorIndexes(s, Separator, SeparatorIndexes, SeparatorLengths)
        End If
        
        If MatchCount > 0 Then
            Select Case Options
                Case StringSplitOptions.None
                    Split = SplitWithEmpties(s, MatchCount, SeparatorIndexes, SeparatorLengths, Count)
                Case StringSplitOptions.RemoveEmptyEntries
                    Split = SplitWithoutEmpties(s, MatchCount, SeparatorIndexes, SeparatorLengths, Count)
                Case Else
                    Error.Argument ArgumentOutOfRange_Enum
            End Select
        Else
            Split = Cor.NewStrings(s)
        End If
    End If
End Function

Private Function FindSeparatorIndexes(ByRef s As String, ByRef Separator As Variant, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim StringSeparator()   As String
    Dim CharSeparator()     As Integer
    
    On Error GoTo Catch
    
    Select Case VarType(Separator)
        Case vbStringArray
            SAPtr(StringSeparator) = SAPtrV(Separator)
            FindSeparatorIndexes = FindStringSeparatorIndexes(s, StringSeparator, SeparatorIndexes, SeparatorLengths)
        Case vbIntegerArray
            SAPtr(CharSeparator) = SAPtrV(Separator)
            FindSeparatorIndexes = FindCharSeparatorIndexes(s, CharSeparator, SeparatorIndexes, SeparatorLengths)
        Case Else
            Error.Argument Argument_CharArrayRequired, "Separator"
    End Select

    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    SAPtr(StringSeparator) = vbNullPtr
    SAPtr(CharSeparator) = vbNullPtr
    Return
End Function

Private Function FindStringSeparatorIndexes(ByRef s As String, ByRef Separator() As String, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim SPtr            As Long
    Dim CharIndex       As Long
    Dim SLength         As Long
    Dim SeparatorPtr    As Long
    Dim SeparatorLength As Long
    Dim SeparatorUB     As Long
    Dim SeparatorLB     As Long
    Dim MatchCount      As Long
    
    SeparatorUB = UBound(Separator)
    SeparatorLB = LBound(Separator)
    SPtr = StrPtr(s)
    SLength = Len(s)
    ReDim SeparatorIndexes(0 To SLength \ 2 + 1)
    ReDim SeparatorLengths(0 To SLength \ 2 + 1)
    
    Do While CharIndex < SLength
        Dim i As Long
        
        For i = SeparatorLB To SeparatorUB
            SeparatorPtr = StrPtr(Separator(i))
            SeparatorLength = Len(Separator(i))
            
            If SeparatorLength > SLength - CharIndex Then
                Exit For
            End If
            
            If OrdinalEquals(SPtr + CharIndex * vbSizeOfChar, SeparatorLength, SeparatorPtr, SeparatorLength) Then
                SeparatorIndexes(MatchCount) = CharIndex
                SeparatorLengths(MatchCount) = SeparatorLength
                MatchCount = MatchCount + 1
                CharIndex = CharIndex + SeparatorLength - 1
                Exit For
            End If
        Next
        
        CharIndex = CharIndex + 1
    Loop
    
    FindStringSeparatorIndexes = MatchCount
End Function

Private Function FindCharSeparatorIndexes(ByRef s As String, ByRef Separator() As Integer, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim Chars()     As Integer
    Dim CharIndex   As Long
    Dim SLength     As Long
    Dim MatchCount  As Long
    Dim SeparatorUB As Long
    Dim SeparatorLB As Long
    
    SeparatorUB = UBound(Separator)
    SeparatorLB = LBound(Separator)
    SLength = Len(s)
    ReDim SeparatorIndexes(0 To SLength \ 2 + 1)
    ReDim SeparatorLengths(0 To SLength \ 2 + 1)
    Chars = AllocChars(s)
    
    On Error GoTo Catch
    
    Do While CharIndex < SLength
        Dim i As Long
        
        For i = SeparatorLB To SeparatorUB
            If Chars(CharIndex) = Separator(i) Then
                SeparatorIndexes(MatchCount) = CharIndex
                SeparatorLengths(MatchCount) = 1
                MatchCount = MatchCount + 1
                Exit For
            End If
        Next
        
        CharIndex = CharIndex + 1
    Loop
    
    GoSub Finally
    FindCharSeparatorIndexes = MatchCount
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars Chars
    Return
End Function

Private Function FindWhitespaceSeparatorIndexes(ByRef s As String, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim Chars()     As Integer
    Dim CharIndex   As Long
    Dim SLength     As Long
    Dim MatchCount  As Long
    
    SLength = Len(s)
    ReDim SeparatorIndexes(0 To SLength \ 2 + 1)
    ReDim SeparatorLengths(0 To SLength \ 2 + 1)
    Chars = AllocChars(s)
    
    On Error GoTo Catch
    
    Do While CharIndex < SLength
        If IsWhiteSpace(Chars(CharIndex)) Then
            SeparatorIndexes(MatchCount) = CharIndex
            SeparatorLengths(MatchCount) = 1
            MatchCount = MatchCount + 1
        End If
        
        CharIndex = CharIndex + 1
    Loop
    
    GoSub Finally
    FindWhitespaceSeparatorIndexes = MatchCount
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars Chars
    Return
End Function

Private Function SplitWithEmpties(ByRef s As String, ByVal MatchCount As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long, ByVal Count As Long) As String()
    Dim Result()            As String
    Dim MaxSubstrings       As Long
    Dim SubstringStartIndex As Long
    
    MaxSubstrings = IIf(MatchCount + 1 < Count, MatchCount, Count - 1)
    ReDim Result(0 To MaxSubstrings)
    
    Dim i As Long
    For i = 0 To MaxSubstrings - 1
        Result(i) = Mid$(s, SubstringStartIndex + 1, SeparatorIndexes(i) - SubstringStartIndex)
        SubstringStartIndex = SeparatorIndexes(i) + SeparatorLengths(i)
    Next
    
    If SubstringStartIndex < Len(s) Then
        Result(MaxSubstrings) = Mid$(s, SubstringStartIndex + 1, Len(s) - SubstringStartIndex)
    End If

    SplitWithEmpties = Result
End Function

Private Function SplitWithoutEmpties(ByRef s As String, ByVal MatchCount As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long, ByVal Count As Long) As String()
    Dim MaxSubstrings As Long
    
    MaxSubstrings = CountSubstrings(s, MatchCount, SeparatorIndexes, SeparatorLengths)
    MaxSubstrings = IIf(MaxSubstrings < Count, MaxSubstrings, Count)
    SplitWithoutEmpties = CreateSubstrings(s, MaxSubstrings, SeparatorIndexes, SeparatorLengths)
End Function

Private Function CountSubstrings(ByRef s As String, ByVal MatchCount As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As Long
    Dim SubstringStartIndex As Long
    
    Dim i As Long
    For i = 0 To MatchCount - 1
        If SeparatorIndexes(i) - SubstringStartIndex > 0 Then
            CountSubstrings = CountSubstrings + 1
        End If
        
        SubstringStartIndex = SeparatorIndexes(i) + SeparatorLengths(i)
    Next
    
    If SubstringStartIndex < Len(s) And (Len(s) - SubstringStartIndex) > 0 Then
        CountSubstrings = CountSubstrings + 1
    End If
End Function

Private Function CreateSubstrings(ByRef s As String, ByVal MaxSubstrings As Long, ByRef SeparatorIndexes() As Long, ByRef SeparatorLengths() As Long) As String()
    Dim Result()            As String
    Dim SubstringStartIndex As Long
    Dim ResultIndex         As Long
    
    ReDim Result(0 To MaxSubstrings - 1)
    
    Dim i As Long
    For i = 0 To MaxSubstrings - 1
        If SeparatorIndexes(i) - SubstringStartIndex > 0 Then
            Result(ResultIndex) = Mid$(s, SubstringStartIndex + 1, SeparatorIndexes(i) - SubstringStartIndex)
            ResultIndex = ResultIndex + 1
        End If
        
        SubstringStartIndex = SeparatorIndexes(i) + SeparatorLengths(i)
    Next
    
    If SubstringStartIndex < Len(s) And (Len(s) - SubstringStartIndex) > 0 Then
        Result(MaxSubstrings - 1) = Mid$(s, SubstringStartIndex + 1, Len(s) - SubstringStartIndex)
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
Public Function Equals(ByRef a As String, ByRef b As String, Optional ByVal ComparisonType As StringComparison = Ordinal) As Boolean
    Select Case ComparisonType
        Case StringComparison.Ordinal
            Equals = OrdinalEquals(StrPtr(a), Len(a), StrPtr(b), Len(b))
        Case StringComparison.OrdinalIgnoreCase
            Equals = OrdinalEqualsIgnoreCase(StrPtr(a), Len(a), StrPtr(b), Len(b))
        Case Else
            Error.ArgumentOutOfRange "ComparisonType", ArgumentOutOfRange_Enum
    End Select
End Function

' this is a convenience function for when strings are being passed around in Variants.
Friend Function EqualsV(ByRef a As Variant, ByRef b As Variant, Optional ByVal ComparisonType As StringComparison = Ordinal) As Boolean
    Select Case ComparisonType
        Case StringComparison.Ordinal
            EqualsV = OrdinalEquals(StrPtr(a), Len(a), StrPtr(b), Len(b))
        Case StringComparison.OrdinalIgnoreCase
            EqualsV = OrdinalEqualsIgnoreCase(StrPtr(a), Len(a), StrPtr(b), Len(b))
        Case Else
            Error.ArgumentOutOfRange "ComparisonType", ArgumentOutOfRange_Enum
    End Select
End Function

Friend Function LastChar(ByRef s As String) As Integer
    Dim Length As Long
    Dim LastCharIndex As Long
    
    Length = Len(s)
    
    If Length > 0 Then
        LastCharIndex = Length - 1
        LastChar = MemWord(StrPtr(s) + LastCharIndex * vbSizeOfChar)
    End If
End Function

Friend Function CharAt(ByRef s As String, ByVal Index As Long) As Integer
    Debug.Assert Index >= 0 And Index < Len(s)
    
    CharAt = MemWord(StrPtr(s) + Index * vbSizeOfChar)
End Function

Friend Function AllChars(ByRef s As String, ByVal Predicate As Long) As Boolean
    AllChars = CheckChars(s, False, Predicate)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitComparisonStrings
    Set mCharChecker = InitDelegate(mCharCheckerDelegate)
End Sub

Private Sub Class_Terminate()
    ReleaseComparisonStrings
    
    ' Normally we could let light-weight delegate object release
    ' on their own, but, this class will be kept alive for the
    ' duration of the application, so when the delegation object
    ' is finally released, the Delegation module may have already
    ' be deallocated, removing the Release function the delegate
    ' object would have called. Since the delegation object is
    ' owned by this class, no memory needs deallocating, so we
    ' can just unhook the object pointer instead and let the
    ' normal class tear-down take care of the allocated delegate.
    ObjectPtr(mCharChecker) = vbNullPtr
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MakeStringRange(ByRef s As String, ByRef StartIndex As Variant, ByRef Count As Variant) As ListRange
    If IsMissing(StartIndex) Then
        If Not IsMissing(Count) Then _
            Error.Argument Argument_ParamRequired, "StartIndex"
            
        MakeStringRange.Index = Len(s) - 1
        MakeStringRange.Count = Len(s)
    Else
        MakeStringRange.Index = StartIndex
        MakeStringRange.Count = CLngOrDefault(Count, MakeStringRange.Index + 1)
    End If
End Function

Private Sub InitComparisonStrings()
    With mComparisonHeader1
        .cbElements = vbSizeOfChar * 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    With mComparisonHeader2
        .cbElements = vbSizeOfChar * 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    
    SAPtr(mComparisonString1) = VarPtr(mComparisonHeader1)
    SAPtr(mComparisonString2) = VarPtr(mComparisonHeader2)
End Sub

Private Sub ReleaseComparisonStrings()
    SAPtr(mComparisonString1) = vbNullPtr
    SAPtr(mComparisonString2) = vbNullPtr
End Sub

Private Function FormatCore(ByRef Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant) As String
    If SAPtr(Args) = vbNullPtr Then _
        Error.ArgumentNull "Args", ArgumentNull_Array
    
    Dim Output As StringBuilder
    Set Output = StringBuilderCache.Acquire
    Output.AppendFormatArrayEx Provider, Format, Args
    FormatCore = StringBuilderCache.GetStringAndRelease(Output)
End Function

Private Function IndexOfAnyCore(ByRef s As String, ByRef OfAny As Variant, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal Step As Long) As Long
    Dim OfAnyChars()    As Integer
    Dim Chars()         As Integer
    Dim OfAnyLB         As Long
    Dim OfAnyUB         As Long
        
    On Error GoTo Catch
    OfAnyChars = AsChars(OfAny)
    
    If SAPtr(OfAnyChars) = vbNullPtr Then _
        Error.ArgumentNull "OfAny", ArgumentNull_Array
        
    Chars = AllocChars(s)
    OfAnyLB = LBound(OfAnyChars)
    OfAnyUB = UBound(OfAnyChars)
    
    Dim i   As Long
    Dim j   As Long
    Dim Ch  As Integer
    For i = StartIndex To EndIndex Step Step
        Ch = Chars(i)
        
        For j = OfAnyLB To OfAnyUB
            If Ch = OfAnyChars(j) Then
                IndexOfAnyCore = i
                GoSub Finally
                Exit Function
            End If
        Next
    Next
    
    IndexOfAnyCore = -1
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    FreeChars OfAnyChars
    FreeChars Chars
    Return
End Function

Private Function TrimCore(ByRef s As String, ByRef TrimChars As Variant, ByVal TrimHead As Boolean, ByVal TrimTail As Boolean) As String
    Select Case VarType(TrimChars)
        Case vbString, vbIntegerArray
            Dim Chars() As Integer
            Chars = AsChars(TrimChars)
            TrimCore = TrimGeneral(s, Chars, TrimHead, TrimTail)
            FreeChars Chars
            
        Case vbMissing
            TrimCore = TrimWhiteSpace(s, TrimHead, TrimTail)
            
        Case Else
            Error.Argument Argument_CharArrayRequired, "TrimChars"
    End Select
End Function

Private Function TrimGeneral(ByRef s As String, ByRef TrimChars() As Integer, ByVal TrimHead As Boolean, ByVal TrimTail As Boolean) As String
    Dim i           As Long
    Dim TrimCharsLB As Long
    Dim TrimCharsUB As Long
    Dim Ch          As Integer
    Dim Length      As Long
    Dim StartIndex  As Long
    Dim EndIndex    As Long
    Dim Chars()     As Integer
    
    Chars = AllocChars(s)
    TrimCharsLB = LBound(TrimChars)
    TrimCharsUB = UBound(TrimChars)
    Length = Len(s)
    
    If TrimHead Then
        Do While StartIndex < Length
            Ch = Chars(StartIndex)
            For i = TrimCharsLB To TrimCharsUB
                If Ch = TrimChars(i) Then Exit For
            Next i
            If i > TrimCharsUB Then Exit Do
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
            For i = TrimCharsLB To TrimCharsUB
                If Ch = TrimChars(i) Then Exit For
            Next i
            If i > TrimCharsUB Then Exit Do
            EndIndex = EndIndex - 1
        Loop
        
        If EndIndex < 0 Then
            GoTo Finally
        End If
    End If
        
    TrimGeneral = Mid$(s, StartIndex + 1, EndIndex - StartIndex + 1)
    
Finally:
    FreeChars Chars
End Function

Private Function TrimWhiteSpace(ByRef s As String, ByVal TrimHead As Boolean, ByVal TrimTail As Boolean) As String
    Dim StartIndex  As Long
    Dim EndIndex    As Long
    Dim i           As Long
    Dim Chars()     As Integer
    
    Chars = AllocChars(s)
    If TrimHead Then
        For i = 0 To UBound(Chars)
            If Not IsWhiteSpace(Chars(i)) Then
                StartIndex = i
                Exit For
            End If
        Next
        
        If i > UBound(Chars) Then
            GoTo Finally
        End If
    End If
    
    EndIndex = Len(s) - 1
    If TrimTail Then
        For i = UBound(Chars) To StartIndex Step -1
            If Not IsWhiteSpace(Chars(i)) Then
                EndIndex = i
                Exit For
            End If
        Next
        
        If i < 0 Then
            GoTo Finally
        End If
    End If
        
    TrimWhiteSpace = Mid$(s, StartIndex + 1, EndIndex - StartIndex + 1)
    
Finally:
    FreeChars Chars
End Function

Private Function CharCount(ByRef s As String, ByVal Char As Integer) As Long
    If LenB(s) > 0 Then
        Dim Chars() As Integer
        Chars = AllocChars(s)

        Dim i As Long
        For i = 0 To UBound(Chars)
            If Chars(i) = Char Then
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
        .pvData = MemLong(SAPtrV(Value) + PVDATA_OFFSET) + (Range.Index - LBound(Value)) * vbSizeOfPtr
    End With
    
    SAPtr(SubStrings) = VarPtr(ArrayStruct)
    JoinStrings = VBA.Join(SubStrings, Separator)
    SAPtr(SubStrings) = vbNullPtr
End Function

Private Function JoinVariants(ByRef Separator As String, ByRef Value As Variant, ByRef Range As ListRange) As String
    Dim SubStrings() As String
    Dim Variants() As Variant
    Dim i As Long
    
    ReDim SubStrings(0 To Range.Count - 1)
    SAPtr(Variants) = SAPtrV(Value)
    
    For i = 0 To Range.Count - 1
        SubStrings(i) = Object.ToString(Variants(Range.Index + i))
    Next
    
    SAPtr(Variants) = vbNullPtr
    JoinVariants = VBA.Join(SubStrings, Separator)
End Function

Private Function JoinArray(ByRef Separator As String, ByRef Value As Variant, ByRef Range As ListRange) As String
    Dim Strings() As String
    ReDim Strings(0 To Range.Count - 1)
    
    Dim i As Long
    For i = 0 To Range.Count - 1
        Strings(i) = Object.ToString(Value(Range.Index + i))
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

' Adapted and modified from Chris Lucas XBeat/VBSpeed implementation
Private Function OrdinalEquals(ByVal Ptr1 As Long, ByVal Len1 As Long, ByVal Ptr2 As Long, ByVal Len2 As Long) As Boolean
    Dim i   As Long
    Dim Tmp As Long
    
    ' Make an informed decision as to whether we should continue
    If Len1 <> Len2 Then
        Exit Function
    End If
    
    mComparisonHeader1.pvData = Ptr1
    mComparisonHeader2.pvData = Ptr2
    
    Select Case Len1
        Case 0
            OrdinalEquals = True
            Exit Function
        Case 1
            OrdinalEquals = (mComparisonString1(0) And &HFFFF&) = (mComparisonString2(0) And &HFFFF&)
            Exit Function
    End Select
    
    Len1 = Len1 * vbSizeOfChar
    Len2 = Len2 * vbSizeOfChar
    Tmp = Len1 \ 4
    ' The first two characters come cheap
    If mComparisonString1(i) <> mComparisonString2(i) Then GoTo BailOut Else i = i + 1
   
DoLoop:
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If mComparisonString1(i) <> mComparisonString2(i) Then GoTo NotEqual Else i = i + 1
        If i <= Tmp Then GoTo DoLoop
    
NotEqual:
    ' some characters don't match, but we need to check to
    ' see if it happened after the end of the string, a
    ' nasty side-effect of cascading ifs
    If Tmp * 4 = Len1 Then
        If i >= Tmp Then OrdinalEquals = True
    Else
        If i > Tmp Then OrdinalEquals = True
    End If
    
BailOut:
    ' Lengths don't match, let's do absolutely nothing
End Function

' Adapted and modified from Chris Lucas XBeat/VBSpeed implementation
Private Function OrdinalEqualsIgnoreCase(ByVal Ptr1 As Long, ByVal SLen As Long, ByVal Ptr2 As Long, ByVal SLen2 As Long) As Boolean
    Dim i       As Long
    Dim Tmp1    As Long
    Dim Tmp2    As Long
    Dim Tmp3    As Long
    Dim Tmp4    As Long
    Dim Alt     As Long
    
    If SLen <> SLen2 Then
        Exit Function
    End If
    
    If SLen = 0 Then
        OrdinalEqualsIgnoreCase = True
        Exit Function
    End If
    
    SLen = SLen * 2
    SLen2 = SLen2 * 2
    mComparisonHeader1.pvData = Ptr1
    mComparisonHeader2.pvData = Ptr2

    For i = 0 To SLen \ 4 - 1
        Tmp1 = mComparisonString1(i)
        Tmp2 = (Tmp1 And &HFFFF&)
        Tmp3 = mComparisonString2(i)
        Tmp4 = (Tmp3 And &HFFFF&)
                
        Select Case Tmp2
            Case 97& To 122&: Alt = Tmp2 - 32
            Case 65& To 90&: Alt = Tmp2 + 32
            Case 49&: Alt = 185
            Case 50&: Alt = 178
            Case 51&: Alt = 179
            Case 138&: Alt = 154
            Case 140&: Alt = 156
            Case 142&: Alt = 158
            Case 154&: Alt = 138
            Case 156&: Alt = 140
            Case 158&: Alt = 142
            Case 159&: Alt = 255
            Case 178&: Alt = 50
            Case 179&: Alt = 51
            Case 185&: Alt = 49
            Case 192& To 214&: Alt = Tmp2 + 32
            Case 216& To 222&: Alt = Tmp2 + 32
            Case 224& To 246&: Alt = Tmp2 - 32
            Case 248& To 254&: Alt = Tmp2 - 32
            Case 255&: Alt = 376
            Case 338&: Alt = 339
            Case 339&: Alt = 338
            Case 352&: Alt = 353
            Case 353&: Alt = 352
            Case 376&: Alt = 255
            Case 381&: Alt = 382
            Case 382&: Alt = 381
        End Select
                
        If Alt <> Tmp4 Then
            If Tmp2 <> Tmp4 Then Exit Function
        End If
        Tmp2 = (Tmp1 And &HFFFF0000)
        Tmp4 = (Tmp3 And &HFFFF0000)
        Select Case Tmp2
            Case &H610000 To &H7A0000: Alt = Tmp2 - &H200000
            Case &H410000 To &H5A0000: Alt = Tmp2 + &H200000
            Case &H310000: Alt = &HB90000
            Case &H320000: Alt = &HB20000
            Case &H330000: Alt = &HB30000
            Case &H8A0000: Alt = &H9A0000
            Case &H8C0000: Alt = &H9C0000
            Case &H8E0000: Alt = &H9E0000
            Case &H9B0000: Alt = &H8A0000
            Case &H9C0000: Alt = &H8C0000
            Case &H9E0000: Alt = &H8E0000
            Case &H9F0000: Alt = &HFF0000
            Case &HB20000: Alt = &H320000
            Case &HB30000: Alt = &H970000
            Case &HB90000: Alt = &H310000
            Case &HC00000 To &HD60000: Alt = Tmp2 + &H200000
            Case &HD80000 To &HDE0000: Alt = Tmp2 + &H200000
            Case &HE00000 To &HF60000: Alt = Tmp2 - &H200000
            Case &HF80000 To &HFE0000: Alt = Tmp2 - &H200000
            Case &HFF0000: Alt = &H1780000
            Case &H1520000: Alt = &H1530000
            Case &H1530000: Alt = &H1520000
            Case &H1600000: Alt = &H1610000
            Case &H1610000: Alt = &H1600000
            Case &H1780000: Alt = &HFF0000
            Case &H17D0000: Alt = &H17E0000
            Case &H17E0000: Alt = &H17D0000
        End Select
        If Alt <> Tmp4 Then
            If Tmp2 <> Tmp4 Then Exit Function
        End If
    Next i
    
    If (SLen \ 2 And 1) Then
        Tmp2 = (mComparisonString1(i) And &HFFFF&)
        Tmp4 = (mComparisonString2(i) And &HFFFF&)
        
        Select Case Tmp2
            Case 97& To 122&: Alt = Tmp2 - 32
            Case 65& To 90&: Alt = Tmp2 + 32
            Case 49&: Alt = 185
            Case 50&: Alt = 178
            Case 51&: Alt = 179
            Case 138&: Alt = 154
            Case 140&: Alt = 156
            Case 142&: Alt = 158
            Case 154&: Alt = 138
            Case 156&: Alt = 140
            Case 158&: Alt = 142
            Case 159&: Alt = 255
            Case 178&: Alt = 50
            Case 179&: Alt = 51
            Case 185&: Alt = 49
            Case 192& To 214&: Alt = Tmp2 + 32
            Case 216& To 222&: Alt = Tmp2 + 32
            Case 224& To 246&: Alt = Tmp2 - 32
            Case 248& To 254&: Alt = Tmp2 - 32
            Case 255&: Alt = 376
            Case 338&: Alt = 339
            Case 339&: Alt = 338
            Case 352&: Alt = 353
            Case 353&: Alt = 352
            Case 376&: Alt = 255
            Case 381&: Alt = 382
            Case 382&: Alt = 381
        End Select
        
        If Tmp2 <> Tmp4 Then
            If Alt <> Tmp4 Then Exit Function
        End If
    End If

    OrdinalEqualsIgnoreCase = True
End Function

Private Function CheckChars(ByRef s As String, ByVal SuccessfulState As Boolean, ByVal Predicate As Long) As Boolean
    Dim Chars() As Integer
    Dim i       As Long

    mCharCheckerDelegate.pfn = Predicate
    Chars = AllocChars(s)
    CheckChars = Not SuccessfulState
    
    For i = 0 To UBound(Chars)
        If mCharChecker.Invoke(Chars(i)) = SuccessfulState Then
            CheckChars = SuccessfulState
            Exit For
        End If
    Next
    
    FreeChars Chars
End Function



