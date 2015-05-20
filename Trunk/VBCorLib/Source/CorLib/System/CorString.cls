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

Private mStringUtils As New StringUtils


''
' Returns if the string is null.
'
' @param s The string to be tested for null.
' @return Indication if the string is null.
' @remarks If the string is null it will return True. An empty string
' will return false. A null string is equal to StrPtr(s) = 0.
'
Public Function IsNull(ByRef s As String) As Boolean
    IsNull = (StrPtr(s) = vbNullPtr)
End Function

''
' Returns if the string is empty or contains only white-space.
'
' @param s The string to be tested for white-space.
' @return If the string has a length of zero or contains only white-space then True is returned, othersize False.
'
Public Function IsNullOrWhiteSpace(ByRef s As String) As Boolean
    If LenB(s) = 0 Then
        IsNullOrWhiteSpace = True
    Else
        IsNullOrWhiteSpace = ContainsOnlyWhiteSpace(s)
    End If
End Function

''
' Returns a string from the characters in an integer array.
'
' @param Chars The array to create the string from.
' @param Index The starting index in the array.
' @param Count The number of characters to use.
' @return A string created from the chars in an integer array.
'
Public Function FromCharArray(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Validation.CheckValidSingleDimArray Chars, Parameter_Chars

    Dim Range As ListRange
    Range = OptionalRange(Index, Count, LBound(Chars), CorArray.LengthFirstDim(Chars))
    
    If Range.Index < LBound(Chars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count - 1 > UBound(Chars) Then _
        Error.InvalidOffsetLength
    
    If Range.Count = 0 Then
        FromCharArray = ""
    Else
        FromCharArray = SysAllocStringLen(VarPtr(Chars(Range.Index)), Range.Count)
    End If
End Function

''
' Formats a string by replacing each argument with the formatted equivalent.
'
' @param FormatText The string containing formatting information.
' @param args A list of arguments to use in replacing the formatting information.
' @return A formatted string.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\Includes\CorString.Format.txt"
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
'
Public Function FormatEx(ByVal Provider As IFormatProvider, ByRef FormatText As String, ParamArray Args() As Variant) As String
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
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
'
Public Function FormatArray(ByRef FormatText As String, ByRef Args() As Variant) As String
    FormatArray = FormatCore(Nothing, FormatText, Args)
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
'
Public Function FormatArrayEx(ByVal Provider As IFormatProvider, ByRef FormatText As String, ByRef Args() As Variant) As String
    FormatArrayEx = FormatCore(Provider, FormatText, Args)
End Function

''
' Creates and array of chars (Integers) from the specified string.
'
' @param s The string to create the chars from.
' @param startindex The start index in the string to begin converting to a char array. This is zero-based.
' @param length The number of characters to convert to chars.
' @return An array containing the converted characters from the string.
'
Public Function ToCharArray(ByRef s As String, Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As Integer()
    Dim Range As ListRange
    Range = OptionalRange(StartIndex, Length, 0, Len(s), Parameter_StartIndex, Parameter_Length)
    
    If Range.Index < 0 Then _
        Error.NegativeNumber "StartIndex"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Length"
    If Range.Index + Range.Count > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_IndexCount
    
    Dim Chars() As Integer
    If Range.Count > 0 Then
        ReDim Chars(0 To Range.Count - 1)
        CopyMemory Chars(0), ByVal StrPtr(s) + Range.Index * SizeOfChar, Range.Count * SizeOfChar
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
' @remarks The paddingChar parameter can accept either a number or chr$ value.
'
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
' @remarks The paddingChar parameter can accept either a number or chr$ value.
'
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
' @remarks Unlike VB's LTrim function, this can trim multiple character types,
' not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from the beginning.</p>
'
Public Function TrimStart(ByRef s As String, Optional ByRef TrimChars As Variant) As String
    TrimStart = TrimCore(s, TrimChars, True, False)
End Function

''
' Trims the end of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's RTrim function, this can trim multiple character types,
' not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from the end.</p>
'
Public Function TrimEnd(ByRef s As String, Optional ByRef TrimChars As Variant) As String
    TrimEnd = TrimCore(s, TrimChars, False, True)
End Function

''
' Trims both the beginning and end of the string.
'
' @param s The string to be trimmed.
' @param trimChars The characters to remove.
' @return The trimmed string.
' @remarks Unlike VB's Trim function, this can trim multiple character types,
' not just spaces.
' <p>The <i>trimChars</i> can be either a string of the characters or an integer array
' of characters that will be trimmed from both ends.</p>
'
Public Function Trim(ByRef s As String, Optional ByRef TrimChars As Variant) As String
    Trim = TrimCore(s, TrimChars, True, True)
End Function

''
' Joins an array or subarray of strings into a single string, separated by the <i>Delimiter</i>.
'
' @param Separator A string inserted between each element when joining into a string.
' @param Value An array of strings or variants to join together.
' @param StartIndex The starting element in <i>Arr</i> to begin joining from.
' @param Count The number of strings elements to join together.
' @return A string containing elements joined from the array separated by the delimiter.
'
Public Function Join(ByRef Separator As String, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As String
    If IsObject(Value) Then
        Join = JoinCollection(Separator, Value)
        Exit Function
    End If
    
    Validation.CheckValidSingleDimArray Value, Parameter_Value
    
    Dim Range As ListRange
    Range = OptionalRange(StartIndex, Count, LBound(Value), CorArray.LengthFirstDim(Value), Parameter_StartIndex)
    
    If Range.Index < LBound(Value) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count - 1 > UBound(Value) Then _
        Error.ArgumentOutOfRange "", ArgumentString.Argument_InvalidOffLen
    
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
' Removes a substring from the string, returning a new string with the gap filled
' with the remaining characters.
'
' @param s The string to remove characters from.
' @param StartIndex The index of the first character to begin removing. This index is zero-based.
' @param Count The number of characters to be removed.
' @return A new string with the substring removed.
'
Public Function Remove(ByRef s As String, ByVal StartIndex As Long, Optional ByRef Count As Variant) As String
    Dim CharCount       As Long
    Dim RemovalCount    As Long
    CharCount = Len(s)
    RemovalCount = OptionalLong(Count, CharCount - StartIndex)
    
    If StartIndex < 0 Then _
        Error.NegativeNumber "StartIndex"
    If RemovalCount < 0 Then _
        Error.NegativeNumber "Count"
    If StartIndex + RemovalCount > Len(s) Then _
        Error.ArgumentOutOfRange "", ArgumentString.Argument_InvalidOffLen
    
    Dim SPtr As Long
    SPtr = StrPtr(s)
    Remove = SysAllocStringLen(SPtr, CharCount - RemovalCount)
    CopyMemory ByVal StrPtr(Remove) + StartIndex * SizeOfChar, ByVal SPtr + (StartIndex + RemovalCount) * SizeOfChar, (CharCount - (StartIndex + RemovalCount)) * SizeOfChar
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
'
Public Function IndexOfAny(ByRef s As String, ByRef OfAny As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    If Len(s) = 0 Then
        IndexOfAny = -1
        Exit Function
    End If
    
    Dim Range As ListRange
    Range = OptionalStepRange(StartIndex, Count, 0, Len(s), Parameter_StartIndex)
    
    If Range.Index < 0 Then _
        Error.NegativeNumber "StartIndex"
    If Range.Index > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count > Len(s) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Count
    
    IndexOfAny = IndexOfAnyCore(s, OfAny, Range.Index, Range.Index + Range.Count - 1, 1)
End Function

''
' Returns the index of the first character found in the <i>OfAny<i> set of characters
' searching from the end of the string.
'
' @param s The string to be searched.
' @param OfAny The set of characters to search for (String or Integer() Array).
' @param StartIndex The starting index of <i>s</i> to begin searching.
' @param Count The number of characters in <i>s</i> to search.
' @return The index of a character from the set, or -1 if none are found.
' @remarks The indexing of <i>s</i> is zero-based.
'
Public Function LastIndexOfAny(ByRef s As String, ByRef OfAny As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    If Len(s) = 0 Then
        LastIndexOfAny = -1
        Exit Function
    End If

    Dim Range As ListRange
    Range = OptionalReverseStepRange(StartIndex, Count, Len(s) - 1, Len(s), Parameter_StartIndex)
    
    If Range.Index < 0 Then _
        Error.NegativeNumber "StartIndex"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index - Range.Count + 1 < 0 Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Count
    
    LastIndexOfAny = IndexOfAnyCore(s, OfAny, Range.Index, Range.Index - Range.Count + 1, -1)
End Function


''
' Inserts a string into an existing string, returning the result.
'
' @param s The string to insert to.
' @param StartIndex The starting position to insert the string.
' @param Value The string to be inserted.
' @return A string containing the new string created from the insertion.
' @remarks Index is zero-based.
'
Public Function Insert(ByRef s As String, ByVal StartIndex As Long, ByRef Value As String) As String
    If StartIndex < 0 Or StartIndex > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex"
    
    Dim ValueLength As Long
    ValueLength = Len(Value)
    
    If ValueLength = 0 Then
        Insert = s
        Exit Function
    End If
    
    Dim SourcePtr As Long
    Dim TargetPtr As Long
    
    SourcePtr = StrPtr(s)
    Insert = SysAllocStringLen(ByVal 0&, ValueLength + Len(s))
    TargetPtr = StrPtr(Insert)
    
    ValueLength = ValueLength * 2
    StartIndex = StartIndex * 2
    CopyMemory ByVal TargetPtr, ByVal SourcePtr, StartIndex
    CopyMemory ByVal TargetPtr + StartIndex, ByVal StrPtr(Value), ValueLength
    CopyMemory ByVal TargetPtr + StartIndex + ValueLength, ByVal SourcePtr + StartIndex, LenB(Insert) - (StartIndex + ValueLength)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friends
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function EqualsIgnoreCase(ByRef String1 As String, ByRef String2 As String) As Boolean
    EqualsIgnoreCase = mStringUtils.EqualStrings(String1, String2, vbTextCompare)
End Function

Friend Function Equals(ByRef String1 As String, ByRef String2 As String) As Boolean
    Equals = mStringUtils.EqualStrings(String1, String2, vbBinaryCompare)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FormatCore(ByRef Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant)
    Dim Output As StringBuilder
    Set Output = StringBuilderCache.Acquire
    Output.InternalAppendFormat Provider, Format, Args
    FormatCore = StringBuilderCache.GetStringAndRelease(Output)
End Function

Private Function IndexOfAnyCore(ByRef s As String, ByRef OfAny As Variant, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal Step As Long) As Long
    Dim OfAnyChars() As Integer
    OfAnyChars = AsChars(OfAny)
    If SAPtr(OfAnyChars) = vbNullPtr Then _
        Error.ArgumentNull "OfAny", ArgumentNull_Array
    
    Dim Chars() As Integer
    Chars = AllocChars(s)
        
    Dim i       As Long
    Dim OfAnyLB As Long
    Dim OfAnyUB As Long
    
    OfAnyLB = LBound(OfAnyChars)
    OfAnyUB = UBound(OfAnyChars)
    For i = StartIndex To EndIndex Step Step
        Dim Ch As Integer
        Ch = Chars(i)
        Dim j As Long
        For j = OfAnyLB To OfAnyUB
            If Ch = OfAnyChars(j) Then
                IndexOfAnyCore = i
                GoTo Finally
            End If
        Next
    Next
    
    IndexOfAnyCore = -1
    
Finally:
    FreeChars OfAnyChars
    FreeChars Chars
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
            If Not Char.IsWhiteSpaceChar(Chars(i)) Then
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
            If Not Char.IsWhiteSpaceChar(Chars(i)) Then
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

Private Function ContainsOnlyWhiteSpace(ByRef s As String) As Boolean
    ContainsOnlyWhiteSpace = True
    
    Dim Chars() As Integer
    Chars = AllocChars(s)
    
    Dim i As Long
    For i = 0 To UBound(Chars)
        If Not Char.IsWhiteSpaceChar(Chars(i)) Then
            ContainsOnlyWhiteSpace = False
            Exit For
        End If
    Next
        
    FreeChars Chars
End Function

Friend Function CharCount(ByRef s As String, ByVal Char As Integer) As Long
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
        .cbElements = SizeOfPtr
        .cDims = 1
        .cElements = Range.Count
        .fFeatures = &H100
        .pvData = MemLong(CorArray.ArrayPointer(Value) + PVDATA_OFFSET) + (Range.Index - LBound(Value)) * SizeOfPtr
    End With
    
    SAPtr(SubStrings) = VarPtr(ArrayStruct)
    JoinStrings = VBA.Join(SubStrings, Separator)
    SAPtr(SubStrings) = vbNullPtr
End Function

Private Function JoinVariants(ByRef Separator As String, ByRef Value As Variant, ByRef Range As ListRange) As String
    Dim SubStrings() As String
    ReDim SubStrings(0 To Range.Count - 1)
    
    Dim Variants() As Variant
    SAPtr(Variants) = CorArray.ArrayPointer(Value)
    
    Dim i As Long
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
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Dim AppendSeparator As Boolean
    Dim Item            As Variant
    Dim ResultBuilder   As StringBuilder
    Set ResultBuilder = StringBuilderCache.Acquire
    
    On Error GoTo NonEnumerable
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


