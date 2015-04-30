VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cString"
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
' Module: cString
'

''
' Provides static functions to format and manipulate string values.
'
' @remarks This class cannot be directly instantiated. The methods are accessed
' through the class name itself.
' <pre>
' s = cString.Trim(s, "\")
' </pre>
'
Option Explicit

Private mStringUtils    As New StringUtils
Private mString         As WordBuffer
Private mTrimChars      As WordBuffer
Private mJoinSA         As SafeArray1d
Private mJoinArray()    As String
Private mStringChars    As SafeArray1d
Private mVariantChars   As SafeArray1d


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
    If IsNullCharArray(Chars) Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
    If IsMultiDimCharArray(Chars) Then _
        Error.Rank

    Dim Range As ListRange
    Range = GetOptionalRange(Index, Count, LBound(Chars), ArrayLength(Chars))
    
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
' @include "..\Includes\cString.Format.txt"
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
    Range = GetOptionalRange(StartIndex, Length, 0, Len(s), Parameter_StartIndex, Parameter_Length)
    
    If Range.Index < 0 Then _
        Error.NegativeNumber "StartIndex"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Length"
    If Range.Index + Range.Count > Len(s) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_IndexCount
    
    Dim Chars() As Integer
    If Range.Count > 0 Then
        ReDim Chars(0 To Range.Count - 1)
        CopyMemory Chars(0), ByVal StrPtr(s) + Range.Index * 2, Range.Count * 2
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
' @param Arr An array of strings or variants to join together.
' @param Delimiter A string inserted between each element when joining into a string.
' @param Index The starting element in <i>Arr</i> to begin joining from.
' @param Count The number of strings elements to join together.
' @return A string containing elements joined from the array separated by the delimiter.
'
Public Function Join(ByRef Arr As Variant, Optional ByRef Delimiter As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim pSA         As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim lb          As Long
    Dim ub          As Long
    Dim ElemSize    As Long
    
    Select Case VarType(Arr)
        Case vbStringArray:      ElemSize = 4
        Case vbVariantArray:     ElemSize = 16
        Case Else
            Throw Cor.NewArgumentException("A String or Variant array is required.", "Arr")
    End Select
    
    pSA = GetArrayPointer(Arr, True)
    lb = LBound(Arr)
    ub = UBound(Arr)
    
    ElemIndex = GetOptionalLong(Index, lb)
    ElemCount = GetOptionalLong(Count, ub - ElemIndex + 1)
    
    Dim Result As Long
    Result = VerifyArrayRange(pSA, ElemIndex, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    With mJoinSA
        .cElements = ElemCount
        .fFeatures = MemWord(pSA + FFEATURES_OFFSET)
        .pvData = MemLong(pSA + PVDATA_OFFSET) + (ElemIndex - lb) * ElemSize
    End With
    Join = VBA.Join(mJoinArray, Delimiter)
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
Public Function Remove(ByRef s As String, ByVal StartIndex As Long, ByVal Count As Long) As String
    Dim CharCount As Long
    CharCount = Len(s)
    
    Dim Result As Long
    Result = VerifyListRange(CharCount, StartIndex, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, StartIndex, "StartIndex", Count, "Count")
    
    Dim sp As Long
    sp = StrPtr(s)
    Remove = SysAllocStringLen(sp, CharCount - Count)
    Call CopyMemory(ByVal StrPtr(Remove) + StartIndex * 2, ByVal sp + (StartIndex + Count) * 2, (CharCount - (StartIndex + Count)) * 2)
End Function

''
' Returns the index of the first character found in the <i>OfAny<i> set of characters.
'
' @param s The string to be searched.
' @param OfAny The set of characters to search for (String or Integer() Array).
' @param Index The starting index of <i>s</i> to begin searching.
' @param Count The number of characters in <i>s</i> to search.
' @return The index of a character from the set, or -1 if none are found.
' @remarks The indexing of <i>s</i> is zero-based.
'
Public Function IndexOfAny(ByRef s As String, ByRef OfAny As Variant, Optional ByVal Index As Variant, Optional ByVal Count As Variant) As Long
    IndexOfAny = InternalIndexOfAny(s, OfAny, Index, Count, True)
End Function

''
' Returns the index of the first character found in the <i>OfAny<i> set of characters
' searching from the end of the string.
'
' @param s The string to be searched.
' @param OfAny The set of characters to search for (String or Integer() Array).
' @param Index The starting index of <i>s</i> to begin searching.
' @param Count The number of characters in <i>s</i> to search.
' @return The index of a character from the set, or -1 if none are found.
' @remarks The indexing of <i>s</i> is zero-based.
'
Public Function LastIndexOfAny(ByRef s As String, ByRef OfAny As Variant, Optional ByVal Index As Variant, Optional ByVal Count As Variant) As Long
    LastIndexOfAny = InternalIndexOfAny(s, OfAny, Index, Count, False)
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
        Throw Cor.NewArgumentOutOfRangeException("StartIndex")
    
    Dim ValueLength As Long
    ValueLength = Len(Value)
    
    If ValueLength = 0 Then
        Insert = s
        Exit Function
    End If
    
    Dim PtrS As Long
    PtrS = StrPtr(s)
    
    Insert = Space$(ValueLength + Len(s))
    
    Dim PtrInsert As Long
    PtrInsert = StrPtr(Insert)
    
    ValueLength = ValueLength * 2
    StartIndex = StartIndex * 2
    CopyMemory ByVal PtrInsert, ByVal PtrS, StartIndex
    CopyMemory ByVal PtrInsert + StartIndex, ByVal StrPtr(Value), ValueLength
    CopyMemory ByVal PtrInsert + StartIndex + ValueLength, ByVal PtrS + StartIndex, LenB(Insert) - (StartIndex + ValueLength)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friends
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function EqualsIgnoreCase(ByRef String1 As String, ByRef String2 As String) As Boolean
    EqualsIgnoreCase = mStringUtils.EqualStrings(String1, String2, vbTextCompare)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FormatCore(ByRef Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant)
    Dim Output As StringBuilder
    Set Output = StringBuilders.GetBuilder
    Output.InternalAppendFormat Provider, Format, Args
    FormatCore = StringBuilders.Release(Output)
End Function

Private Function InternalIndexOfAny(ByRef s As String, ByRef OfAny As Variant, ByRef Index As Variant, ByRef Count As Variant, ByVal ForwardSearch As Boolean) As Long
    mString.SA.pvData = StrPtr(s)
    mString.SA.cElements = Len(s)
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    If ForwardSearch Then
        Result = GetOptionalArrayRange(VarPtr(mString.SA), Index, ElemIndex, Count, ElemCount)
    Else
        Result = GetOptionalArrayRangeReverse(VarPtr(mString.SA), Index, ElemIndex, Count, ElemCount)
    End If
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "s", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))

    Dim Chars() As Integer
    Select Case VarType(OfAny)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(OfAny)
            mTrimChars.SA.cElements = Len(OfAny)
            SAPtr(Chars) = VarPtr(mTrimChars.SA)
            
        Case vbIntegerArray
            SAPtr(Chars) = GetArrayPointer(OfAny, True)

        Case Else
            Throw Cor.NewArgumentException("A String or Integer Array is required.", "OfAny")
    End Select

    If ForwardSearch Then
        InternalIndexOfAny = FindIndexOfAny(Chars, ElemIndex, ElemIndex + ElemCount - 1, 1)
    Else
        InternalIndexOfAny = FindIndexOfAny(Chars, ElemIndex, ElemIndex - ElemCount + 1, -1)
    End If

    SAPtr(Chars) = 0
End Function

Private Function FindIndexOfAny(ByRef Chars() As Integer, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal Direction As Long) As Long
    Dim lb As Long
    lb = LBound(Chars)
    
    Dim ub As Long
    ub = UBound(Chars)
    
    Dim i As Long
    For i = StartIndex To EndIndex Step Direction
        Dim j As Long
        For j = lb To ub
            If mString.Data(i) = Chars(j) Then
                FindIndexOfAny = i
                Exit Function
            End If
        Next j
    Next i
    
    FindIndexOfAny = -1
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
        Chars = AttachToString(s)

        Dim i As Long
        For i = 0 To UBound(Chars)
            If Chars(i) = Char Then
                CharCount = CharCount + 1
            End If
        Next

        DetachStringChars Chars
    End If
End Function

Private Function AttachToString(ByRef s As String) As Integer()
    InitAttachment StrPtr(s), Len(s), mStringChars
    SAPtr(AttachToString) = VarPtr(mStringChars)
End Function

Private Sub DetachStringChars(ByRef Chars() As Integer)
    Debug.Assert SAPtr(Chars) = VarPtr(mStringChars)
    
    mStringChars.cLocks = 0
    SAPtr(Chars) = vbNullPtr
End Sub

Private Function AttachToVariant(ByRef v As Variant) As Integer()
    InitAttachment StrPtr(v), Len(v), mVariantChars
    SAPtr(AttachToVariant) = VarPtr(mVariantChars)
End Function

Private Sub DetachVariantChars(ByRef Chars() As Integer)
    Debug.Assert SAPtr(Chars) = VarPtr(mVariantChars)
    
    mVariantChars.cLocks = 0
    SAPtr(Chars) = vbNullPtr
End Sub

Private Sub InitAttachment(ByVal Ptr As Long, ByVal Length As Long, ByRef Attachment As SafeArray1d)
    Debug.Assert Attachment.cLocks = 0
    
    Attachment.cbElements = 2
    Attachment.cDims = 1
    Attachment.cElements = Length
    Attachment.cLocks = 1
    Attachment.pvData = Ptr
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitWordBuffer mString, 0, 0
    InitWordBuffer mTrimChars, 0, 0
    
    With mJoinSA
        .cbElements = 4
        .cDims = 1
    End With
    SAPtr(mJoinArray) = VarPtr(mJoinSA)
End Sub

Private Sub Class_Terminate()
    SAPtr(mJoinArray) = 0
End Sub
