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
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: cString
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

Private mBuilder        As StringBuilder
Private mString         As WordBuffer
Private mTrimChars      As WordBuffer
Private mJoinSA         As SafeArray1d
Private mJoinArray()    As String



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
' Returns a string from the characters in an integer array.
'
' @param Chars The array to create the string from.
' @param Index The starting index in the array.
' @param Count The number of characters to use.
' @return A string created from the chars in an integer array.
'
Public Function FromCharArray(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As String
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    
    Result = GetOptionalArrayRange(SAPtr(Chars), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Chars", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    If ElemCount = 0 Then
        ' We don't want a Null string returned.
        FromCharArray = ""
        Exit Function
    End If
    
    FromCharArray = SysAllocStringLen(VarPtr(Chars(ElemIndex)), ElemCount)
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
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16))
    Format = FormatArrayEx(Nothing, FormatText, vArgs)
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
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20))
    FormatEx = FormatArrayEx(Provider, FormatText, vArgs)
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
    FormatArray = FormatArrayEx(Nothing, FormatText, Args)
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
    mBuilder.Length = 0
    Call mBuilder.InternalAppendFormat(Provider, FormatText, Args)
    FormatArrayEx = mBuilder.ToString
End Function

''
' Creates and array of chars (Integers) from the specified string.
'
' @param s The string to create the chars from.
' @param startindex The start index in the string to begin converting to a char array. This is zero-based.
' @param length The number of characters to convert to chars.
' @return An array containing the converted characters from the string.
'
Public Function ToCharArray(ByRef s As String, Optional ByRef startIndex As Variant, Optional ByRef Length As Variant) As Integer()
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(Len(s), startIndex, ElemIndex, Length, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "StartIndex", ElemCount, "Length", IsMissing(startIndex))
    
    Dim Ret() As Integer
    If ElemCount > 0 Then
        ReDim Ret(0 To ElemCount - 1)
        Call CopyMemory(Ret(0), ByVal StrPtr(s) + ElemIndex * 2, ElemCount * 2)
    Else
        Ret = Cor.NewIntegers()
    End If
    ToCharArray = Ret
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
Public Function PadLeft(ByRef s As String, ByVal totalWidth As Long, Optional ByRef paddingChar As Variant = 32) As String
    Dim l As Long
    l = Len(s)
    If totalWidth > l Then
        PadLeft = String$(totalWidth - l, paddingChar) & s
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
Public Function PadRight(ByRef s As String, ByVal totalWidth As Long, Optional ByRef paddingChar As Variant = 32) As String
    Dim l As Long
    l = Len(s)
    If totalWidth > l Then
        PadRight = s & String$(totalWidth - l, paddingChar)
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
    Dim Chars() As Integer
    Dim startIndex As Long
    
    Select Case VarType(TrimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(TrimChars)
            mTrimChars.SA.cElements = Len(TrimChars)
            startIndex = TrimStartIndex(s, mTrimChars.Data)
            
        Case vbIntegerArray
            SAPtr(Chars) = GetArrayPointer(TrimChars, True)
            startIndex = TrimStartIndex(s, Chars)
            SAPtr(Chars) = 0
        
        Case vbError
            startIndex = SzTrimStartIndex(s)
        
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
    
    If startIndex = 0 Then Exit Function
    TrimStart = Mid$(s, startIndex)
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
    Dim Chars()     As Integer
    Dim EndIndex    As Long
    
    Select Case VarType(TrimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(TrimChars)
            mTrimChars.SA.cElements = Len(TrimChars)
            EndIndex = TrimEndIndex(s, mTrimChars.Data)
            
        Case vbIntegerArray
            SAPtr(Chars) = GetArrayPointer(TrimChars, True)
            EndIndex = TrimEndIndex(s, Chars)
            SAPtr(Chars) = 0
        
        Case vbError
            EndIndex = SzTrimEndIndex(s)
        
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
    
    If EndIndex = 0 Then Exit Function
    TrimEnd = Left$(s, EndIndex)
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
    Dim Chars() As Integer
    Dim startIndex As Long
    Dim EndIndex As Long
    
    Select Case VarType(TrimChars)
        Case vbString
            mTrimChars.SA.pvData = StrPtr(TrimChars)
            mTrimChars.SA.cElements = Len(TrimChars)
            startIndex = TrimStartIndex(s, mTrimChars.Data)
            EndIndex = TrimEndIndex(s, mTrimChars.Data)
            
        Case vbIntegerArray
            SAPtr(Chars) = GetArrayPointer(TrimChars, True)
            startIndex = TrimStartIndex(s, Chars)
            EndIndex = TrimEndIndex(s, Chars)
            SAPtr(Chars) = 0
            
        Case vbError    ' assume missing
            startIndex = SzTrimStartIndex(s)
            EndIndex = SzTrimEndIndex(s)
        
        Case Else
            Throw Cor.NewArgumentException("Invalid character set.")
    End Select
    
    If startIndex = 0 Or EndIndex = 0 Then Exit Function
    Trim = Mid$(s, startIndex, EndIndex - startIndex + 1)
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
Public Function Remove(ByRef s As String, ByVal startIndex As Long, ByVal Count As Long) As String
    Dim CharCount As Long
    CharCount = Len(s)
    
    Dim Result As Long
    Result = VerifyListRange(CharCount, startIndex, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, startIndex, "StartIndex", Count, "Count")
    
    Dim sp As Long
    sp = StrPtr(s)
    Remove = SysAllocStringLen(sp, CharCount - Count)
    Call CopyMemory(ByVal StrPtr(Remove) + startIndex * 2, ByVal sp + (startIndex + Count) * 2, (CharCount - (startIndex + Count)) * 2)
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
' @param Index The starting position to insert the string.
' @param Value The string to be inserted.
' @return A string containing the new string created from the insertion.
' @remarks Index is zero-based.
'
Public Function Insert(ByRef s As String, ByVal Index As Long, ByRef Value As String) As String
    If Index < 0 Or Index > Len(s) Then _
        Throw Cor.NewArgumentOutOfRangeException(cString.Format("Index must be between 0 and {0}.", Len(s)), "Index", Index)
    
    Dim ValueLength As Long
    ValueLength = Len(Value)
    
    If ValueLength = 0 Then
        Insert = s
        Exit Function
    End If
    
    Dim PtrS As Long
    PtrS = StrPtr(s)
    
    Insert = Space$(ValueLength + Len(s))
    
    Dim ptrInsert As Long
    ptrInsert = StrPtr(Insert)
    
    ValueLength = ValueLength * 2
    Index = Index * 2
    Call CopyMemory(ByVal ptrInsert, ByVal PtrS, Index)
    Call CopyMemory(ByVal ptrInsert + Index, ByVal StrPtr(Value), ValueLength)
    Call CopyMemory(ByVal ptrInsert + Index + ValueLength, ByVal PtrS + Index, LenB(Insert) - (Index + ValueLength))
End Function

Public Function CharCount(ByRef s As String, ByVal Char As Variant) As Long
    Dim ch As Integer
    Select Case VarType(Char)
        Case vbLong, vbInteger, vbByte: ch = Char
        Case vbString:                  ch = AscW(Char)
        Case Else:                      Throw Cor.NewArgumentException("Invalid Char type.", "Char")
    End Select
    
    mString.SA.pvData = StrPtr(s)
    mString.SA.cElements = Len(s)
    
    Dim i As Long
    For i = 0 To Len(s) - 1
        If mString.Data(i) = ch Then CharCount = CharCount + 1
    Next i
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Private Function FindIndexOfAny(ByRef Chars() As Integer, ByVal startIndex As Long, ByVal EndIndex As Long, ByVal Direction As Long) As Long
    Dim lb As Long
    lb = LBound(Chars)
    
    Dim ub As Long
    ub = UBound(Chars)
    
    Dim i As Long
    For i = startIndex To EndIndex Step Direction
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

Private Function TrimStartIndex(ByRef s As String, ByRef TrimChars() As Integer) As Long
    Dim i   As Long
    Dim lb  As Long
    Dim ub  As Long
    Dim ch  As Integer
    Dim pos As Long
    Dim l   As Long
    
    lb = LBound(TrimChars)
    ub = UBound(TrimChars)
    l = Len(s)
    With mString
        .SA.pvData = StrPtr(s)
        .SA.cElements = Len(s)
        Do While pos < l
            ch = .Data(pos)
            For i = lb To ub
                If ch = TrimChars(i) Then Exit For
            Next i
            If i > ub Then Exit Do
            pos = pos + 1
        Loop
        If pos = l Then Exit Function
        TrimStartIndex = pos + 1
    End With
End Function

' This method is optimized for the default set of whitespace.
Private Function SzTrimStartIndex(ByRef s As String) As Long
    Dim i As Long
    
    With mString
        .SA.pvData = StrPtr(s)
        .SA.cElements = Len(s)
        For i = 0 To .SA.cElements - 1
            Select Case .Data(i)
                Case &H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF
                Case Else
                    SzTrimStartIndex = i + 1
                    Exit Function
            End Select
        Next i
    End With
End Function

Private Function TrimEndIndex(ByRef s As String, ByRef TrimChars() As Integer) As Long
    Dim i   As Long
    Dim lb  As Long
    Dim ub  As Long
    Dim ch  As Integer
    Dim pos As Long
    
    lb = LBound(TrimChars)
    ub = UBound(TrimChars)
    With mString
        .SA.pvData = StrPtr(s)
        .SA.cElements = Len(s)
        
        pos = .SA.cElements - 1
        Do While pos >= 0
            ch = .Data(pos)
            For i = lb To ub
                If ch = TrimChars(i) Then Exit For
            Next i
            If i > ub Then Exit Do
            pos = pos - 1
        Loop
        If pos < 0 Then Exit Function
        TrimEndIndex = pos + 1
    End With
End Function

' This method is optimized for the default set of whitespace.
Private Function SzTrimEndIndex(ByRef s As String) As Long
    Dim i As Long
    
    With mString
        .SA.pvData = StrPtr(s)
        .SA.cElements = Len(s)
        For i = .SA.cElements - 1 To 0 Step -1
            Select Case .Data(i)
                Case &H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF
                Case Else
                    SzTrimEndIndex = i + 1
                    Exit Function
            End Select
        Next i
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBuilder = New StringBuilder
    Call InitWordBuffer(mString, 0, 0)
    Call InitWordBuffer(mTrimChars, 0, 0)
    
    With mJoinSA
        .cbElements = 4
        .cDims = 1
    End With
    SAPtr(mJoinArray) = VarPtr(mJoinSA)
End Sub

Private Sub Class_Terminate()
    SAPtr(mJoinArray) = 0
End Sub
