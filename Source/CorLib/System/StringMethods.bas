Attribute VB_Name = "StringMethods"
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: StringMethods
'

''
' These are methods generally forwarded by the CorString class. These exist here for
' optimization with large processing of strings in a looping or similar fashion.
'
' Calling public methods on the CorString public variable can be slow because the call has to go
' through a normal COM call, which can degrade performance. By moving some of the methods to
' a module VB can optimize calls increasing performance.
'
Option Explicit

Public Function Equals(ByRef a As String, ByRef b As String, ByVal ComparisonType As StringComparison) As Boolean
    If ComparisonType = StringComparison.TextCompare Then
        Equals = StrComp(a, b, vbTextCompare)
    Else
        Equals = CompareHelper(StrPtr(a), Len(a), StrPtr(b), Len(b), ComparisonType) = 0
    End If
End Function

' there are places where strings will be held in variants and we don't want to
' have to convert them to a string variable causing additional string allocations.
Public Function EqualsV(ByRef a As Variant, ByRef b As Variant, ByVal ComparisonType As StringComparison) As Boolean
    If ComparisonType = StringComparison.TextCompare Then
        EqualsV = StrComp(a, b, vbTextCompare)
    Else
        EqualsV = CompareHelper(StrPtr(a), Len(a), StrPtr(b), Len(b), ComparisonType) = 0
    End If
End Function

Public Function Compare(ByRef StrA As String, ByRef StrB As String, ByVal ComparisonType As StringComparison) As Long
    If ComparisonType = StringComparison.TextCompare Then
        Compare = StrComp(StrA, StrB, vbTextCompare)
    Else
        Compare = CompareHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), ComparisonType)
    End If
End Function

Public Function CompareV(ByRef StrA As Variant, ByRef StrB As Variant, ByVal ComparisonType As StringComparison) As Long
    If ComparisonType = StringComparison.TextCompare Then
        CompareV = StrComp(StrA, StrB, vbTextCompare)
    Else
        CompareV = CompareHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), ComparisonType)
    End If
End Function

Public Function CompareEx(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByVal Length As Long, ByVal ComparisonType As StringComparison) As Long
    Dim PtrA As Long
    Dim PtrB As Long
    Dim LengthA As Long
    Dim LengthB As Long
    
    If ComparisonType = StringComparison.TextCompare Then _
        Error.NotSupported NotSupported_StringComparison

    ValidateAndGetLengths StrA, IndexA, StrB, IndexB, Length, LengthA, LengthB
    GetOffsetPointers StrA, IndexA, StrB, IndexB, PtrA, PtrB
    CompareEx = CompareHelper(PtrA, LengthA, PtrB, LengthB, ComparisonType)
End Function

Private Function CompareHelper(ByVal PtrA As Long, ByVal LengthA As Long, ByVal PtrB As Long, ByVal LengthB As Long, ByVal ComparisonType As StringComparison) As Long
    Dim Result As Long
    
    ' do a quick check first
    If LengthA = 0 And LengthB = 0 Then
        Exit Function
    End If
    
    Select Case ComparisonType
        Case StringComparison.Ordinal
            Result = CompareStringOrdinal(PtrA, LengthA, PtrB, LengthB, BOOL_FALSE)
        Case StringComparison.OrdinalIgnoreCase
            Result = CompareStringOrdinal(PtrA, LengthA, PtrB, LengthB, BOOL_TRUE)
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
    
    ' CompareString and CompareStringOrdinal return a windows defined result so we adjust for typical comparison results.
    ' CSTR_LESS_THAN = 1
    ' CSTR_EQUAL = 2
    ' CSTR_GREATER_THAN = 3
    CompareHelper = Result - 2
End Function

Public Function CompareCultural(ByRef StrA As String, ByRef StrB As String, ByRef Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    CompareCultural = CompareCulturalHelper(StrPtr(StrA), Len(StrA), StrPtr(StrB), Len(StrB), Culture, Options)
End Function

Public Function CompareCulturalEx(ByRef StrA As String, ByVal IndexA As Long, ByRef StrB As String, ByVal IndexB As Long, ByVal Length As Long, ByRef Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    Dim PtrA As Long
    Dim PtrB As Long
    Dim LengthA As Long
    Dim LengthB As Long
    
    ValidateAndGetLengths StrA, IndexA, StrB, IndexB, Length, LengthA, LengthB
    GetOffsetPointers StrA, IndexA, StrB, IndexB, PtrA, PtrB
    CompareCulturalEx = CompareCulturalHelper(PtrA, LengthA, PtrB, LengthB, Culture, Options)
End Function

Private Function CompareCulturalHelper(ByVal PtrA As Long, ByVal LengthA As Long, ByVal PtrB As Long, ByVal LengthB As Long, ByRef Culture As CultureInfo, ByVal Options As CompareOptions) As Long
    Dim Flags As Long
    
    If Culture Is Nothing Then _
        Error.ArgumentNull "Culture"
        
    If LengthA = 0 And LengthB = 0 Then
        Exit Function
    End If
    
    Select Case Options
        Case CompareOptions.OrdinalOption
            CompareCulturalHelper = CompareStringOrdinal(PtrA, LengthA, PtrB, LengthB, BOOL_FALSE) - 2
        Case CompareOptions.OrdinalIgnoreCaseOption
            CompareCulturalHelper = CompareStringOrdinal(PtrA, LengthA, PtrB, LengthB, BOOL_TRUE) - 2
        Case Else
            Flags = GetCompareFlags(Options)
            CompareCulturalHelper = CompareString(Culture.LCID, Flags, PtrA, LengthA, PtrB, LengthB) - 2
    End Select
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


