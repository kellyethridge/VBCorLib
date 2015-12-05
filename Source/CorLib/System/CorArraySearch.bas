Attribute VB_Name = "CorArraySearch"
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
' Module: CorArrayHelper
'

''
' This helper class contains supporting code for the CorArray class.
'
Option Explicit


' Attempt to use a specialized search for a specific data type.
Public Function TrySZBinarySearch(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Length As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    RetVal = SZBinarySearch(pSA, VarPtr(CLng(Value)), StartIndex, Length, AddressOf SZCompareLongs)
        Case vbString:                  RetVal = SZBinarySearch(pSA, VarPtr(StrPtr(Value)), StartIndex, Length, AddressOf SZCompareStrings)
        Case vbDouble:                  RetVal = SZBinarySearch(pSA, VarPtr(CDbl(Value)), StartIndex, Length, AddressOf SZCompareDoubles)
        Case vbObject, vbDataObject:    RetVal = SZBinarySearch(pSA, VarPtr(Value), StartIndex, Length, AddressOf SZCompareComparables)
        Case vbInteger:                 RetVal = SZBinarySearch(pSA, VarPtr(CInt(Value)), StartIndex, Length, AddressOf SZCompareIntegers)
        Case vbSingle:                  RetVal = SZBinarySearch(pSA, VarPtr(CSng(Value)), StartIndex, Length, AddressOf SZCompareSingles)
        Case vbCurrency:                RetVal = SZBinarySearch(pSA, VarPtr(CCur(Value)), StartIndex, Length, AddressOf SZCompareCurrencies)
        Case vbDate:                    RetVal = SZBinarySearch(pSA, VarPtr(CDate(Value)), StartIndex, Length, AddressOf SZCompareDates)
        Case vbBoolean:                 RetVal = SZBinarySearch(pSA, VarPtr(CBool(Value)), StartIndex, Length, AddressOf SZCompareBooleans)
        Case vbByte:                    RetVal = SZBinarySearch(pSA, VarPtr(CByte(Value)), StartIndex, Length, AddressOf SZCompareBytes)
        Case Else
            Exit Function
    End Select
    TrySZBinarySearch = True
End Function

Public Function TrySZIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZIndexOf(pSA, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualLongs)
        Case vbString:                  RetVal = SZIndexOf(pSA, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualStrings)
        Case vbDouble:                  RetVal = SZIndexOf(pSA, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualDoubles)
        Case vbDate:                    RetVal = SZIndexOf(pSA, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualDates)
        Case vbObject, vbDataObject:    RetVal = SZIndexOf(pSA, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualObjects)
        Case vbInteger:                 RetVal = SZIndexOf(pSA, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualIntegers)
        Case vbSingle:                  RetVal = SZIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualSingles)
        Case vbByte:                    RetVal = SZIndexOf(pSA, VarPtr(CByte(Value)), StartIndex, Count, AddressOf EqualBytes)
        Case vbBoolean:                 RetVal = SZIndexOf(pSA, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualBooleans)
        Case vbCurrency:                RetVal = SZIndexOf(pSA, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualCurrencies)
        Case Else: Exit Function
    End Select
    TrySZIndexOf = True
End Function

Public Function TrySZLastIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZLastIndexOf(pSA, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualLongs)
        Case vbString:                  RetVal = SZLastIndexOf(pSA, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualStrings)
        Case vbDouble:                  RetVal = SZLastIndexOf(pSA, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualDoubles)
        Case vbDate:                    RetVal = SZLastIndexOf(pSA, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualDates)
        Case vbObject, vbDataObject:    RetVal = SZLastIndexOf(pSA, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualObjects)
        Case vbInteger:                 RetVal = SZLastIndexOf(pSA, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualIntegers)
        Case vbSingle:                  RetVal = SZLastIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualSingles)
        Case vbByte:                    RetVal = SZLastIndexOf(pSA, VarPtr(CByte(Value)), StartIndex, Count, AddressOf EqualBytes)
        Case vbBoolean:                 RetVal = SZLastIndexOf(pSA, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualBooleans)
        Case vbCurrency:                RetVal = SZLastIndexOf(pSA, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualCurrencies)
        Case Else: Exit Function
    End Select
    TrySZLastIndexOf = True
End Function

Private Function SZBinarySearch(ByVal ArrayPtr As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim ElemSize    As Long
    Dim pvData      As Long
    Dim pLowElem    As Long
    Dim pHighElem   As Long
    Dim Delegate    As Delegate
    Dim Comparer    As Func_T_T_Long
    
    ElemSize = SafeArrayGetElemsize(ArrayPtr)
    pvData = MemLong(ArrayPtr + PVDATA_OFFSET)
    pLowElem = Index - SafeArrayGetLBound(ArrayPtr, 1)
    pHighElem = pLowElem + Count - 1
    Set Comparer = InitDelegate(Delegate, ComparerAddress)
    
    Dim pMiddleElem As Long
    Do While pLowElem <= pHighElem
        pMiddleElem = (pLowElem + pHighElem) \ 2
        Select Case Comparer.Invoke(ByVal pvData + pMiddleElem * ElemSize, ByVal pValue)
            Case 0
                SZBinarySearch = pMiddleElem + SafeArrayGetLBound(ArrayPtr, 1)
                Exit Function
            Case Is > 0
                pHighElem = pMiddleElem - 1
            Case Else
                pLowElem = pMiddleElem + 1
        End Select
    Loop
    
    SZBinarySearch = (Not pLowElem) + SafeArrayGetLBound(ArrayPtr, 1)
End Function

Private Function SZIndexOf(ByVal ArrayPtr As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim pvData      As Long
    Dim ElemSize    As Long
    Dim Comparer    As Func_T_T_Boolean
    
    Set Comparer = NewDelegate(ComparerAddress)
    ElemSize = SafeArrayGetElemsize(ArrayPtr)
    pvData = MemLong(ArrayPtr + PVDATA_OFFSET)
    Index = Index - SafeArrayGetLBound(ArrayPtr, 1)
    
    Dim i As Long
    For i = Index To Index + Count - 1
        If Comparer.Invoke(ByVal pvData + i * ElemSize, ByVal pValue) Then
            SZIndexOf = i + SafeArrayGetLBound(ArrayPtr, 1)
            Exit Function
        End If
    Next

    SZIndexOf = SafeArrayGetLBound(ArrayPtr, 1) - 1
End Function

Private Function SZLastIndexOf(ByVal ArrayPtr As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim pvData      As Long
    Dim ElemSize    As Long
    Dim Comparer    As Func_T_T_Boolean
    
    Set Comparer = NewDelegate(ComparerAddress)
    ElemSize = SafeArrayGetElemsize(ArrayPtr)
    pvData = MemLong(ArrayPtr + PVDATA_OFFSET)
    Index = Index - SafeArrayGetLBound(ArrayPtr, 1)
    
    Dim i As Long
    For i = Index To Index - Count + 1 Step -1
        If Comparer.Invoke(ByVal pvData + i * ElemSize, ByVal pValue) Then
            SZLastIndexOf = i + SafeArrayGetLBound(ArrayPtr, 1)
            Exit Function
        End If
    Next

    SZLastIndexOf = SafeArrayGetLBound(ArrayPtr, 1) - 1
End Function

Private Function SZCompareLongs(ByRef x As Long, ByRef y As Long) As Long
    If x > y Then
        SZCompareLongs = 1
    ElseIf x < y Then
        SZCompareLongs = -1
    End If
End Function

Private Function SZCompareIntegers(ByRef x As Integer, ByRef y As Integer) As Long
    If x > y Then
        SZCompareIntegers = 1
    ElseIf x < y Then
        SZCompareIntegers = -1
    End If
End Function

Private Function SZCompareStrings(ByRef x As String, ByRef y As String) As Long
    SZCompareStrings = StrComp(x, y, vbBinaryCompare)
End Function

Private Function SZCompareDoubles(ByRef x As Double, ByRef y As Double) As Long
    If x > y Then
        SZCompareDoubles = 1
    ElseIf x < y Then
        SZCompareDoubles = -1
    End If
End Function

Private Function SZCompareSingles(ByRef x As Single, ByRef y As Single) As Long
    If x > y Then
        SZCompareSingles = 1
    ElseIf x < y Then
        SZCompareSingles = -1
    End If
End Function

Private Function SZCompareBytes(ByRef x As Byte, ByRef y As Byte) As Long
    If x > y Then
        SZCompareBytes = 1
    ElseIf x < y Then
        SZCompareBytes = -1
    End If
End Function

Private Function SZCompareBooleans(ByRef x As Boolean, ByRef y As Boolean) As Long
    If x > y Then
        SZCompareBooleans = 1
    ElseIf x < y Then
        SZCompareBooleans = -1
    End If
End Function

Private Function SZCompareDates(ByRef x As Date, ByRef y As Date) As Long
    SZCompareDates = DateDiff("s", y, x)
End Function

Private Function SZCompareCurrencies(ByRef x As Currency, ByRef y As Currency) As Long
    If x > y Then
        SZCompareCurrencies = 1
    ElseIf x < y Then
        SZCompareCurrencies = -1
    End If
End Function

Private Function SZCompareComparables(ByRef x As Object, ByRef y As Variant) As Long
    Dim XComparable As IComparable
    Set XComparable = x
    SZCompareComparables = XComparable.CompareTo(y)
End Function

Private Function EqualLongs(ByRef x As Long, ByRef y As Long) As Boolean
    EqualLongs = (x = y)
End Function

Private Function EqualStrings(ByRef x As String, ByRef y As String) As Boolean
    EqualStrings = CorString.Equals(x, y)
End Function

Private Function EqualDoubles(ByRef x As Double, ByRef y As Double) As Boolean
    EqualDoubles = (x = y)
End Function

Private Function EqualIntegers(ByRef x As Integer, ByRef y As Integer) As Boolean
    EqualIntegers = (x = y)
End Function

Private Function EqualSingles(ByRef x As Single, ByRef y As Single) As Boolean
    EqualSingles = (x = y)
End Function

Private Function EqualDates(ByRef x As Date, ByRef y As Date) As Boolean
    EqualDates = (DateDiff("s", x, y) = 0)
End Function

Private Function EqualBytes(ByRef x As Byte, ByRef y As Byte) As Boolean
    EqualBytes = (x = y)
End Function

Private Function EqualBooleans(ByRef x As Boolean, ByRef y As Boolean) As Boolean
    EqualBooleans = (x = y)
End Function

Private Function EqualCurrencies(ByRef x As Currency, ByRef y As Currency) As Boolean
    EqualCurrencies = (x = y)
End Function

Private Function EqualObjects(ByRef x As Object, ByRef y As Object) As Boolean
    If Not x Is Nothing Then
        If TypeOf x Is IObject Then
            Dim Obj As IObject
            Set Obj = x
            EqualObjects = Obj.Equals(y)
        Else
            EqualObjects = x Is y
        End If
    Else
        EqualObjects = y Is Nothing
    End If
End Function


