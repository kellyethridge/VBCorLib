Attribute VB_Name = "ArrayHelpers"
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

Public Type SortItems
    SA      As SafeArray1d
    Buffer  As Long
End Type

Private mSortItems      As SortItems
Private mHasSortItems   As Boolean
Private mSortKeys       As SortItems
Public SortComparer     As IComparer


''
' Returns a pointer to the SafeArray structure.
'
' If a non-array type is passed in, then zero is returned.
'
' We no longer consider an empty array of objects or UDT's to
' have been null originally. We now simply consider it an empty array.
'
Public Function GetArrayPointer(ByRef Arg As Variant) As Long
    If IsArray(Arg) Then
        GetArrayPointer = MemLong(VB6.vbaRefVarAry(Arg))
    End If
End Function

Public Function VSAPtr(ByRef Value As Variant) As Long
    If Not IsArray(Value) Then _
        Error.Argument Argument_ArrayRequired
        
    VSAPtr = MemLong(vbaRefVarAry(Value))
End Function

Public Function Len1D(ByRef Arr As Variant) As Long
    Len1D = UBound(Arr) - LBound(Arr) + 1
End Function

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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Optimized sort routines. There could have been one
'   all-purpose sort routine, but it would be too slow.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TrySZSort(ByVal pSA As Long, ByVal Left As Long, ByVal Right As Long) As Boolean
    Dim pfn As Long
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    pfn = FuncAddr(AddressOf QuickSortLong)
        Case vbString:                  pfn = FuncAddr(AddressOf QuickSortString)
        Case vbDouble, vbDate:          pfn = FuncAddr(AddressOf QuickSortDouble)
        Case vbObject, vbDataObject:    pfn = FuncAddr(AddressOf QuickSortObject)
        Case vbVariant:                 pfn = FuncAddr(AddressOf QuickSortVariant)
        Case vbInteger:                 pfn = FuncAddr(AddressOf QuickSortInteger)
        Case vbSingle:                  pfn = FuncAddr(AddressOf QuickSortSingle)
        Case vbByte:                    pfn = FuncAddr(AddressOf QuickSortByte)
        Case vbCurrency:                pfn = FuncAddr(AddressOf QuickSortCurrency)
        Case vbBoolean:                 pfn = FuncAddr(AddressOf QuickSortBoolean)
        Case Else: Exit Function
    End Select
    
    Dim Sorter As Action_T_T_T
    Set Sorter = NewDelegate(pfn)
    Sorter.Invoke pSA, ByVal Left, ByVal Right

    TrySZSort = True
End Function

Public Function TrySZSortWithItems(ByVal KeysPtr As Long, ByVal Left As Long, ByVal Right As Long, ByVal ItemsPtr) As Boolean

End Function

Public Sub SetSortKeys(ByVal pSA As Long)
    CopyMemory mSortKeys.SA, ByVal pSA, vbSizeOfSafeArray1d
    Select Case mSortKeys.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else: mSortKeys.Buffer = CoTaskMemAlloc(mSortKeys.SA.cbElements)
    End Select
End Sub

Public Sub ClearSortKeys()
    If mSortKeys.Buffer Then
        CoTaskMemFree mSortKeys.Buffer
    End If
    mSortKeys.Buffer = 0
End Sub

Public Sub SetSortItems(ByVal pSA As Long)
    CopyMemory mSortItems.SA, ByVal pSA, vbSizeOfSafeArray1d
    Select Case mSortItems.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else
            mSortItems.Buffer = CoTaskMemAlloc(mSortItems.SA.cbElements)
    End Select
    mHasSortItems = True
End Sub

Public Sub ClearSortItems()
    If mHasSortItems Then
        CoTaskMemFree mSortItems.Buffer
        mSortItems.Buffer = 0
        mHasSortItems = False
    End If
End Sub

Public Sub SwapSortItems(ByRef Items As SortItems, ByVal i As Long, ByVal j As Long)
    With Items.SA
        Select Case .cbElements
            Case 1:     Helper.Swap1 ByVal .pvData + i, ByVal .pvData + j
            Case 2:     Helper.Swap2 ByVal .pvData + i * 2, ByVal .pvData + j * 2
            Case 4:     Helper.Swap4 ByVal .pvData + i * 4, ByVal .pvData + j * 4
            Case 8:     Helper.Swap8 ByVal .pvData + i * 8, ByVal .pvData + j * 8
            Case 16:    Helper.Swap16 ByVal .pvData + i * 16, ByVal .pvData + j * 16
            Case Else
                ' primarily for UDTs
                CopyMemory ByVal Items.Buffer, ByVal .pvData + i * .cbElements, .cbElements
                CopyMemory ByVal .pvData + i * .cbElements, ByVal .pvData + j * .cbElements, .cbElements
                CopyMemory ByVal .pvData + j * .cbElements, ByVal Items.Buffer, .cbElements
        End Select
    End With
End Sub

Private Sub QuickSortLong(ByRef Keys() As Long, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Long, t As Long
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortLong Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortLong Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortString(ByRef Keys() As String, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As String
    Do While Left < Right
        i = Left: j = Right: x = StringRef(Keys((i + j) \ 2))
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap4 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortString Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortString Keys, i, Right
            Right = j
        End If
        StringPtr(x) = 0
    Loop
End Sub

Private Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant, Key As IComparable
    Do While Left < Right
        i = Left: j = Right: Set x = Keys((i + j) \ 2)
        Do
            Set Key = Keys(i): Do While Key.CompareTo(x) < 0: i = i + 1: Set Key = Keys(i): Loop
            Set Key = Keys(j): Do While Key.CompareTo(x) > 0: j = j - 1: Set Key = Keys(j): Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap4 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortObject Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortObject Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortInteger(ByRef Keys() As Integer, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Integer, t As Integer
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortInteger Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortInteger Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortByte(ByRef Keys() As Byte, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Byte, t As Byte
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortByte Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortByte Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortDouble(ByRef Keys() As Double, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Double, t As Double
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortDouble Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortDouble Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortSingle(ByRef Keys() As Single, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Single, t As Single
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortSingle Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortSingle Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortCurrency(ByRef Keys() As Currency, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Currency, t As Currency
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortCurrency Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortCurrency Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortBoolean(ByRef Keys() As Boolean, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Boolean, t As Boolean
    Do While Left < Right
        i = Left: j = Right: x = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < x: i = i + 1: Loop
            Do While Keys(j) > x: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortBoolean Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortBoolean Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortVariant(ByRef Keys() As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd x, Keys((i + j) \ 2)
        Do
            Do While Comparer.Default.Compare(Keys(i), x) < 0: i = i + 1: Loop
            Do While Comparer.Default.Compare(Keys(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap16 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortVariant Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortVariant Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Public Sub QuickSortGeneral(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd x, Keys((i + j) \ 2)
        Do
            Do While SortComparer.Compare(Keys(i), x) < 0: i = i + 1: Loop
            Do While SortComparer.Compare(Keys(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then SwapSortItems mSortKeys, i, j: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortGeneral Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortGeneral Keys, i, Right
            Right = j
        End If
    Loop
End Sub


