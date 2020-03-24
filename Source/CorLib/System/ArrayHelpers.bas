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
' Module: ArrayHelpers
'

''
' This helper class contains supporting code for the CorArray class.
'
Option Explicit

Public Type SortItems
    SA      As SafeArray1d
    Buffer  As Long
End Type

Private Type StringSortContext
    Keys()          As String
    KeyPtrs()       As Long
    LCID            As Long
    ComparisonType  As Long
    Comparer        As IComparer
End Type

Private mSortItems      As SortItems
Private mHasSortItems   As Boolean
Private mSortKeys       As SortItems

Public SortComparer     As IComparer

Public Function ReverseByteCopy(ByRef Bytes() As Byte) As Byte()
    Dim ub As Long
    ub = UBound(Bytes)
    
    Dim Ret() As Byte
    ReDim Ret(0 To ub)
    
    Dim i As Long
    For i = 0 To ub
        Ret(i) = Bytes(ub - i)
    Next i
    
    ReverseByteCopy = Ret
End Function

' Attempt to use a specialized search for a specific data type.
Public Function TrySZBinarySearch(ByRef Arr As Variant, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Length As Long, ByRef RetVal As Long) As Boolean
    Select Case VarType(Arr) And &HFF
        Case vbLong:                    RetVal = SZBinarySearch(Arr, VarPtr(CLng(Value)), StartIndex, Length, AddressOf SZCompareLongs)
        Case vbString:                  RetVal = SZBinarySearch(Arr, VarPtr(StrPtr(Value)), StartIndex, Length, AddressOf SZCompareStrings)
        Case vbDouble:                  RetVal = SZBinarySearch(Arr, VarPtr(CDbl(Value)), StartIndex, Length, AddressOf SZCompareDoubles)
        Case vbObject, vbDataObject:    RetVal = SZBinarySearch(Arr, VarPtr(Value), StartIndex, Length, AddressOf SZCompareComparables)
        Case vbInteger:                 RetVal = SZBinarySearch(Arr, VarPtr(CInt(Value)), StartIndex, Length, AddressOf SZCompareIntegers)
        Case vbSingle:                  RetVal = SZBinarySearch(Arr, VarPtr(CSng(Value)), StartIndex, Length, AddressOf SZCompareSingles)
        Case vbCurrency:                RetVal = SZBinarySearch(Arr, VarPtr(CCur(Value)), StartIndex, Length, AddressOf SZCompareCurrencies)
        Case vbDate:                    RetVal = SZBinarySearch(Arr, VarPtr(CDate(Value)), StartIndex, Length, AddressOf SZCompareDates)
        Case vbBoolean:                 RetVal = SZBinarySearch(Arr, VarPtr(CBool(Value)), StartIndex, Length, AddressOf SZCompareBooleans)
        Case vbByte:                    RetVal = SZBinarySearch(Arr, VarPtr(CByte(Value)), StartIndex, Length, AddressOf SZCompareBytes)
        Case vbUserDefinedType
            If IsInt64Array(Arr) Then
                RetVal = SZBinarySearch(Arr, VarPtr(CInt64(Value)), StartIndex, Length, AddressOf SZCompareCurrencies)
            Else
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    TrySZBinarySearch = True
End Function

Public Function TrySZIndexOf(ByRef Arr As Variant, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case VarType(Arr) And &HFF
        Case vbLong:                    RetVal = SZIndexOf(Arr, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualLongs)
        Case vbString:                  RetVal = SZIndexOf(Arr, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualStrings)
        Case vbDouble:                  RetVal = SZIndexOf(Arr, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualDoubles)
        Case vbDate:                    RetVal = SZIndexOf(Arr, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualDates)
        Case vbObject, vbDataObject:    RetVal = SZIndexOf(Arr, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualObjects)
        Case vbInteger:                 RetVal = SZIndexOf(Arr, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualIntegers)
        Case vbSingle:                  RetVal = SZIndexOf(Arr, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualSingles)
        Case vbByte:                    RetVal = SZIndexOf(Arr, VarPtr(CByte(Value)), StartIndex, Count, AddressOf EqualBytes)
        Case vbBoolean:                 RetVal = SZIndexOf(Arr, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualBooleans)
        Case vbCurrency:                RetVal = SZIndexOf(Arr, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualCurrencies)
        Case vbUserDefinedType
            If IsInt64Array(Arr) Then
                RetVal = SZIndexOf(Arr, VarPtr(CInt64(Value)), StartIndex, Count, AddressOf EqualCurrencies)
            Else
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    TrySZIndexOf = True
End Function

Public Function TrySZLastIndexOf(ByRef Arr As Variant, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case VarType(Arr) And &HFF
        Case vbLong:                    RetVal = SZLastIndexOf(Arr, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualLongs)
        Case vbString:                  RetVal = SZLastIndexOf(Arr, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualStrings)
        Case vbDouble:                  RetVal = SZLastIndexOf(Arr, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualDoubles)
        Case vbDate:                    RetVal = SZLastIndexOf(Arr, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualDates)
        Case vbObject, vbDataObject:    RetVal = SZLastIndexOf(Arr, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualObjects)
        Case vbInteger:                 RetVal = SZLastIndexOf(Arr, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualIntegers)
        Case vbSingle:                  RetVal = SZLastIndexOf(Arr, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualSingles)
        Case vbByte:                    RetVal = SZLastIndexOf(Arr, VarPtr(CByte(Value)), StartIndex, Count, AddressOf EqualBytes)
        Case vbBoolean:                 RetVal = SZLastIndexOf(Arr, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualBooleans)
        Case vbCurrency:                RetVal = SZLastIndexOf(Arr, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualCurrencies)
        Case vbUserDefinedType
            If IsInt64Array(Arr) Then
                RetVal = SZLastIndexOf(Arr, VarPtr(CInt64(Value)), StartIndex, Count, AddressOf EqualCurrencies)
            Else
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    TrySZLastIndexOf = True
End Function

Private Function SZBinarySearch(ByRef Arr As Variant, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim ArrayPtr    As Long
    Dim ElemSize    As Long
    Dim PVData      As Long
    Dim pLowElem    As Long
    Dim pHighElem   As Long
    Dim ComparerDel As Delegate
    Dim Comparer    As Func_T_T_Long
    
    ArrayPtr = SAPtrV(Arr)
    ElemSize = SafeArrayGetElemsize(ArrayPtr)
    PVData = MemLong(ArrayPtr + PVDATA_OFFSET)
    pLowElem = Index - LBound(Arr)
    pHighElem = pLowElem + Count - 1
    Set Comparer = InitDelegate(ComparerDel, ComparerAddress)
    
    Dim pMiddleElem As Long
    Do While pLowElem <= pHighElem
        pMiddleElem = (pLowElem + pHighElem) \ 2
        Select Case Comparer.Invoke(ByVal PVData + pMiddleElem * ElemSize, ByVal pValue)
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

Private Function SZIndexOf(ByRef Arr As Variant, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim ArrayPtr    As Long
    Dim PVData      As Long
    Dim ElemSize    As Long
    Dim Comparer    As Func_T_T_Boolean
    
    ArrayPtr = SAPtrV(Arr)
    Set Comparer = NewDelegate(ComparerAddress)
    ElemSize = SafeArrayGetElemsize(ArrayPtr)
    PVData = MemLong(ArrayPtr + PVDATA_OFFSET)
    Index = Index - SafeArrayGetLBound(ArrayPtr, 1)
    
    Dim i As Long
    For i = Index To Index + Count - 1
        If Comparer.Invoke(ByVal PVData + i * ElemSize, ByVal pValue) Then
            SZIndexOf = i + SafeArrayGetLBound(ArrayPtr, 1)
            Exit Function
        End If
    Next

    SZIndexOf = SafeArrayGetLBound(ArrayPtr, 1) - 1
End Function

Private Function SZLastIndexOf(ByRef Arr As Variant, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim ArrayPtr    As Long
    Dim PVData      As Long
    Dim ElemSize    As Long
    Dim Comparer    As Func_T_T_Boolean
    
    ArrayPtr = SAPtrV(Arr)
    Set Comparer = NewDelegate(ComparerAddress)
    ElemSize = SafeArrayGetElemsize(ArrayPtr)
    PVData = MemLong(ArrayPtr + PVDATA_OFFSET)
    Index = Index - SafeArrayGetLBound(ArrayPtr, 1)
    
    Dim i As Long
    For i = Index To Index - Count + 1 Step -1
        If Comparer.Invoke(ByVal PVData + i * ElemSize, ByVal pValue) Then
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
Public Function TrySZSort(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long) As Boolean
    Dim Sorter As Action_T_T_T
    Dim pfn As Long
    
    On Error GoTo Catch
    
    Select Case VarType(Keys) And &HFF
        Case vbLong:                    pfn = FuncAddr(AddressOf QuickSortLong)
        Case vbString
            SortStringsWithComparer Keys, Left, Right, Nothing
            TrySZSort = True
            Exit Function
        Case vbDouble, vbDate:          pfn = FuncAddr(AddressOf QuickSortDouble)
        Case vbObject, vbDataObject:    pfn = FuncAddr(AddressOf QuickSortObject)
        Case vbVariant:                 pfn = FuncAddr(AddressOf QuickSortVariant)
        Case vbInteger:                 pfn = FuncAddr(AddressOf QuickSortInteger)
        Case vbSingle:                  pfn = FuncAddr(AddressOf QuickSortSingle)
        Case vbByte:                    pfn = FuncAddr(AddressOf QuickSortByte)
        Case vbCurrency:                pfn = FuncAddr(AddressOf QuickSortCurrency)
        Case vbBoolean:                 pfn = FuncAddr(AddressOf QuickSortBoolean)
        Case vbUserDefinedType
            If IsInt64Array(Keys) Then
                ' we can sort Int64 as Currency because they are both a signed 64-bit number.
                pfn = FuncAddr(AddressOf QuickSortCurrency)
            Else
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    Set Sorter = NewDelegate(pfn)
    Sorter.Invoke SAPtrV(Keys), ByVal Left, ByVal Right

    TrySZSort = True
    Exit Function
    
Catch:
    If Err.Number = 13 Then _
        Error.InvalidOperation InvalidOperation_Comparer_Arg
    
    Throw Err
End Function

Public Sub SortStringsWithComparer(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long, ByVal Comparer As IComparer)
    If Comparer Is Nothing Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareString, Statics.Comparer.Default.LCID, CompareOptions.None
    ElseIf Comparer Is Statics.Comparer.Default Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareString, Statics.Comparer.Default.LCID, CompareOptions.None
    ElseIf Comparer Is Statics.Comparer.DefaultInvariant Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareString, Statics.Comparer.DefaultInvariant.LCID, CompareOptions.None
    ElseIf Comparer Is StringComparer.BinaryCompare Then
        SortStrings Keys, Left, Right, AddressOf QuickSortStrComp, 0, vbBinaryCompare
    ElseIf Comparer Is StringComparer.TextCompare Then
        SortStrings Keys, Left, Right, AddressOf QuickSortStrComp, 0, vbTextCompare
    ElseIf Comparer Is StringComparer.Ordinal Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareStringOrdinal, 0, BOOL.BOOL_FALSE
    ElseIf Comparer Is StringComparer.OrdinalIgnoreCase Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareStringOrdinal, 0, BOOL.BOOL_TRUE
    ElseIf Comparer Is StringComparer.InvariantCulture Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareString, CultureInfo.InvariantCulture.LCID, CompareOptions.None
    ElseIf Comparer Is StringComparer.InvariantCultureIgnoreCase Then
        SortStrings Keys, Left, Right, AddressOf QuickSortCompareString, CultureInfo.InvariantCulture.LCID, CompareOptions.IgnoreCase
    Else
        SortStrings Keys, Left, Right, AddressOf QuickSortStringComparer, 0, 0, Comparer
    End If
End Sub

Private Sub SortStrings(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long, ByVal pfn As Long, ByVal LCID As Long, ByVal ComparisonType As Long, Optional ByVal Comparer As IComparer)
    Dim Sorter  As Action_T_T_T
    Dim Context As StringSortContext
    Dim pSA     As Long
    
    Set Sorter = NewDelegate(pfn)
    
    Context.LCID = LCID
    Context.ComparisonType = ComparisonType
    Set Context.Comparer = Comparer
    
    pSA = SAPtrV(Keys)
    SAPtr(Context.Keys) = pSA
    SAPtr(Context.KeyPtrs) = pSA
    
    On Error GoTo Catch
    Sorter.Invoke Context, ByVal Left, ByVal Right
    GoSub Finally
    Exit Sub
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    SAPtr(Context.KeyPtrs) = vbNullPtr
    SAPtr(Context.Keys) = vbNullPtr
    Return
End Sub

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
            Case 1:     Helper.Swap1 ByVal .PVData + i, ByVal .PVData + j
            Case 2:     Helper.Swap2 ByVal .PVData + i * 2, ByVal .PVData + j * 2
            Case 4:     Helper.Swap4 ByVal .PVData + i * 4, ByVal .PVData + j * 4
            Case 8:     Helper.Swap8 ByVal .PVData + i * 8, ByVal .PVData + j * 8
            Case 16:    Helper.Swap16 ByVal .PVData + i * 16, ByVal .PVData + j * 16
            Case Else
                ' primarily for UDTs
                CopyMemory ByVal Items.Buffer, ByVal .PVData + i * .cbElements, .cbElements
                CopyMemory ByVal .PVData + i * .cbElements, ByVal .PVData + j * .cbElements, .cbElements
                CopyMemory ByVal .PVData + j * .cbElements, ByVal Items.Buffer, .cbElements
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

Private Sub QuickSortCompareStringOrdinal(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, t As Long, m As Long
    Dim PtrX As Long
    Dim LenX As Long
    
    If Not mHasSortItems Then
        If Right - Left < 15 Then
            InsertionSortCompareStringOrdinal Context, Left, Right
            Exit Sub
        End If
    End If
    
    Do While Left < Right
        i = Left: j = Right: m = (i + j) \ 2
        PtrX = Context.KeyPtrs(m)
        LenX = Len(Context.Keys(m))
        
        Do
            Do While CompareStringOrdinal(Context.KeyPtrs(i), Len(Context.Keys(i)), PtrX, LenX, Context.ComparisonType) = CSTR_LESS_THAN
                i = i + 1
            Loop
            
            Do While CompareStringOrdinal(Context.KeyPtrs(j), Len(Context.Keys(j)), PtrX, LenX, Context.ComparisonType) = CSTR_GREATER_THAN
                j = j - 1
            Loop
            
            If i > j Then Exit Do
            
            If i < j Then
                t = Context.KeyPtrs(i)
                Context.KeyPtrs(i) = Context.KeyPtrs(j)
                Context.KeyPtrs(j) = t
                
                If mHasSortItems Then SwapSortItems mSortItems, i, j
            End If
            
            i = i + 1: j = j - 1
        Loop While i <= j
        
        If j - Left <= Right - i Then
            If Left < j Then QuickSortCompareStringOrdinal Context, Left, j
            Left = i
        Else
            If i < Right Then QuickSortCompareStringOrdinal Context, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub InsertionSortCompareStringOrdinal(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim PtrX As Long
    Dim LenX As Long
    
    i = Left
    
    Do While i <= Right
        PtrX = Context.KeyPtrs(i)
        LenX = Len(Context.Keys(i))
        j = i - 1
        
        Do While j >= Left
            If CompareStringOrdinal(Context.KeyPtrs(j), Len(Context.Keys(j)), PtrX, LenX, Context.ComparisonType) <> CSTR_GREATER_THAN Then
                Exit Do
            End If
            
            Context.KeyPtrs(j + 1) = Context.KeyPtrs(j)
            j = j - 1
        Loop
        
        Context.KeyPtrs(j + 1) = PtrX
        i = i + 1
    Loop
End Sub

Private Sub QuickSortCompareString(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, t As Long, m As Long
    Dim PtrX As Long
    Dim LenX As Long
       
    If Not mHasSortItems Then
        If Right - Left < 15 Then
            InsertionSortCompareString Context, Left, Right
            Exit Sub
        End If
    End If
    
    Do While Left < Right
        i = Left: j = Right: m = (i + j) \ 2
        PtrX = Context.KeyPtrs(m)
        LenX = Len(Context.Keys(m))
        
        Do
            Do While CompareString(Context.LCID, Context.ComparisonType, Context.KeyPtrs(i), Len(Context.Keys(i)), PtrX, LenX) = CSTR_LESS_THAN
                i = i + 1
            Loop
            
            Do While CompareString(Context.LCID, Context.ComparisonType, Context.KeyPtrs(j), Len(Context.Keys(j)), PtrX, LenX) = CSTR_GREATER_THAN
                j = j - 1
            Loop
            
            If i > j Then Exit Do
            
            If i < j Then
                t = Context.KeyPtrs(i)
                Context.KeyPtrs(i) = Context.KeyPtrs(j)
                Context.KeyPtrs(j) = t
                
                If mHasSortItems Then SwapSortItems mSortItems, i, j
            End If
            
            i = i + 1: j = j - 1
        Loop While i <= j
        
        If j - Left <= Right - i Then
            If Left < j Then QuickSortCompareString Context, Left, j
            Left = i
        Else
            If i < Right Then QuickSortCompareString Context, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub InsertionSortCompareString(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim PtrX As Long
    Dim LenX As Long
    
    i = Left
    
    Do While i <= Right
        PtrX = Context.KeyPtrs(i)
        LenX = Len(Context.Keys(i))
        j = i - 1
        
        Do While j >= Left
            If CompareString(Context.LCID, Context.ComparisonType, Context.KeyPtrs(j), Len(Context.Keys(j)), PtrX, LenX) <> CSTR_GREATER_THAN Then
                Exit Do
            End If
            
            Context.KeyPtrs(j + 1) = Context.KeyPtrs(j)
            j = j - 1
        Loop
        
        Context.KeyPtrs(j + 1) = PtrX
        i = i + 1
    Loop
End Sub

Private Sub QuickSortStrComp(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, t As Long, x As String
    
    If Not mHasSortItems Then
        If Right - Left < 15 Then
            InsertionSortStrComp Context, Left, Right
            Exit Sub
        End If
    End If
    
    Do While Left < Right
        i = Left: j = Right
        StringPtr(x) = Context.KeyPtrs((i + j) \ 2)
        
        Do
            Do While StrComp(Context.Keys(i), x, Context.ComparisonType) < 0: i = i + 1: Loop
            Do While StrComp(Context.Keys(j), x, Context.ComparisonType) > 0: j = j - 1: Loop
            
            If i > j Then Exit Do
            
            If i < j Then
                t = Context.KeyPtrs(i)
                Context.KeyPtrs(i) = Context.KeyPtrs(j)
                Context.KeyPtrs(j) = t
                
                If mHasSortItems Then SwapSortItems mSortItems, i, j
            End If
            
            i = i + 1: j = j - 1
        Loop While i <= j
        
        If j - Left <= Right - i Then
            If Left < j Then QuickSortStrComp Context, Left, j
            Left = i
        Else
            If i < Right Then QuickSortStrComp Context, i, Right
            Right = j
        End If
        
        StringPtr(x) = vbNullPtr
    Loop
End Sub

Private Sub InsertionSortStrComp(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim x As String
    Dim PtrX As Long
    
    i = Left
    
    Do While i <= Right
        PtrX = Context.KeyPtrs(i)
        StringPtr(x) = PtrX
        j = i - 1
        
        Do While j >= Left
            If StrComp(Context.Keys(j), x, Context.ComparisonType) <= 0 Then
                Exit Do
            End If
            
            Context.KeyPtrs(j + 1) = Context.KeyPtrs(j)
            j = j - 1
        Loop
        
        Context.KeyPtrs(j + 1) = PtrX
        StringPtr(x) = vbNullPtr
        i = i + 1
    Loop
End Sub

Private Sub QuickSortStringComparer(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, t As Long, x As String
    
    If Not mHasSortItems Then
        If Right - Left < 15 Then
            InsertionSortStrComp Context, Left, Right
            Exit Sub
        End If
    End If
    
    Do While Left < Right
        i = Left: j = Right
        StringPtr(x) = Context.KeyPtrs((i + j) \ 2)
        
        Do
            Do While Context.Comparer.Compare(Context.Keys(i), x) < 0: i = i + 1: Loop
            Do While Context.Comparer.Compare(Context.Keys(j), x) > 0: j = j - 1: Loop
            
            If i > j Then Exit Do
            
            If i < j Then
                t = Context.KeyPtrs(i)
                Context.KeyPtrs(i) = Context.KeyPtrs(j)
                Context.KeyPtrs(j) = t
                
                If mHasSortItems Then SwapSortItems mSortItems, i, j
            End If
            
            i = i + 1: j = j - 1
        Loop While i <= j
        
        If j - Left <= Right - i Then
            If Left < j Then QuickSortStringComparer Context, Left, j
            Left = i
        Else
            If i < Right Then QuickSortStringComparer Context, i, Right
            Right = j
        End If
        
        StringPtr(x) = vbNullPtr
    Loop
End Sub

Private Sub InsertionSortStringComparer(ByRef Context As StringSortContext, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim x As String
    Dim PtrX As Long
    
    i = Left
    
    Do While i <= Right
        PtrX = Context.KeyPtrs(i)
        StringPtr(x) = PtrX
        j = i - 1
        
        Do While j >= Left
            If Context.Comparer.Compare(Context.Keys(j), x) <= 0 Then
                Exit Do
            End If
            
            Context.KeyPtrs(j + 1) = Context.KeyPtrs(j)
            j = j - 1
        Loop
        
        Context.KeyPtrs(j + 1) = PtrX
        StringPtr(x) = vbNullPtr
        i = i + 1
    Loop
End Sub

Private Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant, Key As IComparable
    Do While Left < Right
        i = Left: j = Right: Set x = Keys((i + j) \ 2)
        Do
            Do While Comparer.Default.Compare(Keys(i), x) < 0: i = i + 1: Loop
            Do While Comparer.Default.Compare(Keys(j), x) > 0: j = j - 1: Loop
            
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

Public Sub QuickSortGeneral(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long, ByRef Comparer As IComparer)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd x, Keys((i + j) \ 2)
        Do
            Do While Comparer.Compare(Keys(i), x) < 0: i = i + 1: Loop
            Do While Comparer.Compare(Keys(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then SwapSortItems mSortKeys, i, j: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortGeneral Keys, Left, j, Comparer
            Left = i
        Else
            If i < Right Then QuickSortGeneral Keys, i, Right, Comparer
            Right = j
        End If
    Loop
End Sub


