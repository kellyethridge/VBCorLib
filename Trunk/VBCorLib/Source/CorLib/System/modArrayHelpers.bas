Attribute VB_Name = "modArrayHelpers"
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
'    Module: modArrayHelpers
'
Option Explicit

Private Declare Function vbaVarRefAry Lib "MSVBVM60.DLL" Alias "__vbaRefVarAry" (ByRef ArrayToDeref As Variant) As Long

Public Type SortItems
    SA      As SafeArray1d
    Buffer  As Long
End Type

Private mSortItems      As SortItems
Private mHasSortItems   As Boolean
Private mSortKeys       As SortItems
Public SortComparer     As IComparer

''
' Retrieves the pointer to an array's SafeArray structure.
'
' @param arr The array to retrieve the pointer to.
' @return A pointer to a SafeArray structure or 0 if the array is null.
'
Public Function GetArrayPointer(ByRef Arr As Variant, Optional ByVal ThrowOnNull As Boolean = False) As Long
    If Not IsArray(Arr) Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_ArrayRequired), "Arr")
    
    GetArrayPointer = MemLong(vbaVarRefAry(Arr))
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any cArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case VariantType(Arr) And &HFF
        Case vbObject, vbUserDefinedType
            If UBound(Arr) < LBound(Arr) Then
                GetArrayPointer = vbNullPtr
            End If
    End Select
    
    If ThrowOnNull Then
        If GetArrayPointer = vbNullPtr Then
            Throw Cor.NewArgumentNullException("Arr", Resources.GetString(ArgumentNull_Array))
        End If
    End If
End Function


' This is a set of comparison routines used by function delegation calls.
' They allow a virtual comparison routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
Public Function CompareLongs(ByRef x As Long, ByRef y As Long) As Long
    If x > y Then
        CompareLongs = 1
    ElseIf x < y Then
        CompareLongs = -1
    End If
End Function
Public Function CompareIntegers(ByRef x As Integer, ByRef y As Integer) As Long
    If x > y Then
        CompareIntegers = 1
    ElseIf x < y Then
        CompareIntegers = -1
    End If
End Function
Public Function CompareStrings(ByRef x As String, ByRef y As String) As Long
    If x > y Then
        CompareStrings = 1
    ElseIf x < y Then
        CompareStrings = -1
    End If
End Function
Public Function CompareDoubles(ByRef x As Double, ByRef y As Double) As Long
    If x > y Then
        CompareDoubles = 1
    ElseIf x < y Then
        CompareDoubles = -1
    End If
End Function
Public Function CompareSingles(ByRef x As Single, ByRef y As Single) As Long
    If x > y Then
        CompareSingles = 1
    ElseIf x < y Then
        CompareSingles = -1
    End If
End Function
Public Function CompareBytes(ByRef x As Byte, ByRef y As Byte) As Long
    If x > y Then
        CompareBytes = 1
    ElseIf x < y Then
        CompareBytes = -1
    End If
End Function
Public Function CompareBooleans(ByRef x As Boolean, ByRef y As Boolean) As Long
    If x > y Then
        CompareBooleans = 1
    ElseIf x < y Then
        CompareBooleans = -1
    End If
End Function
Public Function CompareDates(ByRef x As Date, ByRef y As Date) As Long
    CompareDates = DateDiff("s", y, x)
End Function
Public Function CompareCurrencies(ByRef x As Currency, ByRef y As Currency) As Long
    If x > y Then CompareCurrencies = 1: Exit Function
    If x < y Then CompareCurrencies = -1
End Function
Public Function CompareIComparable(ByRef x As Object, ByRef y As Variant) As Long
    Dim comparableX As IComparable
    Set comparableX = x
    CompareIComparable = comparableX.CompareTo(y)
End Function
Public Function CompareVariants(ByRef x As Variant, ByRef y As Variant) As Long
    Dim Comparable As IComparable
    
    Select Case VarType(x)
        Case vbNull
            If Not IsNull(y) Then
                CompareVariants = -1
            End If
            Exit Function
        Case vbEmpty
            If Not IsEmpty(y) Then
                CompareVariants = -1
            End If
            Exit Function
        Case vbObject, vbDataObject
            If TypeOf x Is IComparable Then
                Set Comparable = x
                CompareVariants = Comparable.CompareTo(y)
                Exit Function
            End If
        Case VarType(y)
            If x < y Then
                CompareVariants = -1
            ElseIf x > y Then
                CompareVariants = 1
            End If
            Exit Function
    End Select
    
    Select Case VarType(y)
        Case vbNull, vbEmpty
            CompareVariants = 1
        Case vbObject, vbDataObject
            If TypeOf y Is IComparable Then
                Set Comparable = y
                CompareVariants = -Comparable.CompareTo(x)
                Exit Function
            Else
                Throw Cor.NewArgumentException("Object must implement IComparable interface.")
            End If
        Case Else
            Throw Cor.NewInvalidOperationException("Specified IComparer failed.")
    End Select
End Function


' This is a set of equality routines used by function delegation calls.
' They allow a virtual equality routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
'
Public Function EqualsLong(ByRef x As Long, ByRef y As Long) As Boolean: EqualsLong = (x = y): End Function
Public Function EqualsString(ByRef x As String, ByRef y As String) As Boolean: EqualsString = (x = y): End Function
Public Function EqualsDouble(ByRef x As Double, ByRef y As Double) As Boolean: EqualsDouble = (x = y): End Function
Public Function EqualsInteger(ByRef x As Integer, ByRef y As Integer) As Boolean: EqualsInteger = (x = y): End Function
Public Function EqualsSingle(ByRef x As Single, ByRef y As Single) As Boolean: EqualsSingle = (x = y): End Function
Public Function EqualsDate(ByRef x As Date, ByRef y As Date) As Boolean: EqualsDate = (DateDiff("s", x, y) = 0): End Function
Public Function EqualsByte(ByRef x As Byte, ByRef y As Byte) As Boolean: EqualsByte = (x = y): End Function
Public Function EqualsBoolean(ByRef x As Boolean, ByRef y As Boolean) As Boolean: EqualsBoolean = (x = y): End Function
Public Function EqualsCurrency(ByRef x As Currency, ByRef y As Currency) As Boolean: EqualsCurrency = (x = y): End Function
Public Function EqualsObject(ByRef x As Object, ByRef y As Object) As Boolean
    If Not x Is Nothing Then
        If TypeOf x Is IObject Then
            Dim Obj As IObject
            Set Obj = x
            EqualsObject = Obj.Equals(y)
        Else
            EqualsObject = x Is y
        End If
    Else
        EqualsObject = y Is Nothing
    End If
End Function

Public Function EqualsVariants(ByRef x As Variant, ByRef y As Variant) As Boolean
    Dim o As IObject
    Select Case VarType(x)
        Case vbObject
            If x Is Nothing Then
                If IsObject(y) Then
                    EqualsVariants = (y Is Nothing)
                End If
            ElseIf TypeOf x Is IObject Then
                Set o = x
                EqualsVariants = o.Equals(y)
            ElseIf IsObject(y) Then
                If y Is Nothing Then Exit Function
                If TypeOf y Is IObject Then
                    Set o = y
                    EqualsVariants = o.Equals(x)
                Else
                    EqualsVariants = (x Is y)
                End If
            End If
        Case vbNull
            EqualsVariants = IsNull(y)
        Case VarType(y)
            EqualsVariants = (x = y)
    End Select
End Function

' This is a set of casting routines used by function delegation calls.
' They allow a virtual casting routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
'
Public Sub WidenLongToDouble(ByRef x As Double, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToSingle(ByRef x As Single, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToString(ByRef x As String, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToCurrency(ByRef x As Currency, ByRef y As Long): x = y: End Sub
Public Sub WidenLongToVariant(ByRef x As Variant, ByRef y As Long): x = y: End Sub
Public Sub WidenIntegerToLong(ByRef x As Long, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToSingle(ByRef x As Single, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToDouble(ByRef x As Double, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToString(ByRef x As String, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToCurrency(ByRef x As Currency, ByRef y As Integer): x = y: End Sub
Public Sub WidenIntegerToVariant(ByRef x As Variant, ByRef y As Integer): x = y: End Sub
Public Sub WidenByteToInteger(ByRef x As Integer, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToLong(ByRef x As Long, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToSingle(ByRef x As Single, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToDouble(ByRef x As Double, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToString(ByRef x As String, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToCurrency(ByRef x As Currency, ByRef y As Byte): x = y: End Sub
Public Sub WidenByteToVariant(ByRef x As Variant, ByRef y As Byte): x = y: End Sub
Public Sub WidenSingleToDouble(ByRef x As Double, ByRef y As Single): x = y: End Sub
Public Sub WidenSingleToString(ByRef x As String, ByRef y As Single): x = y: End Sub
Public Sub WidenSingleToVariant(ByRef x As Variant, ByRef y As Single): x = y: End Sub
Public Sub WidenDateToDouble(ByRef x As Double, ByRef y As Date): x = y: End Sub
Public Sub WidenDateToString(ByRef x As String, ByRef y As Date): x = y: End Sub
Public Sub WidenDateToVariant(ByRef x As Variant, ByRef y As Date): x = y: End Sub
Public Sub WidenObjectToVariant(ByRef x As Variant, ByRef y As Object): Set x = y: End Sub
Public Sub WidenCurrencyToString(ByRef x As String, ByRef y As Currency): x = y: End Sub
Public Sub WidenCurrencyToVariant(ByRef x As Variant, ByRef y As Currency): x = y: End Sub
Public Sub WidenCurrencyToDouble(ByRef x As Double, ByRef y As Currency): x = y: End Sub
Public Sub WidenStringToVariant(ByRef x As Variant, ByRef y As String): x = y: End Sub
Public Sub WidenDoubleToString(ByRef x As String, ByRef y As Double): x = y: End Sub
Public Sub WidenDoubleToVariant(ByRef x As Variant, ByRef y As Double): x = y: End Sub

' Functions used to assign variants to narrower variables.
Public Sub NarrowVariantToLong(ByRef x As Long, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToInteger(ByRef x As Integer, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToDouble(ByRef x As Double, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToString(ByRef x As String, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToSingle(ByRef x As Single, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToByte(ByRef x As Byte, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToDate(ByRef x As Date, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToBoolean(ByRef x As Boolean, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToCurrency(ByRef x As Currency, ByRef y As Variant): x = y: End Sub
Public Sub NarrowVariantToObject(ByRef x As Object, ByRef y As Variant): Set x = y: End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Optimized sort routines. There could have been one
'   all-purpose sort routine, but it would be too slow.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSortKeys(ByVal pSA As Long)
    CopyMemory mSortKeys.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
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
    CopyMemory mSortItems.SA, ByVal pSA, SIZEOF_SAFEARRAY1D
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

Public Sub QuickSortLong(ByRef Keys() As Long, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortString(ByRef Keys() As String, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortInteger(ByRef Keys() As Integer, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortByte(ByRef Keys() As Byte, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortDouble(ByRef Keys() As Double, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortSingle(ByRef Keys() As Single, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortCurrency(ByRef Keys() As Currency, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortBoolean(ByRef Keys() As Boolean, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortVariant(ByRef Keys() As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd x, Keys((i + j) \ 2)
        Do
            Do While CompareVariants(Keys(i), x) < 0: i = i + 1: Loop
            Do While CompareVariants(Keys(j), x) > 0: j = j - 1: Loop
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


