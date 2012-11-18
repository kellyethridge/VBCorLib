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
    Const BYREF_ARRAY As Long = VT_BYREF Or vbArray
    
    Dim vt As Long
    
    vt = VariantType(Arr)
    Select Case vt And BYREF_ARRAY
        ' we have to double deref the original array pointer because
        ' the variant held a pointer to the original array variable.
        Case BYREF_ARRAY
            GetArrayPointer = MemLong(MemLong(VarPtr(Arr) + VARIANTDATA_OFFSET))
            
        ' we won't need to deref again if the original array was dimensioned
        ' as a variant ie:
        '    Dim arr As Variant
        '    ReDim arr(1 To 10) As Long
        '
        ' The passed in variant will be the array variable, not a ByRef
        ' pointer to the array variable.
        Case vbArray
            GetArrayPointer = MemLong(VarPtr(Arr) + VARIANTDATA_OFFSET)
        
        Case Else
            Throw Cor.NewArgumentException(Resources.GetString(Argument_ArrayRequired), "Arr")
    End Select
    
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
    Select Case vt And &HFF
        Case vbObject, vbUserDefinedType
            If UBound(Arr) < LBound(Arr) Then
                GetArrayPointer = vbNullPtr
            End If
    End Select
    
    If ThrowOnNull Then
        If GetArrayPointer = vbNullPtr Then
            Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
        End If
    End If
End Function


' This is a set of comparison routines used by function delegation calls.
' They allow a virtual comparison routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
Public Function CompareLongs(ByRef X As Long, ByRef Y As Long) As Long
    If X > Y Then
        CompareLongs = 1
    ElseIf X < Y Then
        CompareLongs = -1
    End If
End Function
Public Function CompareIntegers(ByRef X As Integer, ByRef Y As Integer) As Long
    If X > Y Then
        CompareIntegers = 1
    ElseIf X < Y Then
        CompareIntegers = -1
    End If
End Function
Public Function CompareStrings(ByRef X As String, ByRef Y As String) As Long
    If X > Y Then
        CompareStrings = 1
    ElseIf X < Y Then
        CompareStrings = -1
    End If
End Function
Public Function CompareDoubles(ByRef X As Double, ByRef Y As Double) As Long
    If X > Y Then
        CompareDoubles = 1
    ElseIf X < Y Then
        CompareDoubles = -1
    End If
End Function
Public Function CompareSingles(ByRef X As Single, ByRef Y As Single) As Long
    If X > Y Then
        CompareSingles = 1
    ElseIf X < Y Then
        CompareSingles = -1
    End If
End Function
Public Function CompareBytes(ByRef X As Byte, ByRef Y As Byte) As Long
    If X > Y Then
        CompareBytes = 1
    ElseIf X < Y Then
        CompareBytes = -1
    End If
End Function
Public Function CompareBooleans(ByRef X As Boolean, ByRef Y As Boolean) As Long
    If X > Y Then
        CompareBooleans = 1
    ElseIf X < Y Then
        CompareBooleans = -1
    End If
End Function
Public Function CompareDates(ByRef X As Date, ByRef Y As Date) As Long
    CompareDates = DateDiff("s", Y, X)
End Function
Public Function CompareCurrencies(ByRef X As Currency, ByRef Y As Currency) As Long
    If X > Y Then CompareCurrencies = 1: Exit Function
    If X < Y Then CompareCurrencies = -1
End Function
Public Function CompareIComparable(ByRef X As Object, ByRef Y As Variant) As Long
    Dim comparableX As IComparable
    Set comparableX = X
    CompareIComparable = comparableX.CompareTo(Y)
End Function
Public Function CompareVariants(ByRef X As Variant, ByRef Y As Variant) As Long
    Dim Comparable As IComparable
    
    Select Case VarType(X)
        Case vbNull
            If Not IsNull(Y) Then
                CompareVariants = -1
            End If
            Exit Function
        Case vbEmpty
            If Not IsEmpty(Y) Then
                CompareVariants = -1
            End If
            Exit Function
        Case vbObject, vbDataObject
            If TypeOf X Is IComparable Then
                Set Comparable = X
                CompareVariants = Comparable.CompareTo(Y)
                Exit Function
            End If
        Case VarType(Y)
            If X < Y Then
                CompareVariants = -1
            ElseIf X > Y Then
                CompareVariants = 1
            End If
            Exit Function
    End Select
    
    Select Case VarType(Y)
        Case vbNull, vbEmpty
            CompareVariants = 1
        Case vbObject, vbDataObject
            If TypeOf Y Is IComparable Then
                Set Comparable = Y
                CompareVariants = -Comparable.CompareTo(X)
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
Public Function EqualsLong(ByRef X As Long, ByRef Y As Long) As Boolean: EqualsLong = (X = Y): End Function
Public Function EqualsString(ByRef X As String, ByRef Y As String) As Boolean: EqualsString = (X = Y): End Function
Public Function EqualsDouble(ByRef X As Double, ByRef Y As Double) As Boolean: EqualsDouble = (X = Y): End Function
Public Function EqualsInteger(ByRef X As Integer, ByRef Y As Integer) As Boolean: EqualsInteger = (X = Y): End Function
Public Function EqualsSingle(ByRef X As Single, ByRef Y As Single) As Boolean: EqualsSingle = (X = Y): End Function
Public Function EqualsDate(ByRef X As Date, ByRef Y As Date) As Boolean: EqualsDate = (DateDiff("s", X, Y) = 0): End Function
Public Function EqualsByte(ByRef X As Byte, ByRef Y As Byte) As Boolean: EqualsByte = (X = Y): End Function
Public Function EqualsBoolean(ByRef X As Boolean, ByRef Y As Boolean) As Boolean: EqualsBoolean = (X = Y): End Function
Public Function EqualsCurrency(ByRef X As Currency, ByRef Y As Currency) As Boolean: EqualsCurrency = (X = Y): End Function
Public Function EqualsObject(ByRef X As Object, ByRef Y As Object) As Boolean
    If Not X Is Nothing Then
        If TypeOf X Is IObject Then
            Dim Obj As IObject
            Set Obj = X
            EqualsObject = Obj.Equals(Y)
        Else
            EqualsObject = X Is Y
        End If
    Else
        EqualsObject = Y Is Nothing
    End If
End Function

Public Function EqualsVariants(ByRef X As Variant, ByRef Y As Variant) As Boolean
    Dim o As IObject
    Select Case VarType(X)
        Case vbObject
            If X Is Nothing Then
                If IsObject(Y) Then
                    EqualsVariants = (Y Is Nothing)
                End If
            ElseIf TypeOf X Is IObject Then
                Set o = X
                EqualsVariants = o.Equals(Y)
            ElseIf IsObject(Y) Then
                If Y Is Nothing Then Exit Function
                If TypeOf Y Is IObject Then
                    Set o = Y
                    EqualsVariants = o.Equals(X)
                Else
                    EqualsVariants = (X Is Y)
                End If
            End If
        Case vbNull
            EqualsVariants = IsNull(Y)
        Case VarType(Y)
            EqualsVariants = (X = Y)
    End Select
End Function

' This is a set of casting routines used by function delegation calls.
' They allow a virtual casting routine to be selected and called without
' needing to modify code. Only the address of the specific routine is needed.
'
Public Sub WidenLongToDouble(ByRef X As Double, ByRef Y As Long): X = Y: End Sub
Public Sub WidenLongToSingle(ByRef X As Single, ByRef Y As Long): X = Y: End Sub
Public Sub WidenLongToString(ByRef X As String, ByRef Y As Long): X = Y: End Sub
Public Sub WidenLongToCurrency(ByRef X As Currency, ByRef Y As Long): X = Y: End Sub
Public Sub WidenLongToVariant(ByRef X As Variant, ByRef Y As Long): X = Y: End Sub
Public Sub WidenIntegerToLong(ByRef X As Long, ByRef Y As Integer): X = Y: End Sub
Public Sub WidenIntegerToSingle(ByRef X As Single, ByRef Y As Integer): X = Y: End Sub
Public Sub WidenIntegerToDouble(ByRef X As Double, ByRef Y As Integer): X = Y: End Sub
Public Sub WidenIntegerToString(ByRef X As String, ByRef Y As Integer): X = Y: End Sub
Public Sub WidenIntegerToCurrency(ByRef X As Currency, ByRef Y As Integer): X = Y: End Sub
Public Sub WidenIntegerToVariant(ByRef X As Variant, ByRef Y As Integer): X = Y: End Sub
Public Sub WidenByteToInteger(ByRef X As Integer, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenByteToLong(ByRef X As Long, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenByteToSingle(ByRef X As Single, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenByteToDouble(ByRef X As Double, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenByteToString(ByRef X As String, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenByteToCurrency(ByRef X As Currency, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenByteToVariant(ByRef X As Variant, ByRef Y As Byte): X = Y: End Sub
Public Sub WidenSingleToDouble(ByRef X As Double, ByRef Y As Single): X = Y: End Sub
Public Sub WidenSingleToString(ByRef X As String, ByRef Y As Single): X = Y: End Sub
Public Sub WidenSingleToVariant(ByRef X As Variant, ByRef Y As Single): X = Y: End Sub
Public Sub WidenDateToDouble(ByRef X As Double, ByRef Y As Date): X = Y: End Sub
Public Sub WidenDateToString(ByRef X As String, ByRef Y As Date): X = Y: End Sub
Public Sub WidenDateToVariant(ByRef X As Variant, ByRef Y As Date): X = Y: End Sub
Public Sub WidenObjectToVariant(ByRef X As Variant, ByRef Y As Object): Set X = Y: End Sub
Public Sub WidenCurrencyToString(ByRef X As String, ByRef Y As Currency): X = Y: End Sub
Public Sub WidenCurrencyToVariant(ByRef X As Variant, ByRef Y As Currency): X = Y: End Sub
Public Sub WidenCurrencyToDouble(ByRef X As Double, ByRef Y As Currency): X = Y: End Sub
Public Sub WidenStringToVariant(ByRef X As Variant, ByRef Y As String): X = Y: End Sub
Public Sub WidenDoubleToString(ByRef X As String, ByRef Y As Double): X = Y: End Sub
Public Sub WidenDoubleToVariant(ByRef X As Variant, ByRef Y As Double): X = Y: End Sub

' Functions used to assign variants to narrower variables.
Public Sub NarrowVariantToLong(ByRef X As Long, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToInteger(ByRef X As Integer, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToDouble(ByRef X As Double, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToString(ByRef X As String, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToSingle(ByRef X As Single, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToByte(ByRef X As Byte, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToDate(ByRef X As Date, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToBoolean(ByRef X As Boolean, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToCurrency(ByRef X As Currency, ByRef Y As Variant): X = Y: End Sub
Public Sub NarrowVariantToObject(ByRef X As Object, ByRef Y As Variant): Set X = Y: End Sub



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
    Dim i As Long, j As Long, X As Long, t As Long
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As String
    Do While Left < Right
        i = Left: j = Right: X = StringRef(Keys((i + j) \ 2))
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
        StringPtr(X) = 0
    Loop
End Sub

Public Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Variant, Key As IComparable
    Do While Left < Right
        i = Left: j = Right: Set X = Keys((i + j) \ 2)
        Do
            Set Key = Keys(i): Do While Key.CompareTo(X) < 0: i = i + 1: Set Key = Keys(i): Loop
            Set Key = Keys(j): Do While Key.CompareTo(X) > 0: j = j - 1: Set Key = Keys(j): Loop
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
    Dim i As Long, j As Long, X As Integer, t As Integer
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Byte, t As Byte
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Double, t As Double
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Single, t As Single
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Currency, t As Currency
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Boolean, t As Boolean
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd X, Keys((i + j) \ 2)
        Do
            Do While CompareVariants(Keys(i), X) < 0: i = i + 1: Loop
            Do While CompareVariants(Keys(j), X) > 0: j = j - 1: Loop
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
    Dim i As Long, j As Long, X As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd X, Keys((i + j) \ 2)
        Do
            Do While SortComparer.Compare(Keys(i), X) < 0: i = i + 1: Loop
            Do While SortComparer.Compare(Keys(j), X) > 0: j = j - 1: Loop
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


