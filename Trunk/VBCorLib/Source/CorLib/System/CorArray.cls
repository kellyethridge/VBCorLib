VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorArray"
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
' Module: CorArray
'

''
' Provides methods for manipulating, searching and sorting arrays.
'
' @remarks This class cannot be directly instantiated. To access the methods
' use the class name directly.
' <pre>
' Debug.Print CorArray.GetLength(MyArray)
' </pre>
'
Option Explicit


' extended safearray descriptor the includes the 16 bytes
' preceding the array descriptor. The 16 bytes contains
' a GUID for objects and user-defined types. This is to
' help in the correct copying of those types of elements.
Private Type GuidSafeArray1d
    Guid    As VBGUID
    SA      As SafeArray1d
End Type

Private mSortDelegator          As FunctionDelegator
Private mSorter                 As ISortRoutine

Private CallbackSorter          As New CallbackSorter


Private Sub Class_Initialize()
    Set mSorter = InitDelegator(mSortDelegator)
End Sub

''
' Uses a callback method to check if an element in an array exists.
'
' @param arr The one-dimensional array to check if an element exists according to the callback criteria.
' @param AddressOfPredicate The AddressOf the callback method used to determine if
' an element matches the specified criteria.
' @return Indication of the element existing in the array.
' @remarks Each array element is passed into a user callback method. The user
' then checks the element value to determine if it matches any criteria necessary
' and returns the result. The criteria is defined outside of the
' <b>Exists</b> method. How the criteria is defined and checked is up to the
' user of the function.
' <p>The function signature for the callback has a specific format.
' <pre>
' Public Function ExistsCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
Public Function Exists(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Boolean
    Exists = (FindIndex(Arr, AddressOfPredicate) >= LBound(Arr))
End Function

''
' Uses a callback method to search an array for the first element that matches the criteria.
'
' @param arr The one-dimensional array to search.
' @param AddressOfPredicate The callback method address used to determine if an element
' matches the specified criteria.
' @return The element value if a match is found, otherwise a default for that element type.
' @remarks <p>The function signature for the callback has a specific format.
' <pre>
' Public Function FindCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray.Find.txt"
Public Function Find(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Variant
    Dim Index As Long
    
    Index = FindIndex(Arr, AddressOfPredicate)
    If Index >= LBound(Arr) Then
        VariantCopyInd Find, Arr(Index)
    Else
        AssignDefaultElementValue Arr, Find
    End If
End Function

''
' Finds the index of the first array element that matches the criteria
' specified by the callback function.
'
' @param arr The one-dimensional array to be searched.
' @param AddressOfPredicate The address of the callback function used to match the array elements.
' @param Index The starting element of the array to beging the matching process.
' @param Count The number of elements to include in the array search.
' @return The index of the matched element, or LowerBound - 1 if none is found.
' @remarks <p>The function signature for the callback has a specific format.
' <pre>
' Public Function FindIndexCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray.Find.txt"
Public Function FindIndex(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
        
    Dim ElementIndex   As Long
    Dim ElementCount   As Long
    ElementIndex = OptionalLong(Index, LBound(Arr))
    ElementCount = OptionalLong(Count, UBound(Arr) - ElementIndex + 1)
    
    If ElementIndex < LBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If ElementCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    If ElementIndex + ElementCount - 1 > UBound(Arr) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_UBound
    
    Dim ElementSize As Long
    ElementSize = SafeArrayGetElemsize(Ptr)
    
    Dim ElementPtr As Long
    ElementPtr = MemLong(Ptr + PVDATA_OFFSET) + (ElementIndex - LBound(Arr)) * ElementSize
    
    Dim Predicate As Func_T_Boolean
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    Dim i As Long
    For i = 1 To ElementCount
        If Predicate.Invoke(ByVal ElementPtr) Then
            FindIndex = ElementIndex + i - 1
            Exit Function
        End If
        ElementPtr = ElementPtr + ElementSize
    Next
    
    FindIndex = LBound(Arr) - 1
End Function

''
' Finds all matching elements in an array, returning an array of the matched elements.
'
' @param arr The one-dimensional array to find all matching elements.
' @param AddressOfPredicate The callback address of the matching criteria method.
' @return An array of all matched elements.
' @remarks <p>The function signature for the callback has a specific format.
' <pre>
' Public Function FindAllCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray.Find.txt"
Public Function FindAll(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Variant
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
    
    Dim LowerBound As Long
    LowerBound = LBound(Arr)
    
    Dim Count As Long
    Count = UBound(Arr) - LowerBound + 1
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(Ptr)
    
    Dim pElem As Long
    pElem = MemLong(Ptr + PVDATA_OFFSET)
    
    Dim Predicate As Func_T_Boolean
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    ' iterate each array element, passing the memory address
    ' to the user supplied AddressOf callback. This will appear
    ' as a ByRef parameter of the array's datatype in the callback.
    Dim i                   As Long
    Dim MatchCount          As Long
    Dim MatchedIndexes()    As Long
    
    ReDim MatchedIndexes(0 To 31)
    For i = LowerBound To LowerBound + Count - 1
        If Predicate.Invoke(ByVal pElem) Then
            If MatchCount > UBound(MatchedIndexes) Then
                ReDim Preserve MatchedIndexes(0 To MatchCount * 2 - 1)
            End If
            MatchedIndexes(MatchCount) = i
            MatchCount = MatchCount + 1
        End If
        pElem = pElem + ElemSize
    Next i
    
    Dim Ret         As Variant
    Dim ArrayType   As Long
    
    ArrayType = VarType(Arr) And &HFF   ' knock off the vbArray bit.
    Select Case ArrayType
        Case vbObject
            ' as an array of objects, we create a new array using the
            ' same Guid value as that of the original array.
            VariantType(Ret) = vbObject Or vbArray
            MemLong(VarPtr(Ret) + VARIANTDATA_OFFSET) = SafeArrayCreateVectorEx(vbObject, 0, MatchCount, Ptr - SizeOfGuid)
            
            For i = 0 To MatchCount - 1
                Set Ret(i) = Arr(MatchedIndexes(i))
            Next i
            
        Case vbUserDefinedType
            Dim UDT As IRecordInfo
            Set UDT = SafeArrayGetRecordInfo(Ptr)    ' gets an object used to manipulate a UDT.
            
            Dim pvData  As Long
            pvData = MemLong(Ptr + PVDATA_OFFSET)
            
            Dim pReturnSafeArray As Long
            pReturnSafeArray = SafeArrayCreateVectorEx(vbUserDefinedType, 0, MatchCount, ObjPtr(UDT))
            
            Dim pDstElem As Long
            pDstElem = MemLong(pReturnSafeArray + PVDATA_OFFSET)
            
            For i = 0 To MatchCount - 1
                UDT.RecordCopy pvData + (MatchedIndexes(i) - LowerBound) * ElemSize, pDstElem
                pDstElem = pDstElem + ElemSize
            Next i
            
            VariantType(Ret) = vbUserDefinedType Or vbArray
            MemLong(VarPtr(Ret) + VARIANTDATA_OFFSET) = pReturnSafeArray
            
        Case Else
            Ret = CreateInstance(ArrayType, MatchCount)
            For i = 0 To MatchCount - 1
                Ret(i) = Arr(MatchedIndexes(i))
            Next i
    End Select
    
    FindAll = Ret
End Function

''
' Finds the last occurence of a matched element in the array.
'
' @param arr The array to search.
' @param AddressOfPredicate A callback address of the matching criteria method.
' @return The last element in the array to match the criteria, or the
' default for the array type if no elements matched.
' @remarks <p>The function signature for the callback has a specific format.
' <pre>
' Public Function FindLastCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray.Find.txt"
Public Function FindLast(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Variant
    Dim Index As Long
    
    Index = FindLastIndex(Arr, AddressOfPredicate)
    If Index >= LBound(Arr) Then
        VariantCopyInd FindLast, Arr(Index)
    Else
        AssignDefaultElementValue Arr, FindLast
    End If
End Function

''
' Finds the index of the last occurence of a matched element in the array.
'
' @param arr The one-dimensional array to search.
' @param AddressOfPredicate The callback address of the matching criteria method.
' @param Index The starting index in the array to begin the backward search.
' @param Count The number of elements to search.
' @return The index of the last occurence of the matched element.
' @remarks <p>The function signature for the callback has a specific format.
' <pre>
' Public Function FindLastIndexCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray.Find.txt"
Public Function FindLastIndex(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr

    Dim ElemIndex As Long
    Dim ElemCount As Long
    ElemIndex = OptionalLong(Index, UBound(Arr))
    ElemCount = OptionalLong(Count, ElemIndex + 1)
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayBounds
    If ElemCount < 0 Then _
        Error.NegativeNumber "Count"
    If ElemIndex - ElemCount + 1 < LBound(Arr) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_LBound
        
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(Ptr)
    
    Dim pElem As Long
    pElem = MemLong(Ptr + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    
    Dim Predicate As Func_T_Boolean
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    Dim i As Long
    For i = ElemCount To 1 Step -1
        If Predicate.Invoke(ByVal pElem) Then
            FindLastIndex = LBound(Arr) + i - 1
            Exit Function
        End If
        pElem = pElem - ElemSize
    Next i
    FindLastIndex = LBound(Arr) - 1
End Function

''
' Returns a boolean indicating if all elements in the array matched the criteria.
'
' @param arr The array to compare against the criteria.
' @param AddressOfPredicate The callback address used to match the criteria.
' @Return Indication of the entire array matching the criteria or not.
' @remarks <p>The function signature for the callback has a specific format.
' <pre>
' Public Function TrueForAllCallback(ByRef x As [Array Datatype]) As Boolean
'     '' return True if x matches your criteria.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
Public Function TrueForAll(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Boolean
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(Ptr)
    
    Dim pElem As Long
    pElem = MemLong(Ptr + PVDATA_OFFSET)
    
    Dim Predicate As Func_T_Boolean
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If Not Predicate.Invoke(ByVal pElem) Then
            Exit Function
        End If
        
        pElem = pElem + ElemSize
    Next i
    TrueForAll = True
End Function

''
' Iterates over an array passing in each element into an Action function to be performed on the element.
'
' @param Arr The one-dimensional array containing the elements to be processed.
' @param AddressOfAction The callback address of the function to process an array element.
' @remarks <p>The callback method must be a Sub with a single ByRef parameter of the same
' type as the array.
' <pre>
' Public Sub ActionMethod(ByRef e As [Data Type])
'     '' Perform action on e.
' End Sub
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then the [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
Public Sub ForEach(ByRef Arr As Variant, ByVal AddressOfAction As Long)
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(Ptr)
    
    Dim pElem As Long
    pElem = MemLong(Ptr + PVDATA_OFFSET)
    
    Dim Action As Action_T
    Set Action = NewDelegate(AddressOfAction)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Action.Invoke ByVal pElem
        pElem = pElem + ElemSize
    Next i
End Sub

''
' Performs a binary search on a given array. A subportion of the array can
' be searched using the startindex and length parameters. A custom user
' comparer can optionally be supplied to perform special comparisons
' between elements in the array.
'
' @param Arr The array to search for a specific value.
' @param Value The value to search for in the array.
' @param Comparer A user supplied object to compare elements within the array.
' @return The index at which the value was found.
' @remarks If the returned value is less than the lower bound of the array, then
' the value indicates where the value would have been found in the array. The
' following is how to convert the return value to the array index:<br><br>
' Lowerbound of 0:  return = Not return<br>
' Other lowerbound: return = (Not (return - lowerbound)) + lowerbound<br>
' @include "..\Includes\CorArray_BinarySearch.txt"
Public Function BinarySearch(ByRef Arr As Variant, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    CheckValidSingleDimArray Arr
    BinarySearch = BinarySearchCore(Arr, LBound(Arr), CorArray.LengthFirstDim(Arr), Value, Comparer)
End Function

''
' Performs a binary search on a given array. A subportion of the array can
' be searched using the startindex and length parameters. A custom user
' comparer can optionally be supplied to perform special comparisons
' between elements in the array.
'
' @param Arr The array to search for a specific value.
' @param Index The starting index in the array to begin searching.
' @param Length The number of elements to search, starting at startindex.
' @param Value The value to search for in the array.
' @param Comparer A user supplied object to compare elements within the array.
' @return The index at which the value was found.
' @remarks If the returned value is less than the lower bound of the array, then
' the value indicates where the value would have been found in the array. The
' following is how to convert the return value to the array index:<br><br>
' Lowerbound of 0:  return = Not return<br>
' Other lowerbound: return = (Not (return - lowerbound)) + lowerbound<br>
' @include "..\Includes\CorArray_BinarySearch.txt"
Public Function BinarySearchEx(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    CheckValidSingleDimArray Arr
    If Index < LBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Length < 0 Then _
        Error.NegativeNumber "Length"
    If Index + Length - 1 > UBound(Arr) Then _
        Error.InvalidOffsetLength
    
    BinarySearchEx = BinarySearchCore(Arr, Index, Length, Value, Comparer)
End Function

''
' Clears a portion of the elements in an array.
'
' @param arr The array to clear elements from.
' @param Index The starting element to being clearing.
' @param Count The number of elements to be cleared.
'
Public Sub Clear(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long)
    Dim pSafeArray  As Long
    
    pSafeArray = ArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    If Index < LBound(Arr) Then _
        Error.IndexOutOfRange IndexOutOfRange_ArrayBounds
    If Length < 0 Then _
        Error.IndexOutOfRange IndexOutOfRange_ArrayBounds
    If Index + Length - 1 > UBound(Arr) Then _
        Error.IndexOutOfRange IndexOutOfRange_ArrayBounds
    
    If Length = 0 Then Exit Sub
    
    Dim IIDOffset As Long
    Select Case VarType(Arr) And &HFF   ' knock off the vbArray bit
        Case vbObject
            IIDOffset = SizeOfGuid
            
        Case vbString, vbUserDefinedType
            IIDOffset = 4
            
        Case Else
            ' we can use ZeroMemory for everything else.
            Dim ElemSize As Long
            ElemSize = SafeArrayGetElemsize(pSafeArray)
            
            Dim pElem As Long
            pElem = MemLong(pSafeArray + PVDATA_OFFSET)
            
            ' optimize by zeroing out the memory directly.
            ZeroMemory ByVal pElem + (Index - LBound(Arr)) * ElemSize, Length * ElemSize
            Exit Sub
    End Select
    
    Dim SubArray As GuidSafeArray1d
    CopyMemory ByVal VarPtr(SubArray.SA) - IIDOffset, ByVal pSafeArray - IIDOffset, Len(SubArray.SA) + IIDOffset
    
    With SubArray.SA
        .fFeatures = .fFeatures Or FADF_FIXEDSIZE Or FADF_STATIC Or FADF_AUTO
        .pvData = .pvData + (Index - LBound(Arr)) * .cbElements
        .cElements = Length
    End With
    SafeArrayDestroyData VarPtr(SubArray.SA)
End Sub

''
' Copies one Array to another Array and performs type casting as necessary.
'
' @param SourceArray The array from which to copy the elements.
' @param DestinationArray The array in which to place the elements.
' @param Count The number of elements to copy.
' @remarks The AddressOfCopier method signature must be two parameters
' of ByRef for the datatypes of each of the arrays. The destination
' array element is the first parameter.
' <h4>Example</h4> This is an example of a callback method declaration.
' <pre>
' Public Sub CopyCallback(ByRef dst As String, ByRef src As MyClass)
'     dst = src.Name
' End Sub
' </pre>
Public Sub Copy(ByRef SourceArray As Variant, ByRef DestinationArray As Variant, ByVal Length As Long)
    If IsNull(SourceArray) Then _
        Error.ArgumentNull "SourceArray", ArgumentNull_Array
    If IsNull(DestinationArray) Then _
        Error.ArgumentNull "DestinationArray", ArgumentNull_Array
    
    CopyEx SourceArray, LBound(SourceArray), DestinationArray, LBound(DestinationArray), Length
End Sub

''
' Copies a section of one Array to another Array and performs type casting as necessary.
'
' @param sourcearray The array from which to copy the elements.
' @param sourceindex The starting element in the source array to begin copying from.
' @param destinationarray The array in which to place the elements.
' @param destinationindex The starting element in the destination array to place elements.
' @param Count The number of elements to copy.
' @param AddressOfCopier A callback address to a method used to assign
' elements from the source array to the destination array.
' @remarks The AddressOfCopier method signature must be two parameters
' of ByRef for the datatypes of each of the arrays. The destination
' array element is the first parameter.
' <h4>Example</h4> This is an example of a callback method declaration.
' <pre>
' Public Sub CopyCallback(ByRef dst As String, ByRef src As MyClass)
'     dst = src.Name
' End Sub
' </pre>
Public Sub CopyEx(ByRef SourceArray As Variant, ByVal SourceIndex As Long, ByRef DestinationArray As Variant, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim pSourceSafeArray As Long
    pSourceSafeArray = ArrayPointer(SourceArray)
    If pSourceSafeArray = vbNullPtr Then _
        Error.ArgumentNull "SourceArray", ArgumentNull_Array
    
    Dim pDestinationSafeArray As Long
    pDestinationSafeArray = ArrayPointer(DestinationArray)
    If pDestinationSafeArray = vbNullPtr Then _
        Error.ArgumentNull "DestinationArray", ArgumentNull_Array
        
    Dim SourceLength As Long
    SourceLength = LengthCore(pSourceSafeArray, 0)
    Dim DestinationLength As Long
    DestinationLength = LengthCore(pDestinationSafeArray, 0)
    
    If Length < 0 Then _
        Error.NegativeNumber "Length"
    If Length + SourceIndex - LBound(SourceArray) > SourceLength Then _
        Error.Argument Argument_LongerThanSrCorArray
    If Length + DestinationIndex - LBound(DestinationArray) > DestinationLength Then _
        Error.Argument Argument_LongerThanDestArray
    If SourceIndex < LBound(SourceArray) Then _
        Error.ArgumentOutOfRange "SourceIndex", ArgumentOutOfRange_ArrayLB
    If DestinationIndex < LBound(DestinationArray) Then _
        Error.ArgumentOutOfRange "DestinationIndex", ArgumentOutOfRange_ArrayLB
    
    Dim SourceArrayType As VbVarType
    SourceArrayType = VarType(SourceArray) And &HFF    ' knock off the vbArray bit.
    
    Dim DestinationArrayType As VbVarType
    DestinationArrayType = VarType(DestinationArray) And &HFF   ' knock off the vbArray bit.
    
    If SourceArrayType = DestinationArrayType Then
        Select Case SourceArrayType
            Case vbObject, vbDataObject
                If IsEqualGUID(ByVal pSourceSafeArray - SizeOfGuid, ByVal pDestinationSafeArray - SizeOfGuid) <> BOOL_FALSE Then
                    CopyReferenceTypes pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
                Else
                    CopyObjectTypesCast pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
                End If
                
            Case vbString, vbVariant
                CopyReferenceTypes pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
                
            Case vbUserDefinedType
                Dim SrcUDT As IRecordInfo
                Dim DstUDT As IRecordInfo
                
                Set SrcUDT = SafeArrayGetRecordInfo(pSourceSafeArray)
                Set DstUDT = SafeArrayGetRecordInfo(pDestinationSafeArray)
                
                If SrcUDT.IsMatchingType(DstUDT) = BOOL_FALSE Then _
                    Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
                    
                CopyReferenceTypes pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
                
            Case Else
                CopyValueTypes pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
        End Select
    Else
        Dim Assigner As Long
        Assigner = GetAssigningFunction(SourceArrayType, DestinationArrayType)
        If Assigner = vbNullPtr Then _
            Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
        
        CopyWithCast pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length, Assigner
    End If
End Sub

''
' Initializes a new array of the specified type in up to 3 dimensions.
'
' @param ElementType The data type the array represents.
' @param length1 The number of elements in the first dimension.
' @param length2 The number of elements in the second dimension.
' @param length3 The number of elements in the third dimension.
' @returns An array of the specified type with the specified dimensions.
' @remarks If a length is zero, then an empty array is returned, not a Null array.
' If one length is defined as zero, then all other defined lengths must be zero.
' This does not mean that lenghts 2 and 3 need to be set to zero if only an empty
' one-dimensional array is being created.
' <p>Elements of a size that are a multiple of 4 will not create a new array
' when being returned. All other datatype sizes will cause a new array
' to be created. The array must also be assigned at the end of the code
' so VB can optimize returning the Variant datatype and not copy it when possible.</p>
'
Public Function CreateInstance(ByVal ElementType As VbVarType, Optional ByVal Length1 As Long = 0, Optional ByRef Length2 As Variant, Optional ByRef Length3 As Variant) As Variant
    Dim Bounds(2)   As SafeArrayBound
    Dim Rank        As Long
    
    Select Case ElementType
        Case vbByte, vbInteger, vbLong, vbDouble, vbSingle, vbBoolean, vbCurrency, vbDate, vbString, vbObject, vbVariant
        Case Else
            Error.Argument Argument_MustBeVbVarType, "ElementType"
    End Select
    
    If Length1 < 0 Then _
        Error.NegativeNumber "Length1"
    
    Rank = 1
    Bounds(0).cElements = Length1
    
    If Not IsMissing(Length2) Then
        If CLng(Length2) < 0 Then _
            Error.NegativeNumber "Length2"
        
        Rank = 2
        Bounds(1).cElements = Length2
    
        If Not IsMissing(Length3) Then
            If CLng(Length3) < 0 Then _
                Error.NegativeNumber "Length3"
            
            Rank = 3
            Bounds(2).cElements = Length3
        End If
    End If
        
    VariantType(CreateInstance) = ElementType Or vbArray
    MemLong(VarPtr(CreateInstance) + VARIANTDATA_OFFSET) = SafeArrayCreate(ElementType, Rank, Bounds(0))
End Function

''
' Returns the number of elements in the specified dimension. If 0 is
' specified, this returns the total number of elements in the array.
'
' @param arr The array to retrieve the number of elements from.
' @param dimension (Optional) the specific dimension to get the number of elements.
' @return The number of elements in the specified dimension or the entire array.
'
Public Function GetLength(ByRef Arr As Variant, Optional ByVal Dimension As Long) As Long
    Dim SafeArrayPtr  As Long
    SafeArrayPtr = ArrayPointer(Arr)
    If SafeArrayPtr = vbNullPtr Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    GetLength = LengthCore(SafeArrayPtr, Dimension)
End Function

Friend Function LengthCore(ByVal SafeArrayPtr As Long, Optional ByVal Dimension As Long) As Long
    Debug.Assert SafeArrayPtr <> 0
    
    Dim Rank As Long
    Rank = SafeArrayGetDim(SafeArrayPtr)
    If Dimension > Rank Then _
        Error.IndexOutOfRange IndexOutOfRange_Dimension
    
    Dim ArrayBoundPtr As Long
    ArrayBoundPtr = SafeArrayPtr + SizeOfSafeArray
    
    If Dimension < 1 Then
        ' optimize for 99% of calls
        If Rank = 1 Then
            LengthCore = MemLong(ArrayBoundPtr)
        Else
            LengthCore = CalculateProductOfDimensions(ArrayBoundPtr, Rank)
        End If
    Else
        LengthCore = GetDimensionSize(ArrayBoundPtr, Rank, Dimension)
    End If
End Function

Private Function CalculateProductOfDimensions(ByVal ArrayBoundPtr As Long, ByVal Rank As Long) As Long
    Dim Result As Long
    
    Result = 1
    Dim Ptr As Long
    For Ptr = ArrayBoundPtr To ArrayBoundPtr + (SizeOfSafeArrayBound * (Rank - 1)) Step SizeOfSafeArrayBound
        Result = Result * MemLong(Ptr)
    Next

    CalculateProductOfDimensions = Result
End Function

Private Function GetDimensionSize(ByVal ArrayBoundPtr As Long, ByVal Rank As Long, ByVal Dimension As Long) As Long
    ArrayBoundPtr = ArrayBoundPtr + (Rank - Dimension) * SizeOfSafeArrayBound
    GetDimensionSize = MemLong(ArrayBoundPtr)
End Function

Friend Function LengthFirstDim(ByRef Arr As Variant) As Long
    LengthFirstDim = UBound(Arr) - LBound(Arr) + 1
End Function

''
' Returns the number of dimensions in the array.
'
' @param Arr The array to retrieve the number of dimensions from.
' @return The number of dimension in the array.
Public Function Rank(ByRef Arr As Variant) As Long
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    If Ptr = vbNullPtr Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    Rank = SafeArrayGetDim(Ptr)
End Function

''
' Returns the index of the first occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param Arr The one-dimensional array to search.
' @param Value The value to search for in the array.
' @param StartIndex The index to the element in the array to begin search from.
' @param Count The number of elements to search, starting from startindex.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function IndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
    
    Dim Range As ListRange
    Range = OptionalStepRange(StartIndex, Count, LBound(Arr), LengthFirstDim(Arr), Parameter_StartIndex)
    
    If Range.Index < LBound(Arr) Or Range.Index > UBound(Arr) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_ArrayBounds
    If Range.Count < 0 Or Range.Index + Range.Count - 1 > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
        
    On Error GoTo errTrap
    If TrySZIndexOf(Ptr, Value, Range.Index, Range.Count, IndexOf) Then
        Exit Function
    End If
    
    Dim i As Long
    For i = Range.Index To Range.Index + Range.Count - 1
        If Object.Equals(Arr(i), Value) Then
            IndexOf = i
            Exit Function
        End If
    Next
    
    IndexOf = LBound(Arr) - 1
    Exit Function
    
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Returns if an array variable is uninitialized.
'
' @param arr The array to test for initialization.
' @return Value indicating if the array is initialized.
Public Function IsNull(ByRef Arr As Variant) As Boolean
    IsNull = ArrayPointer(Arr) = vbNullPtr
End Function

''
' Returns the index of the last occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param Arr The array to search.
' @param Value The value to search for in the array.
' @param StartIndex The index to the element in the array to begin search from.
' @param Count The number of elements to search, starting from startindex.
' @return Value indicating the index the value was found. If no value was found,
' then one less than the lower bound of the array is returned.
'
Public Function LastIndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Ptr As Long
    Ptr = ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = OptionalLong(StartIndex, UBound(Arr))
    ElemCount = OptionalLong(Count, ElemIndex - LBound(Arr) + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_ArrayBounds
    If ElemCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    If ElemIndex - ElemCount + 1 < LBound(Arr) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_LBound
    
    On Error GoTo errTrap
    If TrySZLastIndexOf(Ptr, Value, ElemIndex, ElemCount, LastIndexOf) Then
        Exit Function
    End If
    
    Do While ElemCount > 0
        If Object.Equals(Arr(ElemIndex), Value) Then
            LastIndexOf = ElemIndex
            Exit Function
        End If
        ElemIndex = ElemIndex - 1
        ElemCount = ElemCount - 1
    Loop
    
    LastIndexOf = LBound(Arr) - 1
    Exit Function
    
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Reverses the elements in a subportion of an array.
'
' @param Arr The array to reverse.
' @param Index The starting element to begin reversing elements.
' @param Count The number of elements to reverse in the array.
'
Public Sub Reverse(ByRef Arr As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim Ptr As Long
    Ptr = CorArray.ArrayPointer(Arr)
    
    CheckValidSingleDimArrayPtr Ptr
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = OptionalLong(Index, LBound(Arr))
    ElemCount = OptionalLong(Count, UBound(Arr) - ElemIndex + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If ElemCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If ElemIndex + ElemCount - 1 > UBound(Arr) Then _
        Error.InvalidOffsetLength
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(Ptr)
    
    Dim pLowElem As Long
    pLowElem = MemLong(Ptr + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    
    Dim pHighElem As Long
    pHighElem = pLowElem + (ElemCount - 1) * ElemSize
    
    Select Case ElemSize
        Case 1, 2, 4, 8, 16
            Do While pLowElem < pHighElem
                Select Case ElemSize
                    Case 1: Helper.Swap1 ByVal pLowElem, ByVal pHighElem
                    Case 2: Helper.Swap2 ByVal pLowElem, ByVal pHighElem
                    Case 4: Helper.Swap4 ByVal pLowElem, ByVal pHighElem
                    Case 8: Helper.Swap8 ByVal pLowElem, ByVal pHighElem
                    Case 16: Helper.Swap16 ByVal pLowElem, ByVal pHighElem
                End Select
                pLowElem = pLowElem + ElemSize
                pHighElem = pHighElem - ElemSize
            Loop
            
        Case Else
            Dim pSwapBuffer As Long
            pSwapBuffer = CoTaskMemAlloc(ElemSize)
            
            Do While pLowElem < pHighElem
                CopyMemory ByVal pSwapBuffer, ByVal pLowElem, ElemSize
                CopyMemory ByVal pLowElem, ByVal pHighElem, ElemSize
                CopyMemory ByVal pHighElem, ByVal pSwapBuffer, ElemSize
                pLowElem = pLowElem + ElemSize
                pHighElem = pHighElem - ElemSize
            Loop
            CoTaskMemFree pSwapBuffer
    End Select
End Sub

''
' Sorts an entire array. An optionally supplied comparer object can be
' used to compare special elements, such as userdefined values.
'
' @param Arr The array to sort.
' @param Comparer (Optional) An IComparer object or address of a comparer callback function.
' @remarks The Comparer parameter can be an <b>IComparer</b> object or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As [Array Datatype], ByRef y As [Array Datatype]) As Long
'     '' return a negative value if x is less than y.
'     '' return a positive value if x is greater than y.
'     '' return 0 if x equals y.
' End Function
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray_Sort.txt"
Public Sub Sort(ByRef Arr As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    Dim Length As Long
    Length = UBound(Arr) - LBound(Arr) + 1
    SortKeyEx Arr, Empty, LBound(Arr), Length, Comparer
End Sub

''
' Sorts an array, or subportion, given a startindex and length. An optionally
' supplied comparer object can be used to compare special elements, such as
' userdefined values.
'
' @param Arr The array to sort.
' @param Index The starting index in the array to start sorting.
' @param Count The number of elements to be sorted.
' @param Comparer (Optional) An IComparer object or address of a comparer callback function.
' @remarks The Comparer parameter can be an <b>IComparer</b> object or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As [Array Datatype], ByRef y As [Array Datatype]) As Long
'     '' return a negative value if x is less than y.
'     '' return a positive value if x is greater than y.
'     '' return 0 if x equals y.
' End Function
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
Public Sub SortEx(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, Optional ByRef Comparer As Variant)
    If IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    SortKeyEx Arr, Empty, Index, Length, Comparer
End Sub

''
' Sorts an entire array based on an array of keys. An optionally supplied
' comparer object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param Comparer (Optional) An IComparer object or address of a comparer callback function.
' @remarks The Comparer parameter can be an <b>IComparer</b> object or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As [Array Datatype], ByRef y As [Array Datatype]) As Long
'     '' return a negative value if x is less than y.
'     '' return a positive value if x is greater than y.
'     '' return 0 if x equals y.
' End Function
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
' @include "..\Includes\CorArray.SortKey.txt"
Public Sub SortKey(ByRef Keys As Variant, ByRef Items As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Keys) Then _
        Error.ArgumentNull "Keys", ArgumentNull_Array
    
    Dim Length As Long
    Length = UBound(Keys) - LBound(Keys) + 1
    SortKeyEx Keys, Items, LBound(Keys), Length, Comparer
End Sub

''
' Sorts an array based on an array of keys. A subportion of the array can
' be sorted given a startindex and length. An optionally, supplied comparer
' object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param Index The starting index in the array to start sorting.
' @param Length The number of elements to be sorted.
' @param Comparer (Optional) An IComparer object or address of a comparer callback function.
' @remarks The Comparer parameter can be an <b>IComparer</b> object or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As [Array Datatype], ByRef y As [Array Datatype]) As Long
'     '' return a negative value if x is less than y.
'     '' return a positive value if x is greater than y.
'     '' return 0 if x equals y.
' End Function
' </pre>
' The [Array Datatype] must be replaced with the datatype of the array. If the
' array is an array of Variants, then [Array Datatype] would be a Variant, not
' any specific sub-type within the variants of the array.</p>
'
Public Sub SortKeyEx(ByRef Keys As Variant, ByRef Items As Variant, ByVal Index As Long, ByVal Length As Long, Optional ByRef Comparer As Variant)
    Dim KeysPtr  As Long
    Dim ItemsPtr As Long
    
    KeysPtr = ArrayPointer(Keys)
    CheckValidSingleDimArrayPtr KeysPtr, Parameter_Keys
    
    If Index < LBound(Keys) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum
    If Index + Length - 1 > UBound(Keys) Then _
        Error.InvalidOffsetLength
    
    If Not IsEmpty(Items) Then
        If Not IsArray(Items) Then _
            Error.Argument Argument_ArrayRequired, "Items"
        
        ItemsPtr = ArrayPointer(Items)
        CheckValidSingleDimArrayPtr ItemsPtr, Parameter_Items
        
        ' Items array must match the Keys array in size and bounds.
        If (LBound(Keys) <> LBound(Items)) Or _
           (UBound(Keys) <> UBound(Items)) Then _
            Error.Argument Argument_MatchingBounds
        
        SetSortItems ItemsPtr
    End If
    
    On Error GoTo errTrap
    ClearException
    Select Case VarType(Comparer)
        ' use an IComparer object to sort with.
        Case vbObject
            SetSortKeys KeysPtr
            
            If Comparer Is Nothing Then
                Set SortComparer = VBCorLib.Comparer.Default
            Else
                Set SortComparer = Comparer
            End If
            
            QuickSortGeneral Keys, Index, Index + Length - 1
            
        ' Use a callback method to sort with.
        ' This allows a user to supply a highspeed comparison
        ' routine to compare two elements of the array.
        Case vbLong
            CallbackSorter.Sort KeysPtr, ItemsPtr, Index, Index + Length - 1, Comparer
            
        ' try optimized routines first, then fallback to the default IComparer object.
        Case vbError
            If Not TrySZSort(KeysPtr, Index, Index + Length - 1) Then
                ' If we get here, then no optimized routine existed,
                ' so we'll set our comparer to the default and loop
                ' around again, letting this case statement fall
                ' into the QuickSortGeneral routine.
                SetSortKeys KeysPtr
                Set SortComparer = VBCorLib.Comparer.Default
                QuickSortGeneral Keys, Index, Index + Length - 1
            End If
        Case Else
            Error.Argument Argument_InvalidComparer, "Comparer"
    End Select
    
    CleanUpSortRoutine
    Exit Sub
    
errTrap:
    CleanUpSortRoutine
    Throw
End Sub

Private Sub CleanUpSortRoutine()
    ClearSortKeys
    ClearSortItems
    Set SortComparer = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Internal
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function ArrayPointer(ByRef Arg As Variant) As Long
    Dim ArgType As Integer
    ArgType = VariantType(Arg)
    
    If (ArgType And vbArray) = 0 Then _
        Error.Argument Argument_ArrayRequired
    
    ArrayPointer = MemLong(VB6.vbaRefVarAry(Arg))
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any CorArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case ArgType And &HFF
        Case vbObject, vbUserDefinedType
            If UBound(Arg) < LBound(Arg) Then
                ArrayPointer = vbNullPtr
            End If
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BinarySearchCore(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, ByRef Value As Variant, ByVal Comparer As IComparer) As Long
    On Error GoTo errTrap
    
    If Comparer Is Nothing Then
        If CorArrayHelper.TrySZBinarySearch(ArrayPointer(Arr), Value, Index, Length, BinarySearchCore) Then
            Exit Function
        End If
        
        Set Comparer = Statics.Comparer.Default
    End If
    
    ' perform a search using comparer.
    Dim MiddleIndex As Long
    Dim LowIndex    As Long
    Dim HighIndex   As Long
    
    LowIndex = Index
    HighIndex = LowIndex + Length - 1
    Do While LowIndex <= HighIndex
        MiddleIndex = (LowIndex + HighIndex) \ 2
        Select Case Comparer.Compare(Arr(MiddleIndex), Value)
            Case 0
                BinarySearchCore = MiddleIndex
                Exit Function
            Case Is > 0
                HighIndex = MiddleIndex - 1
            Case Else
                LowIndex = MiddleIndex + 1
        End Select
    Loop
    
    BinarySearchCore = (Not (LowIndex - LBound(Arr))) + LBound(Arr)
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' This function assigns the default value of the specific
' datatype for the array. If the array is a vbLong, then
' a 0 of type long is assiged. This is to ensure that the
' Variant return value contains the same datatype as the array.
'
Private Sub AssignDefaultElementValue(ByRef Arr As Variant, ByRef RetVal As Variant)
    ' we use the VB conversion functions to have easier
    ' readability as to what we are assigning.
    Select Case VarType(Arr) And &HFF
        Case vbLong:        RetVal = CLng(0)
        Case vbInteger:     RetVal = CInt(0)
        Case vbByte:        RetVal = CByte(0)
        Case vbString:      RetVal = vbNullString
        Case vbObject:      Set RetVal = Nothing
        Case vbDouble:      RetVal = CDbl(0)
        Case vbSingle:      RetVal = CSng(0)
        Case vbDate:        RetVal = CDate(#12:00:00 AM#)
        Case vbCurrency:    RetVal = CCur(0)
        Case vbBoolean:     RetVal = False
        Case vbVariant:     RetVal = Empty
        Case vbUserDefinedType
            Dim rec As IRecordInfo
            Set rec = SafeArrayGetRecordInfo(CorArray.ArrayPointer(Arr))
            
            ' Set the datatype for the variant.
            VariantType(RetVal) = vbUserDefinedType
            ' Set the pointer to the new structure created.
            ' This structure is of the same type as contained in the array.
            MemLong(VarPtr(RetVal) + VARIANTDATA_OFFSET) = rec.RecordCreate
            ' Set the pointer to the IRecordInfo object so VB can
            ' inspect this variant and handle it properly.
            MemLong(VarPtr(RetVal) + VARIANTDATA_OFFSET + 4) = ObjectPtr(rec)
            ' Kill our reference to the IRecordInfo without decrementing
            ' the reference count, since we assigned it without adding a reference.
            ObjectPtr(rec) = vbNullPtr
    End Select
End Sub

' Value types can be copied quickly using CopyMemory.
'
Private Sub CopyValueTypes(ByVal srcpSA As Long, ByVal SourceIndex As Long, ByVal dstpSA As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(srcpSA)
    
    Dim pSrcElem As Long
    pSrcElem = MemLong(srcpSA + PVDATA_OFFSET) + (SourceIndex - SafeArrayGetLBound(srcpSA, 1)) * ElemSize
    
    Dim pDstElem As Long
    pDstElem = MemLong(dstpSA + PVDATA_OFFSET) + (DestinationIndex - SafeArrayGetLBound(dstpSA, 1)) * ElemSize
    
    CopyMemory ByVal pDstElem, ByVal pSrcElem, Length * ElemSize
End Sub

' Copies reference types, dealing with possible overlapping source and destination regions.
'
Private Sub CopyReferenceTypes(ByVal srcpSA As Long, ByVal SourceIndex As Long, ByVal dstpSA As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim srcSA       As GuidSafeArray1d
    Dim dstSA       As GuidSafeArray1d
    Dim IIDOffset   As Long
    
    Select Case SafeArrayGetVartype(srcpSA)
        Case vbObject:              IIDOffset = SizeOfGuid
        Case vbUserDefinedType:     IIDOffset = 4
    End Select
    
    CopyMemory ByVal VarPtr(srcSA.SA) - IIDOffset, ByVal srcpSA - IIDOffset, LenB(srcSA.SA) + IIDOffset
    CopyMemory ByVal VarPtr(dstSA.SA) - IIDOffset, ByVal dstpSA - IIDOffset, LenB(dstSA.SA) + IIDOffset
    
    If srcpSA = dstpSA Then
        ' test for overlapping source and destination portion of the same array
        If (DestinationIndex > SourceIndex) And (DestinationIndex < SourceIndex + Length) Then
            Dim ClearSA As GuidSafeArray1d
            ClearSA = srcSA
            With ClearSA.SA
                .pvData = .pvData + (SourceIndex + Length) * .cbElements
                .cElements = DestinationIndex - SourceIndex
            End With
            SafeArrayDestroyData VarPtr(ClearSA.SA)
            
            With srcSA.SA
                CopyMemory ByVal .pvData + (DestinationIndex - .lLbound) * .cbElements, ByVal .pvData + (SourceIndex - .lLbound) * .cbElements, Length * .cbElements
    
                ' If we have moved reference type values, then we need to
                ' fill in the uncovered portion of the original values
                ' with copies. First we have to zero out the uncovered portion
                ' so the copy won't release any references, since we merely
                ' moved the references without making copies to objects or strings.
                ' Once the uncovered portion is set up, we can copy backwards using
                ' the standard reference copy method without stepping on ourselves.
                ZeroMemory ByVal .pvData + (SourceIndex - .lLbound) * .cbElements, (DestinationIndex - SourceIndex) * .cbElements
            End With
            
            Length = DestinationIndex - SourceIndex
            Dim t As Long
            t = SourceIndex
            SourceIndex = DestinationIndex
            DestinationIndex = t
        End If
    End If
    
    SetCopyArrayDescriptor srcSA.SA, SourceIndex, Length
    SetCopyArrayDescriptor dstSA.SA, DestinationIndex, Length
    
    SafeArrayCopyData VarPtr(srcSA.SA), VarPtr(dstSA.SA)
End Sub

Private Sub CopyWithCast(ByVal srcpSA As Long, ByVal SourceIndex As Long, ByVal dstpSA As Long, ByVal DestinationIndex As Long, ByVal Length As Long, ByVal pfn As Long)
    Dim srcSize As Long
    srcSize = SafeArrayGetElemsize(srcpSA)
    
    Dim srcLb   As Long
    srcLb = SafeArrayGetLBound(srcpSA, 1)
    
    Dim src     As Long
    src = MemLong(srcpSA + PVDATA_OFFSET) + (SourceIndex - srcLb) * srcSize
    
    Dim dstSize As Long
    dstSize = SafeArrayGetElemsize(dstpSA)
    
    Dim dstLb   As Long
    dstLb = SafeArrayGetLBound(dstpSA, 1)
    
    Dim dst     As Long
    dst = MemLong(dstpSA + PVDATA_OFFSET) + (DestinationIndex - dstLb) * dstSize
    
    Dim Assigner As Action_T_T
    Set Assigner = NewDelegate(pfn)
    
    Do While Length > 0
        Assigner.Invoke ByVal dst, ByVal src
        src = src + srcSize
        dst = dst + dstSize
        Length = Length - 1
    Loop
End Sub

' Copies objects and checks for compatible interfaces during the copy.
'
Private Sub CopyObjectTypesCast(ByVal srcpSA As Long, ByVal SourceIndex As Long, ByVal dstpSA As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Const SizeOfPointer As Long = 4
    Dim SrcLowerBound As Long
    SrcLowerBound = SafeArrayGetLBound(srcpSA, 1)
    
    Dim DstLowerBound As Long
    DstLowerBound = SafeArrayGetLBound(dstpSA, 1)
    
    Dim pSrcElem As Long
    pSrcElem = MemLong(srcpSA + PVDATA_OFFSET) + (SourceIndex - SrcLowerBound) * SizeOfPointer
    
    Dim pDstElem As Long
    pDstElem = MemLong(dstpSA + PVDATA_OFFSET) + (DestinationIndex - DstLowerBound) * SizeOfPointer
    
    Dim pDstGuid As Long
    pDstGuid = dstpSA - SizeOfGuid
    
    Dim pObj    As Long
    Dim SrcObj  As IVBUnknown
    Dim DstObj  As IVBUnknown
    Do While Length > 0
        ObjectPtr(SrcObj) = MemLong(pSrcElem)
        pObj = 0
        If Not SrcObj Is Nothing Then
            If SrcObj.QueryInterface(ByVal pDstGuid, pObj) = E_NOINTERFACE Then
                ObjectPtr(SrcObj) = vbNullPtr
                ObjectPtr(DstObj) = vbNullPtr
                Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
            End If
        End If
        
        ObjectPtr(DstObj) = MemLong(pDstElem)
        If Not DstObj Is Nothing Then
            DstObj.Release
        End If
        MemLong(pDstElem) = pObj
        
        pSrcElem = pSrcElem + SizeOfPointer
        pDstElem = pDstElem + SizeOfPointer
        Length = Length - 1
    Loop
    
    ObjectPtr(SrcObj) = vbNullPtr
    ObjectPtr(DstObj) = vbNullPtr
End Sub

' Returns the address to a function that can assign one
' datatype to another datatype. If the datatypes are not
' compatible then 0 is returned.
'
Private Function GetAssigningFunction(ByVal srcType As VbVarType, ByVal dstType As VbVarType) As Long
    Dim Ret As Long
    Select Case srcType
        Case vbLong
            Select Case dstType
                Case vbDouble:      Ret = FuncAddr(AddressOf WidenLongToDouble)
                Case vbSingle:      Ret = FuncAddr(AddressOf WidenLongToSingle)
                Case vbString:      Ret = FuncAddr(AddressOf WidenLongToString)
                Case vbCurrency:    Ret = FuncAddr(AddressOf WidenLongToCurrency)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenLongToVariant)
            End Select
        Case vbInteger, vbBoolean
            Select Case dstType
                Case vbLong:        Ret = FuncAddr(AddressOf WidenIntegerToLong)
                Case vbDouble:      Ret = FuncAddr(AddressOf WidenIntegerToDouble)
                Case vbString:      Ret = FuncAddr(AddressOf WidenIntegerToString)
                Case vbSingle:      Ret = FuncAddr(AddressOf WidenIntegerToSingle)
                Case vbCurrency:    Ret = FuncAddr(AddressOf WidenIntegerToCurrency)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenIntegerToVariant)
            End Select
        Case vbByte
            Select Case dstType
                Case vbLong:        Ret = FuncAddr(AddressOf WidenByteToLong)
                Case vbInteger:     Ret = FuncAddr(AddressOf WidenByteToInteger)
                Case vbDouble:      Ret = FuncAddr(AddressOf WidenByteToDouble)
                Case vbString:      Ret = FuncAddr(AddressOf WidenByteToString)
                Case vbSingle:      Ret = FuncAddr(AddressOf WidenByteToSingle)
                Case vbCurrency:    Ret = FuncAddr(AddressOf WidenByteToCurrency)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenByteToVariant)
            End Select
        Case vbDouble
            Select Case dstType
                Case vbString:      Ret = FuncAddr(AddressOf WidenDoubleToString)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenDoubleToVariant)
            End Select
        Case vbSingle
            Select Case dstType
                Case vbDouble:      Ret = FuncAddr(AddressOf WidenSingleToDouble)
                Case vbString:      Ret = FuncAddr(AddressOf WidenSingleToString)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenSingleToVariant)
            End Select
        Case vbDate
            Select Case dstType
                Case vbDouble:      Ret = FuncAddr(AddressOf WidenDateToDouble)
                Case vbString:      Ret = FuncAddr(AddressOf WidenDateToString)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenDateToVariant)
            End Select
        Case vbObject, vbDataObject
            Select Case dstType
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenObjectToVariant)
            End Select
        Case vbCurrency
            Select Case dstType
                Case vbString:      Ret = FuncAddr(AddressOf WidenCurrencyToString)
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenCurrencyToVariant)
                Case vbDouble:      Ret = FuncAddr(AddressOf WidenCurrencyToDouble)
            End Select
        Case vbString
            Select Case dstType
                Case vbVariant:     Ret = FuncAddr(AddressOf WidenStringToVariant)
            End Select
        Case vbVariant
            Select Case dstType
                Case vbLong:                    Ret = FuncAddr(AddressOf NarrowVariantToLong)
                Case vbInteger:                 Ret = FuncAddr(AddressOf NarrowVariantToInteger)
                Case vbString:                  Ret = FuncAddr(AddressOf NarrowVariantToString)
                Case vbDouble:                  Ret = FuncAddr(AddressOf NarrowVariantToDouble)
                Case vbObject, vbDataObject:    Ret = FuncAddr(AddressOf NarrowVariantToObject)
                Case vbSingle:                  Ret = FuncAddr(AddressOf NarrowVariantToSingle)
                Case vbByte:                    Ret = FuncAddr(AddressOf NarrowVariantToByte)
                Case vbDate:                    Ret = FuncAddr(AddressOf NarrowVariantToDate)
                Case vbBoolean:                 Ret = FuncAddr(AddressOf NarrowVariantToBoolean)
                Case vbCurrency:                Ret = FuncAddr(AddressOf NarrowVariantToCurrency)
            End Select
    End Select
    GetAssigningFunction = Ret
End Function

Private Sub SetCopyArrayDescriptor(ByRef desc As SafeArray1d, ByVal Index As Long, ByVal Length As Long)
    With desc
        .pvData = .pvData + (Index - .lLbound) * .cbElements
        .cElements = Length
    End With
End Sub

''
'   Attempt to locate an optimized sort routine for a specific data type.
Private Function TrySZSort(ByVal pSA As Long, ByVal Left As Long, ByVal Right As Long) As Boolean
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
    mSortDelegator.pfn = pfn
    mSorter.Call pSA, Left, Right
    TrySZSort = True
End Function


