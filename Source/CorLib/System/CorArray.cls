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
' Debug.Print CorArray.Length(MyArray)
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

Private Type ArrayInfo
    SAPtr       As Long
    DataPtr     As Long
    ElementSize As Long
End Type

Private CallbackSorter As New CallbackSorter


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
' @include "..\DocIncludes\CorArray.Find.html"
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
' @include "..\DocIncludes\CorArray.Find.html"
Public Function FindIndex(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Info As ArrayInfo
    Dim ElementIndex   As Long
    Dim ElementCount   As Long
    Dim ElementPtr As Long
    Dim Predicate As Func_T_Boolean
    
    ValidateArray Arr
    ElementIndex = CLngOrDefault(Index, LBound(Arr))
    ElementCount = CLngOrDefault(Count, UBound(Arr) - ElementIndex + 1)
    
    If ElementIndex < LBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If ElementCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    If ElementIndex + ElementCount - 1 > UBound(Arr) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_UBound
        
    GetArrayInfo Arr, Info
    ElementPtr = Info.DataPtr + (ElementIndex - LBound(Arr)) * Info.ElementSize
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    Dim i As Long
    For i = 1 To ElementCount
        If Predicate.Invoke(ByVal ElementPtr) Then
            FindIndex = ElementIndex + i - 1
            Exit Function
        End If
        
        ElementPtr = ElementPtr + Info.ElementSize
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
' @include "..\DocIncludes\CorArray.Find.html"
Public Function FindAll(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Variant
    Dim Info                As ArrayInfo
    Dim MatchCount          As Long
    Dim MatchedIndexes()    As Long
    
    ValidateArray Arr
    GetArrayInfo Arr, Info
    MatchedIndexes = FindAllMatches(Arr, Info, AddressOfPredicate, MatchCount)
    
    Select Case VarType(Arr) And &HFF
        Case vbObject
            SetFindAllObjectResult Arr, Info, MatchedIndexes, MatchCount, FindAll
        Case vbUserDefinedType
            SetFindAllUdtResult Arr, Info, MatchedIndexes, MatchCount, FindAll
        Case Else
            SetFindAllResult Arr, Info, MatchedIndexes, MatchCount, FindAll
    End Select
End Function

Private Function FindAllMatches(ByRef Arr As Variant, ByRef Info As ArrayInfo, ByVal AddressOfPredicate As Long, ByRef MatchCount As Long) As Long()
    Dim ElementPtr          As Long
    Dim Predicate           As Func_T_Boolean
    Dim MatchedIndexes()    As Long
    
    ReDim MatchedIndexes(0 To 31)
    ElementPtr = Info.DataPtr
    Set Predicate = NewDelegate(AddressOfPredicate)

    ' iterate each array element, passing the memory address
    ' to the user supplied AddressOf callback. This will appear
    ' as a ByRef parameter of the array's datatype in the callback.
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If Predicate.Invoke(ByVal ElementPtr) Then
            If MatchCount > UBound(MatchedIndexes) Then
                ReDim Preserve MatchedIndexes(0 To MatchCount * 2 - 1)
            End If
            
            MatchedIndexes(MatchCount) = i
            MatchCount = MatchCount + 1
        End If
        
        ElementPtr = ElementPtr + Info.ElementSize
    Next i
    
    FindAllMatches = MatchedIndexes
End Function

Private Sub SetFindAllObjectResult(ByRef Arr As Variant, ByRef Info As ArrayInfo, ByRef MatchedIndexes() As Long, ByVal MatchCount As Long, ByRef Result As Variant)
    Dim i As Long
    
    ' as an array of objects, we create a new array using the
    ' same Guid value as that of the original array.
    VariantType(Result) = vbObject Or vbArray
    MemLong(VarPtr(Result) + VARIANTDATA_OFFSET) = SafeArrayCreateVectorEx(vbObject, 0, MatchCount, Info.SAPtr - vbSizeOfGuid)
    
    For i = 0 To MatchCount - 1
        Set Result(i) = Arr(MatchedIndexes(i))
    Next i
End Sub

Private Sub SetFindAllResult(ByRef Arr As Variant, ByRef Info As ArrayInfo, ByRef MatchedIndexes() As Long, ByVal MatchCount As Long, ByRef Result As Variant)
    Dim i As Long
    
    VariantType(Result) = VarType(Arr)
    MemLong(VarPtr(Result) + VARIANTDATA_OFFSET) = SafeArrayCreateVectorEx(VarType(Arr) And &HFF, 0, MatchCount, Info.SAPtr - vbSizeOfGuid)

    For i = 0 To MatchCount - 1
        Result(i) = Arr(MatchedIndexes(i))
    Next i
End Sub

Private Sub SetFindAllUdtResult(ByRef Arr As Variant, ByRef Info As ArrayInfo, ByRef MatchedIndexes() As Long, ByVal MatchCount As Long, ByRef Result As Variant)
    Dim Record As IRecordInfo
    Dim pReturnSafeArray As Long
    Dim pDstElem As Long
    Dim LowerBound As Long
    Dim i As Long
    
    Set Record = SafeArrayGetRecordInfo(Info.SAPtr)    ' gets an object used to manipulate a UDT.
    pReturnSafeArray = SafeArrayCreateVectorEx(vbUserDefinedType, 0, MatchCount, ObjPtr(Record))
    pDstElem = MemLong(pReturnSafeArray + PVDATA_OFFSET)
    LowerBound = LBound(Arr)
    
    For i = 0 To MatchCount - 1
        Record.RecordCopy Info.DataPtr + (MatchedIndexes(i) - LowerBound) * Info.ElementSize, pDstElem
        pDstElem = pDstElem + Info.ElementSize
    Next i
    
    VariantType(Result) = vbUserDefinedType Or vbArray
    MemLong(VarPtr(Result) + VARIANTDATA_OFFSET) = pReturnSafeArray
End Sub

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
' @include "..\DocIncludes\CorArray.Find.html"
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
' @include "..\DocIncludes\CorArray.Find.html"
Public Function FindLastIndex(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim Info        As ArrayInfo
    Dim ElementPtr  As Long
    Dim Predicate   As Func_T_Boolean
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    
    ValidateArray Arr
    ElemIndex = CLngOrDefault(Index, UBound(Arr))
    ElemCount = CLngOrDefault(Count, ElemIndex + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayBounds
    If ElemCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If ElemIndex - ElemCount + 1 < LBound(Arr) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_LBound
    
    GetArrayInfo Arr, Info
    ElementPtr = Info.DataPtr + (ElemIndex - LBound(Arr)) * Info.ElementSize
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    Dim i As Long
    For i = ElemCount To 1 Step -1
        If Predicate.Invoke(ByVal ElementPtr) Then
            FindLastIndex = LBound(Arr) + i - 1
            Exit Function
        End If
        
        ElementPtr = ElementPtr - Info.ElementSize
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
    Dim Info        As ArrayInfo
    Dim ElementPtr  As Long
    Dim Predicate   As Func_T_Boolean
        
    ValidateArray Arr
    GetArrayInfo Arr, Info
    ElementPtr = Info.DataPtr
    Set Predicate = NewDelegate(AddressOfPredicate)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If Not Predicate.Invoke(ByVal ElementPtr) Then
            Exit Function
        End If
        
        ElementPtr = ElementPtr + Info.ElementSize
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
    Dim Info        As ArrayInfo
    Dim ElementPtr  As Long
    Dim Action      As Action_T
    
    ValidateArray Arr
    GetArrayInfo Arr, Info
    ElementPtr = Info.DataPtr
    Set Action = NewDelegate(AddressOfAction)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Action.Invoke ByVal ElementPtr
        ElementPtr = ElementPtr + Info.ElementSize
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
' @include "..\DocIncludes\CorArray.BinarySearch.html"
Public Function BinarySearch(ByRef Arr As Variant, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ValidateArray Arr
    BinarySearch = InternalBinarySearch(Arr, LBound(Arr), Len1D(Arr), Value, Comparer)
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
' @include "..\DocIncludes\CorArray.BinarySearch.html"
Public Function BinarySearchEx(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ValidateArrayRange Arr, Index, Length, CountName:=NameOfLength
    BinarySearchEx = InternalBinarySearch(Arr, Index, Length, Value, Comparer)
End Function

''
' Clears a portion of the elements in an array.
'
' @param Arr The array to clear elements from.
' @param Index The starting element to being clearing.
' @param Length The number of elements to be cleared.
'
Public Sub Clear(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long)
    Dim pSafeArray  As Long
    
    pSafeArray = SAPtrV(Arr)
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
            IIDOffset = vbSizeOfGuid
            
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
' @param Length The number of elements to copy.
'
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
' @param Sourcearray The array from which to copy the elements.
' @param Sourceindex The starting element in the source array to begin copying from.
' @param Destinationarray The array in which to place the elements.
' @param Destinationindex The starting element in the destination array to place elements.
' @param Length The number of elements to copy.
'
Public Sub CopyEx(ByRef SourceArray As Variant, ByVal SourceIndex As Long, ByRef DestinationArray As Variant, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim pSourceSafeArray As Long
    pSourceSafeArray = SAPtrV(SourceArray)
    If pSourceSafeArray = vbNullPtr Then _
        Error.ArgumentNull "SourceArray", ArgumentNull_Array
    
    Dim pDestinationSafeArray As Long
    pDestinationSafeArray = SAPtrV(DestinationArray)
    If pDestinationSafeArray = vbNullPtr Then _
        Error.ArgumentNull "DestinationArray", ArgumentNull_Array
    
    Dim SourceLength As Long
    SourceLength = LengthCore(pSourceSafeArray)
    Dim DestinationLength As Long
    DestinationLength = LengthCore(pDestinationSafeArray)
    
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum
    If Length + SourceIndex - LBound(SourceArray) > SourceLength Then _
        Error.Argument Arg_LongerThanSrcArray
    If Length + DestinationIndex - LBound(DestinationArray) > DestinationLength Then _
        Error.Argument Arg_LongerThanDestArray
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
                If IsEqualGUID(ByVal pSourceSafeArray - vbSizeOfGuid, ByVal pDestinationSafeArray - vbSizeOfGuid) <> BOOL_FALSE Then
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
' This does not mean that lengths 2 and 3 need to be set to zero if only an empty
' one-dimensional array is being created.
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
        Error.ArgumentOutOfRange "Length1", ArgumentOutOfRange_NeedNonNegNum
    
    Rank = 1
    Bounds(0).cElements = Length1
    
    If Not IsMissing(Length2) Then
        If CLng(Length2) < 0 Then _
            Error.ArgumentOutOfRange "Length2", ArgumentOutOfRange_NeedNonNegNum
        
        Rank = 2
        Bounds(1).cElements = Length2
    
        If Not IsMissing(Length3) Then
            If CLng(Length3) < 0 Then _
                Error.ArgumentOutOfRange "Length3", ArgumentOutOfRange_NeedNonNegNum
            
            Rank = 3
            Bounds(2).cElements = Length3
        End If
    End If
        
    VariantType(CreateInstance) = ElementType Or vbArray
    MemLong(VarPtr(CreateInstance) + VARIANTDATA_OFFSET) = SafeArrayCreate(ElementType, Rank, Bounds(0))
End Function

''
' Gets a 32-bit integer that represents the total number of elements in all dimensions of the array.
'
' @param Arr The array to retrieve the number of elements from.
' @return A 32-bit integer that represents the total number of elements in all the dimensions of the Array; zero if there are no elements in the array.
Public Function Length(ByRef Arr As Variant) As Long
    Dim Ptr As Long
    
    Ptr = GetValidSAPtr(Arr)
    Length = LengthCore(Ptr)
End Function

''
' Gets a 32-bit integer that represents the number of elements in the specified dimension of the array.
'
' @param Arr The array to get the dimension length from.
' @param Dimension A one-based dimension of the array whose length needs to be determined.
' @return A 32-bit integer that represents the number of elements in the specified dimension.
' @include "..\DocIncludes\CorArray.GetLength.html"
Public Function GetLength(ByRef Arr As Variant, ByVal Dimension As Long) As Long
    Dim ArrayPtr    As Long
    Dim Rank        As Long
    Dim BoundPtr    As Long
    
    ArrayPtr = GetValidSAPtr(Arr)
    Rank = SafeArrayGetDim(ArrayPtr)
    
    If Dimension < 1 Or Dimension > Rank Then _
        Error.IndexOutOfRange IndexOutOfRange_Dimension
    
    BoundPtr = ArrayPtr + vbSizeOfSafeArray + (Rank - Dimension) * vbSizeOfSafeArrayBound
    GetLength = MemLong(BoundPtr)
End Function

''
' Returns the number of dimensions in the array.
'
' @param Arr The array to retrieve the number of dimensions from.
' @return The number of dimension in the array.
Public Function Rank(ByRef Arr As Variant) As Long
    Dim Ptr As Long
    
    If Not IsArray(Arr) Then _
        Error.Argument Argument_ArrayRequired
        
    Ptr = SAPtrV(Arr)
    
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
    Dim Ptr     As Long
    Dim Range   As ListRange
        
    ValidateArray Arr
    Range = MakeDefaultStepRange(StartIndex, LBound(Arr), Count, Len1D(Arr), Parameter_StartIndex)
    
    If Range.Index < LBound(Arr) Or Range.Index > UBound(Arr) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_ArrayBounds
    If Range.Count < 0 Or Range.Index + Range.Count - 1 > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    
    On Error GoTo Catch
    Ptr = SAPtrV(Arr)
    
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
    
Catch:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Returns if an array variable is uninitialized.
'
' @param arr The array to test for initialization.
' @return Returns <b>True</b> if the array is uninitialized, <b>False</b> otherwise.
Public Function IsNull(ByRef Arr As Variant) As Boolean
    IsNull = SAPtrV(Arr) = vbNullPtr
End Function

''
' Returns if the array variable is unitialized(null) or empty(no elements).
'
' @param Arr The array to inspectd.
' @return Returns <b>True</b> if the array is unitialized or has zero elements, <b>False</b> otherwise.
Public Function IsNullOrEmpty(ByRef Arr As Variant) As Boolean
    If IsNull(Arr) Then
        IsNullOrEmpty = True
    ElseIf Length(Arr) = 0 Then
        IsNullOrEmpty = True
    End If
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
Public Function LastIndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Ptr         As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    
    ValidateArray Arr
    ElemIndex = CLngOrDefault(StartIndex, UBound(Arr))
    ElemCount = CLngOrDefault(Count, ElemIndex - LBound(Arr) + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_ArrayBounds
    If ElemCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    If ElemIndex - ElemCount + 1 < LBound(Arr) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_LBound
    
    On Error GoTo Catch
    Ptr = SAPtrV(Arr)
    
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
    
Catch:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Reverses the elements in a subportion of an array.
'
' @param Arr The array to reverse.
' @param Index The starting element to begin reversing elements.
' @param Count The number of elements to reverse in the array.
Public Sub Reverse(ByRef Arr As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim Info        As ArrayInfo
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Element1Ptr As Long
    Dim Element2Ptr As Long
    Dim Swapper     As Action_T_T
    
    ValidateArray Arr
    ElemIndex = CLngOrDefault(Index, LBound(Arr))
    ElemCount = CLngOrDefault(Count, UBound(Arr) - ElemIndex + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If ElemCount < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If ElemIndex + ElemCount - 1 > UBound(Arr) Then _
        Error.Argument Argument_InvalidOffLen
    
    GetArrayInfo Arr, Info
    Element1Ptr = Info.DataPtr + (ElemIndex - LBound(Arr)) * Info.ElementSize
    Element2Ptr = Element1Ptr + (ElemCount - 1) * Info.ElementSize
    
    Select Case Info.ElementSize
        Case 1
            Set Swapper = NewDelegate(AddressOf SwapByte)
        Case 2
            Set Swapper = NewDelegate(AddressOf SwapInteger)
        Case 4
            Set Swapper = NewDelegate(AddressOf SwapLong)
        Case 8
            Set Swapper = NewDelegate(AddressOf SwapCurrency)
        Case 16
            Set Swapper = NewDelegate(AddressOf SwapVariant)
        Case Else
            Dim TempPtr As Long
            TempPtr = CoTaskMemAlloc(Info.ElementSize)
            
            Do While Element1Ptr < Element2Ptr
                CopyMemory ByVal TempPtr, ByVal Element1Ptr, Info.ElementSize
                CopyMemory ByVal Element1Ptr, ByVal Element2Ptr, Info.ElementSize
                CopyMemory ByVal Element2Ptr, ByVal TempPtr, Info.ElementSize
                Element1Ptr = Element1Ptr + Info.ElementSize
                Element2Ptr = Element2Ptr - Info.ElementSize
            Loop
            
            CoTaskMemFree TempPtr
            Exit Sub
    End Select
    
    Do While Element1Ptr < Element2Ptr
        Swapper.Invoke ByVal Element1Ptr, ByVal Element2Ptr
        Element1Ptr = Element1Ptr + Info.ElementSize
        Element2Ptr = Element2Ptr - Info.ElementSize
    Loop
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
' @include "..\DocIncludes\CorArray.Sort.html"
Public Sub Sort(ByRef Arr As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    SortKeyEx Arr, Empty, LBound(Arr), Len1D(Arr), Comparer
End Sub

''
' Sorts an array, or subportion, given a startindex and length. An optionally
' supplied comparer object can be used to compare special elements, such as
' userdefined values.
'
' @param Arr The array to sort.
' @param Index The starting index in the array to start sorting.
' @param Length The number of elements to be sorted.
' @param Comparer An IComparer object or address of a comparer callback function.
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
' @include "..\DocIncludes\CorArray.SortKey.html"
Public Sub SortKey(ByRef Keys As Variant, ByRef Items As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Keys) Then _
        Error.ArgumentNull "Keys", ArgumentNull_Array
    
    SortKeyEx Keys, Items, LBound(Keys), Len1D(Keys), Comparer
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
        
    ValidateArrayRange Keys, Index, Length, Parameter_Keys, CountName:=NameOfLength
    KeysPtr = SAPtrV(Keys)
    
    If Not IsEmpty(Items) Then
        ValidateArray Items, Parameter_Items
        
        ' Items array must match the Keys array in size and bounds.
        If (LBound(Keys) <> LBound(Items)) Or _
           (UBound(Keys) <> UBound(Items)) Then _
            Error.Argument Argument_MatchingBounds
        
        ItemsPtr = SAPtrV(Items)
        SetSortItems ItemsPtr
    End If
    
    On Error GoTo Catch
    Select Case VarType(Comparer)
        ' use an IComparer object to sort with.
        Case vbObject
            SetSortKeys KeysPtr
            
            If Comparer Is Nothing Then
                Set SortComparer = CorLib.Comparer.Default
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
                Set SortComparer = CorLib.Comparer.Default
                QuickSortGeneral Keys, Index, Index + Length - 1
            End If
        Case Else
            Error.Argument Argument_InvalidComparer, "Comparer"
    End Select
    
    GoSub Finally
    Exit Sub
    
Catch:
    GoSub Finally
    Throw
Finally:
    ClearSortKeys
    ClearSortItems
    Set SortComparer = Nothing
    Return
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Internal
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function InternalBinarySearch(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, ByRef Value As Variant, ByVal Comparer As IComparer) As Long
    On Error GoTo Catch
    
    If Comparer Is Nothing Then
        If TrySZBinarySearch(SAPtrV(Arr), Value, Index, Length, InternalBinarySearch) Then
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
                InternalBinarySearch = MiddleIndex
                Exit Function
            Case Is > 0
                HighIndex = MiddleIndex - 1
            Case Else
                LowIndex = MiddleIndex + 1
        End Select
    Loop
    
    InternalBinarySearch = (Not (LowIndex - LBound(Arr))) + LBound(Arr)
    Exit Function
    
Catch:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LengthCore(ByVal SafeArrayPtr As Long) As Long
    Dim Rank            As Long
    Dim ArrayBoundPtr   As Long
    
    Debug.Assert SafeArrayPtr <> 0
    
    Rank = SafeArrayGetDim(SafeArrayPtr)
    ArrayBoundPtr = SafeArrayPtr + vbSizeOfSafeArray
    
    ' optimize for 99% of calls
    If Rank = 1 Then
        LengthCore = MemLong(ArrayBoundPtr)
    Else
        LengthCore = CalculateProductOfDimensions(ArrayBoundPtr, Rank)
    End If
End Function

Private Function CalculateProductOfDimensions(ByVal ArrayBoundPtr As Long, ByVal Rank As Long) As Long
    Dim Ptr     As Long
    Dim Result  As Long
    
    Result = 1
    For Ptr = ArrayBoundPtr To ArrayBoundPtr + (vbSizeOfSafeArrayBound * (Rank - 1)) Step vbSizeOfSafeArrayBound
        Result = Result * MemLong(Ptr)
    Next

    CalculateProductOfDimensions = Result
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
            Set rec = SafeArrayGetRecordInfo(SAPtrV(Arr))
            
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
Private Sub CopyValueTypes(ByVal SrcSAPtr As Long, ByVal SourceIndex As Long, ByVal DstSAPtr As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim ElementSize     As Long
    Dim SrcElementPtr   As Long
    Dim DstElementPtr   As Long
    
    ElementSize = SafeArrayGetElemsize(SrcSAPtr)
    SrcElementPtr = MemLong(SrcSAPtr + PVDATA_OFFSET) + (SourceIndex - SafeArrayGetLBound(SrcSAPtr, 1)) * ElementSize
    DstElementPtr = MemLong(DstSAPtr + PVDATA_OFFSET) + (DestinationIndex - SafeArrayGetLBound(DstSAPtr, 1)) * ElementSize
    CopyMemory ByVal DstElementPtr, ByVal SrcElementPtr, Length * ElementSize
End Sub

' Copies reference types, dealing with possible overlapping source and destination regions.
'
Private Sub CopyReferenceTypes(ByVal srcpSA As Long, ByVal SourceIndex As Long, ByVal dstpSA As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim srcSA       As GuidSafeArray1d
    Dim dstSA       As GuidSafeArray1d
    Dim IIDOffset   As Long
    
    Select Case SafeArrayGetVartype(srcpSA)
        Case vbObject:              IIDOffset = vbSizeOfGuid
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

Private Sub CopyWithCast(ByVal SrcSAPtr As Long, ByVal SourceIndex As Long, ByVal DstSAPtr As Long, ByVal DestinationIndex As Long, ByVal Length As Long, ByVal pfn As Long)
    Dim SrcSize     As Long
    Dim Src         As Long
    Dim DstSize     As Long
    Dim Dst         As Long
    Dim Assigner    As Action_T_T
    
    SrcSize = SafeArrayGetElemsize(SrcSAPtr)
    Src = MemLong(SrcSAPtr + PVDATA_OFFSET) + (SourceIndex - SafeArrayGetLBound(SrcSAPtr, 1)) * SrcSize
    DstSize = SafeArrayGetElemsize(DstSAPtr)
    Dst = MemLong(DstSAPtr + PVDATA_OFFSET) + (DestinationIndex - SafeArrayGetLBound(DstSAPtr, 1)) * DstSize
    Set Assigner = NewDelegate(pfn)
    
    Do While Length > 0
        Assigner.Invoke ByVal Dst, ByVal Src
        Src = Src + SrcSize
        Dst = Dst + DstSize
        Length = Length - 1
    Loop
End Sub

' Copies objects and checks for compatible interfaces during the copy.
'
Private Sub CopyObjectTypesCast(ByVal SrcSAPtr As Long, ByVal SourceIndex As Long, ByVal DstSAPtr As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Const SizeOfPointer As Long = 4
    
    Dim SrcLowerBound   As Long
    Dim DstLowerBound   As Long
    Dim SrcElementPtr   As Long
    Dim DstElementPtr   As Long
    Dim DstGuidPtr      As Long
    
    SrcLowerBound = SafeArrayGetLBound(SrcSAPtr, 1)
    DstLowerBound = SafeArrayGetLBound(DstSAPtr, 1)
    SrcElementPtr = MemLong(SrcSAPtr + PVDATA_OFFSET) + (SourceIndex - SrcLowerBound) * SizeOfPointer
    DstElementPtr = MemLong(DstSAPtr + PVDATA_OFFSET) + (DestinationIndex - DstLowerBound) * SizeOfPointer
    DstGuidPtr = DstSAPtr - vbSizeOfGuid
    
    Dim pObj    As Long
    Dim SrcObj  As IVBUnknown
    Dim DstObj  As IVBUnknown
    Do While Length > 0
        ObjectPtr(SrcObj) = MemLong(SrcElementPtr)
        pObj = 0
        
        If Not SrcObj Is Nothing Then
            If SrcObj.QueryInterface(ByVal DstGuidPtr, pObj) = E_NOINTERFACE Then
                ObjectPtr(SrcObj) = vbNullPtr
                ObjectPtr(DstObj) = vbNullPtr
                Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
            End If
        End If
        
        ObjectPtr(DstObj) = MemLong(DstElementPtr)
        If Not DstObj Is Nothing Then
            DstObj.Release
        End If
        MemLong(DstElementPtr) = pObj
        
        SrcElementPtr = SrcElementPtr + SizeOfPointer
        DstElementPtr = DstElementPtr + SizeOfPointer
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

Private Sub GetArrayInfo(ByRef Arr As Variant, ByRef Info As ArrayInfo)
    Info.SAPtr = SAPtrV(Arr)
    Info.DataPtr = MemLong(Info.SAPtr + PVDATA_OFFSET)
    Info.ElementSize = SafeArrayGetElemsize(Info.SAPtr)
End Sub


