VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArray"
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
'    Module: cArray
'

''
' Provides methods for manipulating, searching and sorting arrays.
'
' @remarks This class cannot be directly instantiated. To access the methods
' use the class name directly.
' <pre>
' Debug.Print cArray.GetLength(MyArray)
' </pre>
'
Option Explicit


' extended safearray descriptor the includes the 16 bytes
' preceding the array descriptor. The 16 bytes contains
' a GUID for objects and user-defined types. This is to
' help in the correct copying of those types of elements.
Private Type GuidSafeArray1d
    Guid As VBGUID
    SA As SafeArray1d
End Type

''
' Represents the available array types that can be created using
' CreateInstance. The values corrispond the vbVarType enum.
'
' @param ciLong Same as vbLong
' @param ciInteger Same as vbInteger
' @param ciByte Same as vbByte
' @param ciDouble Same as vbDouble
' @param ciSingle Same as vbSingle
' @param ciBoolean Same as vbBoolean
' @param ciDate Same as vbDate
' @param ciString Same as vbString
' @param ciCurrency Same as vbCurrency
' @param ciVariant Same as vbVariant
' @param ciObject Same as vbObject
'
Public Enum ciArrayTypes
    ciLong = vbLong
    ciInteger = vbInteger
    ciByte = vbByte
    ciDouble = vbDouble
    ciSingle = vbSingle
    ciBoolean = vbBoolean
    ciDate = vbDate
    ciString = vbString
    ciCurrency = vbCurrency
    ciVariant = vbVariant
    ciObject = vbObject
End Enum



Private mBinarySearchDelegator  As FunctionDelegator
Private mBinarySearchComparer   As ITwoArgReturnLong

Private mIndexOfDelegator       As FunctionDelegator
Private mIndexOfComparer        As ITwoArgReturnBool 'ITwoRefReturnBool

Private mAssignDelegator        As FunctionDelegator
Private mAssigner               As ITwoArgReturnVoid

Private mSortDelegator          As FunctionDelegator
Private mSorter                 As ISortRoutine

Private mFindDelegator          As FunctionDelegator
Private mFinder                 As IOneArgReturnBool

Private mForEachDelegator       As FunctionDelegator
Private mForEach                As IOneArgReturnVoid

Private CallbackSorter          As New CallbackSorter


''
' Uses a callback method to check if an element in an array exists.
'
' @param arr The array to check if an element exists according to the callback criteria.
' @param AddressOfMatch The AddressOf the callback method used to determine if
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
Public Function Exists(ByRef Arr As Variant, ByVal AddressOfMatch As Long) As Boolean
    Exists = (FindIndex(Arr, AddressOfMatch) >= LBound(Arr))
End Function

''
' Uses a callback method to search an array for the first element that matches the criteria.
'
' @param arr The array to search.
' @param AddressOfMatch The callback method address used to determine if an element
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
' @include "..\Includes\cArray.Find.txt"
Public Function Find(ByRef Arr As Variant, ByVal AddressOfMatch As Long) As Variant
    Dim Index As Long
    
    Index = FindIndex(Arr, AddressOfMatch)
    If Index >= LBound(Arr) Then
        Call VariantCopyInd(Find, Arr(Index))
    Else
        Call AssignDefaultElementValue(Arr, Find)
    End If
End Function

''
' Finds the index of the first array element that matches the criteria
' specified by the callback function.
'
' @param arr The array to be searched.
' @param AddressOfMatch The address of the callback function used to match the array elements.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindIndex(ByRef Arr As Variant, ByVal AddressOfMatch As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    Dim Result      As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Result = GetOptionalArrayRange(pSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    
    mFindDelegator.pfn = AddressOfMatch
    
    Dim i As Long
    For i = 1 To ElemCount
        If mFinder.Call(pElem) = True Then
            FindIndex = LBound(Arr) + i - 1
            Exit Function
        End If
        pElem = pElem + ElemSize
    Next i
    FindIndex = LBound(Arr) - 1
End Function

''
' Finds all matching elements in an array, returning an array of the matched elements.
'
' @param arr The array to find all matching elements.
' @param AddressOfMatch The callback address of the matching criteria method.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindAll(ByRef Arr As Variant, ByVal AddressOfMatch As Long) As Variant
    mFindDelegator.pfn = AddressOfMatch
    
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr, True)
    
    Dim LowerBound As Long
    LowerBound = SafeArrayGetLBound(pSafeArray, 1)
    
    Dim Count As Long
    Count = SafeArrayGetUBound(pSafeArray, 1) - LowerBound + 1
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET)
    
    
    ' iterate each array element, passing the memory address
    ' to the user supplied AddressOf callback. This will appear
    ' as a ByRef parameter of the array's datatype in the callback.
    Dim i                   As Long
    Dim MatchCount          As Long
    Dim MatchedIndexes()    As Long
    
    ReDim MatchedIndexes(0 To 31)
    For i = LowerBound To LowerBound + Count - 1
        If mFinder.Call(pElem) = True Then
            If MatchCount > UBound(MatchedIndexes) Then ReDim Preserve MatchedIndexes(0 To MatchCount * 2 - 1)
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
            MemLong(VarPtr(Ret) + VARIANTDATA_OFFSET) = SafeArrayCreateVectorEx(vbObject, 0, MatchCount, pSafeArray - SIZEOF_GUID)
            'If MatchCount > 0 Then ReDim ret(0 To MatchCount - 1)
            
            For i = 0 To MatchCount - 1
                Set Ret(i) = Arr(MatchedIndexes(i))
            Next i
            
        Case vbUserDefinedType
            Dim UDT As IRecordInfo
            Set UDT = SafeArrayGetRecordInfo(pSafeArray)    ' gets an object used to manipulate a UDT.
            
            Dim pvData  As Long
            pvData = MemLong(pSafeArray + PVDATA_OFFSET)
            
            Dim pReturnSafeArray As Long
            pReturnSafeArray = SafeArrayCreateVectorEx(vbUserDefinedType, 0, MatchCount, ObjPtr(UDT))
            
            Dim pDstElem As Long
            pDstElem = MemLong(pReturnSafeArray + PVDATA_OFFSET)
            
            For i = 0 To MatchCount - 1
                Call UDT.RecordCopy(pvData + (MatchedIndexes(i) - LowerBound) * ElemSize, pDstElem)
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
' @param AddressOfMatch A callback address of the matching criteria method.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindLast(ByRef Arr As Variant, ByVal AddressOfMatch As Long) As Variant
    Dim Index As Long
    
    Index = FindLastIndex(Arr, AddressOfMatch)
    If Index >= LBound(Arr) Then
        Call VariantCopyInd(FindLast, Arr(Index))
    Else
        Call AssignDefaultElementValue(Arr, FindLast)
    End If
End Function

''
' Finds the index of the last occurence of a matched element in the array.
'
' @param arr The array to search.
' @param AddressOfMatch The callback address of the matching criteria method.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindLastIndex(ByRef Arr As Variant, ByVal AddressOfMatch As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    mFindDelegator.pfn = AddressOfMatch
    
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    Dim Result      As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Result = GetOptionalArrayRangeReverse(pSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    
    Dim i As Long
    For i = ElemCount To 1 Step -1
        If mFinder.Call(pElem) = True Then
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
' @param AddressOfMatch The callback address used to match the criteria.
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
Public Function TrueForAll(ByRef Arr As Variant, ByVal AddressOfMatch As Long) As Boolean
    mFindDelegator.pfn = AddressOfMatch
    
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr, True)
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If mFinder.Call(pElem) = False Then Exit Function
        pElem = pElem + ElemSize
    Next i
    TrueForAll = True
End Function

''
' Iterates over an array passing in each element into an Action function to be performed on the element.
'
' @param Arr The array containing the elements to be processed.
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
    mForEachDelegator.pfn = AddressOfAction
    
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr, True)
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Call mForEach.Call(pElem)
        pElem = pElem + ElemSize
    Next i
End Sub

''
' Performs a binary search on a given array. A subportion of the array can
' be searched using the startindex and length parameters. A custom user
' comparer can optionally be supplied to perform special comparisons
' between elements in the array.
'
' @param arr The array to search for a specific value.
' @param value The value to search for in the array.
' @param Index The starting index in the array to begin searching.
' @param Count The number of elements to search, starting at startindex.
' @param comparer A user supplied object to compare elements within the array.
' @return The index at which the value was found.
' @remarks If the returned value is less than the lower bound of the array, then
' the value indicates where the value would have been found in the array. The
' following is how to convert the return value to the array index:<br><br>
' Lowerbound of 0:  return = Not return<br>
' Other lowerbound: return = (Not (return - lowerbound)) + lowerbound<br>
' @include "..\Includes\cArray_BinarySearch.txt"
Public Function BinarySearch(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalArrayRange(pSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
        
    On Error GoTo errTrap
    If Comparer Is Nothing Then
        If TrySZBinarySearch(pSafeArray, Value, ElemIndex, ElemCount, BinarySearch) Then Exit Function
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    ' perform a search using comparer.
    Dim MiddleIndex As Long
    Dim LowIndex    As Long
    Dim HighIndex   As Long
    
    LowIndex = ElemIndex
    HighIndex = LowIndex + ElemCount - 1
    Do While LowIndex <= HighIndex
        MiddleIndex = (LowIndex + HighIndex) \ 2
        Select Case Comparer.Compare(Arr(MiddleIndex), Value)
            Case 0
                BinarySearch = MiddleIndex
                Exit Function
            Case Is > 0
                HighIndex = MiddleIndex - 1
            Case Else
                LowIndex = MiddleIndex + 1
        End Select
    Loop
    
    Dim LowerBound As Long
    LowerBound = SafeArrayGetLBound(pSafeArray, 1)
    BinarySearch = (Not (LowIndex - LowerBound)) + LowerBound
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Clears a portion of the elements in an array.
'
' @param arr The array to clear elements from.
' @param Index The starting element to being clearing.
' @param Count The number of elements to be cleared.
'
Public Sub Clear(ByRef Arr As Variant, ByVal Index As Long, ByVal Count As Long)
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    Dim Result As Long
    Result = VerifyArrayRange(pSafeArray, Index, Count)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", Index, "Index", Count, "Count")
    
    If Count = 0 Then Exit Sub
    
    Dim IIDOffset As Long
    Select Case VarType(Arr) And &HFF   ' knock off the vbArray bit
        Case vbObject
            IIDOffset = SIZEOF_GUID
            
        Case vbString, vbUserDefinedType
            IIDOffset = 4
            
        Case Else
            ' we can use ZeroMemory for everything else.
            Dim ElemSize As Long
            ElemSize = SafeArrayGetElemsize(pSafeArray)
            
            Dim pElem As Long
            pElem = MemLong(pSafeArray + PVDATA_OFFSET)
            
            ' optimize by zeroing out the memory directly.
            Call ZeroMemory(ByVal pElem + (Index - LBound(Arr)) * ElemSize, Count * ElemSize)
            Exit Sub
    End Select
    
    Dim SubArray As GuidSafeArray1d
    Call CopyMemory(ByVal VarPtr(SubArray.SA) - IIDOffset, ByVal pSafeArray - IIDOffset, Len(SubArray.SA) + IIDOffset)
    
    With SubArray.SA
        .fFeatures = .fFeatures Or FADF_FIXEDSIZE Or FADF_STATIC Or FADF_AUTO
        .pvData = .pvData + (Index - LBound(Arr)) * .cbElements
        .cElements = Count
    End With
    Call SafeArrayDestroyData(VarPtr(SubArray.SA))
End Sub

''
' Copies one Array to another Array and performs type casting as necessary.
'
' @param SourceArray The array from which to copy the elements.
' @param DestinationArray The array in which to place the elements.
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
Public Sub Copy(ByRef SourceArray As Variant, ByRef DestinationArray As Variant, ByVal Length As Long, Optional ByVal AddressOfCopier As Long)
    If IsNull(SourceArray) Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "SourceArray")
    If IsNull(DestinationArray) Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "DestinationArray")
    
    InternalCopy SourceArray, LBound(SourceArray), DestinationArray, LBound(DestinationArray), Length, AddressOfCopier
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
Public Sub CopyEx(ByRef SourceArray As Variant, ByVal SourceIndex As Long, ByRef DestinationArray As Variant, ByVal DestinationIndex As Long, ByVal Length As Long, Optional ByVal AddressOfCopier As Long = 0)
    InternalCopy SourceArray, SourceIndex, DestinationArray, DestinationIndex, Length, AddressOfCopier
End Sub

''
' Initializes a new array of the specified type in up to 3 dimensions.
'
' @param ArrayType The data type the array represents.
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
Public Function CreateInstance(ByVal ArrayType As ciArrayTypes, Optional ByVal length1 As Long = 0, Optional ByRef length2 As Variant, Optional ByRef length3 As Variant) As Variant
    Dim bounds(2)   As SafeArrayBound
    Dim Rank        As Long
    Dim Length      As Long
    
    Rank = 1
    If length1 < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "length1", length1)
    bounds(2).cElements = length1
    
    If Not IsMissing(length2) Then
        Rank = 2
        Length = length2
        If Length < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "length2", length2)
        If Length = 0 And bounds(2).cElements <> 0 Then _
            Throw Cor.NewArgumentException("If defining one dimension as 0, then all other defined dimensions must be 0.", "length2")
        
        bounds(1).cElements = Length
    End If
    If Not IsMissing(length3) Then
        If Rank <> 2 Then Throw Cor.NewArgumentException("Cannot specifiy dimension without dimension 2.", "length3")
        Rank = 3
        Length = length3
        If Length < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "length3", length3)
        If Length = 0 And bounds(1).cElements <> 0 Then _
            Throw Cor.NewArgumentException("If defining one dimension as 0, then all other defined dimensions must be 0.", "length2")
        
        bounds(0).cElements = Length
    End If
        
    Dim Ret As Variant
    VariantType(Ret) = ArrayType Or vbArray
    MemLong(VarPtr(Ret) + VARIANTDATA_OFFSET) = SafeArrayCreate(ArrayType, Rank, bounds(3 - Rank))
    
    ' Elements of a size that are a multiple of 4 will not create a new array
    ' when being returned. All other datatype sizes will cause a new array
    ' to be created. The array must also be assigned at the end of the code
    ' so VB can optimize returning the Variant datatype and not copy it when possible.
    CreateInstance = Ret
End Function

''
' Creates an array of values.
'
' @param arraytype The target datatype of the resultant array and the values to be converted to.
' @param arrayelements The values to be put into the new array.
' @return An array of the specified datatype and containing the array elements with conversion.
' Or an empty array if no elements are defined.
' <p>Elements of a size that are a multiple of 4 will not create a new array
' when being returned. All other datatype sizes will cause a new array
' to be created.</p>
'
Public Function NewArray(ByVal ArrayType As ciArrayTypes, ParamArray ArrayElements() As Variant) As Variant
    Dim Args() As Variant
    
    Call Helper.Swap4(ByVal ArrPtr(Args), ByVal Helper.DerefEBP(16))
    NewArray = InternalNewArray(ArrayType, Args)
End Function

''
' Returns the number of elements in the specified dimension. If 0 is
' specified, this returns the total number of elements in the array.
'
' @param arr The array to retrieve the number of elements from.
' @param dimension (Optional) the specific dimension to get the number of elements.
' @return The number of elements in the specified dimension or the entire array.
'
Public Function GetLength(ByRef Arr As Variant, Optional ByVal Dimension As Long = 0) As Long
    Dim pSafeArray  As Long
'    Dim Rank        As Long
    
    pSafeArray = GetArrayPointer(Arr, True)
'    Rank = SafeArrayGetDim(pSafeArray)
'    If Dimension > Rank Then _
'        Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(IndexOutOfRange_Dimension))
'
'    pSafeArray = pSafeArray + SIZEOF_SAFEARRAY
'    If Dimension < 1 Then
'        ' optimize for 99% of calls
'        If Rank = 1 Then
'            GetLength = MemLong(pSafeArray)
'            Exit Function
'        End If
'
'        ' calculate the product of all the dimension sizes.
'        GetLength = 1
'        Dim i As Long
'        For i = pSafeArray To pSafeArray + (SIZEOF_SAFEARRAYBOUND * (Rank - 1)) Step SIZEOF_SAFEARRAYBOUND
'            GetLength = GetLength * MemLong(i)
'        Next i
'    Else
'        ' get the size of a specific dimension.
'        pSafeArray = pSafeArray + (Rank - Dimension) * SIZEOF_SAFEARRAYBOUND
'        GetLength = MemLong(pSafeArray)
'    End If
    GetLength = InternalGetLength(pSafeArray, Dimension)
End Function

Private Function InternalGetLength(ByVal pSafeArray As Long, ByVal Dimension As Long) As Long
    Dim Rank As Long
    Rank = SafeArrayGetDim(pSafeArray)
    If Dimension > Rank Then _
        Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(IndexOutOfRange_Dimension))
        
    pSafeArray = pSafeArray + SIZEOF_SAFEARRAY
    If Dimension < 1 Then
        ' optimize for 99% of calls
        If Rank = 1 Then
            InternalGetLength = MemLong(pSafeArray)
            Exit Function
        End If
        
        ' calculate the product of all the dimension sizes.
        InternalGetLength = 1
        Dim i As Long
        For i = pSafeArray To pSafeArray + (SIZEOF_SAFEARRAYBOUND * (Rank - 1)) Step SIZEOF_SAFEARRAYBOUND
            InternalGetLength = InternalGetLength * MemLong(i)
        Next i
    Else
        ' get the size of a specific dimension.
        pSafeArray = pSafeArray + (Rank - Dimension) * SIZEOF_SAFEARRAYBOUND
        InternalGetLength = MemLong(pSafeArray)
    End If
End Function

''
' Returns the number of dimensions in the array.
'
' @param arr The array to retrieve the number of dimensions from.
' @return The number of dimension in the array.
Public Function GetRank(ByRef Arr As Variant) As Long
    GetRank = SafeArrayGetDim(GetArrayPointer(Arr, True))
End Function

''
' Returns the index of the first occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param arr The array to search.
' @param value The value to search for in the array.
' @param Index The index to the element in the array to begin search from.
' @param count The number of elements to search, starting from startindex.
' @param comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function IndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    Dim Result      As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Result = GetOptionalArrayRange(pSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    On Error GoTo errTrap
    If Comparer Is Nothing Then
        ' if no comparer was supplied, then try some optimized routines.
        If TrySZIndexOf(pSafeArray, Value, ElemIndex, ElemCount, IndexOf) Then Exit Function
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    ' perform a search using comparer.
    Do While ElemCount > 0
        If Comparer.Compare(Arr(ElemIndex), Value) = 0 Then
            IndexOf = ElemIndex
            Exit Function
        End If
        ElemIndex = ElemIndex + 1
        ElemCount = ElemCount - 1
    Loop
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
    IsNull = (GetArrayPointer(Arr) = vbNullPtr)
End Function

''
' Returns the index of the last occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param arr The array to search.
' @param value The value to search for in the array.
' @param Index The index to the element in the array to begin search from.
' @param Count The number of elements to search, starting from startindex.
' @param comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
'
Public Function LastIndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr, True)
    
    Dim Result      As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Result = GetOptionalArrayRangeReverse(pSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))

    On Error GoTo errTrap
    If Comparer Is Nothing Then
        ' if no comparer was supplied, then try some optimized routines.
        If TrySZLastIndexOf(pSafeArray, Value, ElemIndex, ElemCount, LastIndexOf) Then Exit Function
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    ' perform a search using comparer.
    Do While ElemCount > 0
        If Comparer.Compare(Arr(ElemIndex), Value) = 0 Then
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
' @param arr The array to reverse.
' @param Index The starting element to begin reversing elements.
' @param Count The number of elements to reverse in the array.
'
Public Sub Reverse(ByRef Arr As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    Dim Result      As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Result = GetOptionalArrayRange(pSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Arr", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pLowElem As Long
    pLowElem = MemLong(pSafeArray + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    
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
                Call CopyMemory(ByVal pSwapBuffer, ByVal pLowElem, ElemSize)
                Call CopyMemory(ByVal pLowElem, ByVal pHighElem, ElemSize)
                Call CopyMemory(ByVal pHighElem, ByVal pSwapBuffer, ElemSize)
                pLowElem = pLowElem + ElemSize
                pHighElem = pHighElem - ElemSize
            Loop
            Call CoTaskMemFree(pSwapBuffer)
    End Select
End Sub

''
' Sorts an entire array. An optionally supplied comparer object can be
' used to compare special elements, such as userdefined values.
'
' @param Keys The array to sort.
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
' @include "..\Includes\cArray_Sort.txt"
Public Sub Sort(ByRef Keys As Variant, Optional ByVal Comparer As Variant)
    Call InternalSort(Keys, Empty, , , Comparer)
End Sub

''
' Sorts an array, or subportion, given a startindex and length. An optionally
' supplied comparer object can be used to compare special elements, such as
' userdefined values.
'
' @param Keys The array to sort.
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
Public Sub SortEx(ByRef Keys As Variant, ByVal Index As Long, ByVal Count As Long, Optional ByVal Comparer As Variant)
    Call InternalSort(Keys, Empty, Index, Count, Comparer)
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
' @include "..\Includes\cArray.SortKey.txt"
Public Sub SortKey(ByRef Keys As Variant, ByRef Items As Variant, Optional ByVal Comparer As Variant)
    Call InternalSort(Keys, Items, , , Comparer)
End Sub

''
' Sorts an array based on an array of keys. A subportion of the array can
' be sorted given a startindex and length. An optionally, supplied comparer
' object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
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
Public Sub SortKeyEx(ByRef Keys As Variant, ByRef Items As Variant, ByVal Index As Long, ByVal Count As Long, Optional ByVal Comparer As Variant)
    Call InternalSort(Keys, Items, Index, Count, Comparer)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This is Friend to allow access from the Constructors.NewArray function.
'
Friend Function InternalNewArray(ByVal ArrayType As ciArrayTypes, ByRef Args() As Variant) As Variant
    Dim UpperBound  As Long
    UpperBound = UBound(Args)
    
    Dim Ret As Variant
    Ret = CreateInstance(ArrayType, UpperBound + 1)
    
    If UpperBound >= 0 Then
        Call CopyEx(Args, 0, Ret, 0, UpperBound + 1)
    End If
    
    InternalNewArray = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBinarySearchComparer = InitDelegator(mBinarySearchDelegator)
    Set mIndexOfComparer = InitDelegator(mIndexOfDelegator)
    Set mAssigner = InitDelegator(mAssignDelegator)
    Set mSorter = InitDelegator(mSortDelegator)
    Set mFinder = InitDelegator(mFindDelegator)
    Set mForEach = InitDelegator(mForEachDelegator)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
            Set rec = SafeArrayGetRecordInfo(GetArrayPointer(Arr))
            
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
            ObjectPtr(rec) = 0
    End Select
End Sub

''
' This is the primary copy function. It determines the type of arrays
' being used as source and destination and calls the appropriate copy method
' for any possible casting between arrays.
'
Private Sub InternalCopy(ByRef SourceArray As Variant, ByVal SourceIndex As Long, ByRef DestinationArray As Variant, ByVal DestinationIndex As Long, ByVal Length As Long, ByVal AddressOfCopier As Long)
    Dim pSourceSafeArray As Long
    pSourceSafeArray = GetArrayPointer(SourceArray)
    Dim pDestinationSafeArray As Long
    pDestinationSafeArray = GetArrayPointer(DestinationArray)
    Dim SourceLength As Long
    SourceLength = InternalGetLength(pSourceSafeArray, 0)
    Dim DestinationLength As Long
    DestinationLength = InternalGetLength(pDestinationSafeArray, 0)
    
    If pSourceSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "SourceArray")
    If pDestinationSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "DestinationArray")
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length", Length)
    If Length + SourceIndex - LBound(SourceArray) > SourceLength Then _
        Throw Cor.NewArgumentException(Resources.GetString(Arg_LongerThanSrcArray))
    If Length + DestinationIndex - LBound(DestinationArray) > DestinationLength Then _
        Throw Cor.NewArgumentException(Resources.GetString(Arg_LongerThanDestArray))
    If SourceIndex < LBound(SourceArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_ArrayLB), "SourceIndex")
    If DestinationIndex < LBound(DestinationArray) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_ArrayLB), "DestinationIndex")
    
    Dim srcArrayType As VbVarType
    srcArrayType = VarType(SourceArray) And &HFF    ' knock off the vbArray bit.
    
    Dim dstArrayType As VbVarType
    dstArrayType = VarType(DestinationArray) And &HFF   ' knock off the vbArray bit.
    
    If AddressOfCopier = 0 Then
        If srcArrayType = dstArrayType Then
            Select Case srcArrayType
                Case vbObject, vbDataObject
                    If IsEqualGUID(ByVal pSourceSafeArray - SIZEOF_GUID, ByVal pDestinationSafeArray - SIZEOF_GUID) <> BOOL_FALSE Then
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
            Dim pfn As Long
            pfn = GetAssigningFunction(srcArrayType, dstArrayType)
            If pfn = 0 Then _
                Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
            
            CopyWithCast pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length, pfn
        End If
    Else
        CopyWithCast pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length, AddressOfCopier
    End If
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
        Case vbObject:              IIDOffset = SIZEOF_GUID
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
    
    mAssignDelegator.pfn = pfn
    
    Do While Length > 0
        Call mAssigner.Call(dst, src)
        src = src + srcSize
        dst = dst + dstSize
        Length = Length - 1
    Loop
End Sub

' Copies objects and checks for compatible interfaces during the copy.
'
Private Sub CopyObjectTypesCast(ByVal srcpSA As Long, ByVal SourceIndex As Long, ByVal dstpSA As Long, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim srcLowerBound As Long
    srcLowerBound = SafeArrayGetLBound(srcpSA, 1)
    
    Dim dstLowerBound As Long
    dstLowerBound = SafeArrayGetLBound(dstpSA, 1)
    
    Dim pSrcElem As Long
    pSrcElem = MemLong(srcpSA + PVDATA_OFFSET) + (SourceIndex - srcLowerBound) * 4
    
    Dim pDstElem As Long
    pDstElem = MemLong(dstpSA + PVDATA_OFFSET) + (DestinationIndex - dstLowerBound) * 4
    
    Dim pDstGuid As Long
    pDstGuid = dstpSA - SIZEOF_GUID
    
    Dim pObj    As Long
    Dim srcObj  As IVBUnknown
    Dim dstObj  As IVBUnknown
    Do While Length > 0
        ObjectPtr(srcObj) = MemLong(pSrcElem)
        pObj = 0
        If Not srcObj Is Nothing Then
            If srcObj.QueryInterface(ByVal pDstGuid, pObj) = E_NOINTERFACE Then
                ObjectPtr(srcObj) = 0
                ObjectPtr(dstObj) = 0
                Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
            End If
        End If
        
        ObjectPtr(dstObj) = MemLong(pDstElem)
        If Not dstObj Is Nothing Then Call dstObj.Release
        MemLong(pDstElem) = pObj
        
        pSrcElem = pSrcElem + 4
        pDstElem = pDstElem + 4
        Length = Length - 1
    Loop
    
    ObjectPtr(srcObj) = 0
    ObjectPtr(dstObj) = 0
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

''
' Perform the sorting. If no comparer is supplied, then optimized sorting routines
' will be attempted, otherwise, the general routine will be used with the comparer.
Private Sub InternalSort(ByRef Keys As Variant, ByRef Items As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByRef Comparer As Variant)
    Dim pKeysSafeArray  As Long
    Dim pItemsSafeArray As Long
    Dim ElemIndex       As Long
    Dim ElemCount       As Long
    
    pKeysSafeArray = GetArrayPointer(Keys)
    
    Dim Result As Long
    Result = GetOptionalArrayRange(pKeysSafeArray, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Keys", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    ' set up items to be sorted with keys
    If IsArray(Items) Then
        pItemsSafeArray = GetArrayPointer(Items)
        If pItemsSafeArray = 0 Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Items")
        
        ' Items array must match the Keys array in size and bounds.
        If (LBound(Keys) <> LBound(Items)) Or _
           (UBound(Keys) <> UBound(Items)) Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_MatchingBounds))
        
        Call SetSortItems(pItemsSafeArray)
    End If
    
    On Error GoTo errTrap
    Call ClearException
    Do
        Select Case VarType(Comparer)
            ' use an IComparer object to sort with.
            Case vbObject
                Call SetSortKeys(pKeysSafeArray)
                
                If Comparer Is Nothing Then
                    Set SortComparer = VBCorLib.Comparer.Default
                Else
                    Set SortComparer = Comparer
                End If
                
                Call QuickSortGeneral(Keys, ElemIndex, ElemIndex + ElemCount - 1)
                Exit Do
                
            ' Use a callback method to sort with.
            ' This allows a user to supply a highspeed comarison
            ' routine to compare two elements of the array.
            Case vbLong
                Call CallbackSorter.Sort(pKeysSafeArray, pItemsSafeArray, ElemIndex, ElemIndex + ElemCount - 1, Comparer)
                Exit Do
                
            ' try optimized routines first, then fallback to the default IComparer object.
            Case vbError
                If TrySZSort(pKeysSafeArray, ElemIndex, ElemIndex + ElemCount - 1) Then Exit Do
                
                ' If we get here, then no optimized routine existed,
                ' so we'll set our comparer to the default and loop
                ' around again, letting this case statement fall
                ' into the QuickSortGeneral routine.
                Set Comparer = VBCorLib.Comparer.Default
            
            Case Else
                Throw Cor.NewArgumentException("Invalid Comparer.", "Comparer")
        End Select
    Loop
    
errTrap:
    Call ClearSortKeys
    Call ClearSortItems
    Set SortComparer = Nothing
    
    Dim Ex As Exception
    If Catch(Ex) Then Throw Ex
End Sub

Private Sub SetCopyArrayDescriptor(ByRef desc As SafeArray1d, ByVal Index As Long, ByVal Length As Long)
    With desc
        .pvData = .pvData + (Index - .lLbound) * .cbElements
        .cElements = Length
    End With
End Sub

''
' This is an optimized search routine that uses a function pointer
' to call a specific comparison routine.
Private Function SZBinarySearch(ByVal pSA As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    mBinarySearchDelegator.pfn = pfn
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSA)
    
    Dim pvData As Long
    pvData = MemLong(pSA + PVDATA_OFFSET)
    
    Dim pLowElem As Long
    pLowElem = Index - SafeArrayGetLBound(pSA, 1)
    
    Dim pHighElem As Long
    pHighElem = pLowElem + Count - 1
    
    Dim pMiddleElem As Long
    Do While pLowElem <= pHighElem
        pMiddleElem = (pLowElem + pHighElem) \ 2
        Select Case mBinarySearchComparer.Call(pvData + pMiddleElem * ElemSize, pValue)
            Case 0
                SZBinarySearch = pMiddleElem + SafeArrayGetLBound(pSA, 1)
                Exit Function
            Case Is > 0
                pHighElem = pMiddleElem - 1
            Case Else
                pLowElem = pMiddleElem + 1
        End Select
    Loop
    
    SZBinarySearch = (Not pLowElem) + SafeArrayGetLBound(pSA, 1)
End Function

Private Function SZIndexOf(ByVal pSA As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    mIndexOfDelegator.pfn = pfn
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSA)
    
    Dim pvData As Long
    pvData = MemLong(pSA + PVDATA_OFFSET)
    
    Index = Index - SafeArrayGetLBound(pSA, 1)
    Do While Count > 0
        If mIndexOfComparer.Call(pvData + Index * ElemSize, pValue) Then
            SZIndexOf = Index + SafeArrayGetLBound(pSA, 1)
            Exit Function
        End If
        Count = Count - 1
        Index = Index + 1
    Loop

    SZIndexOf = SafeArrayGetLBound(pSA, 1) - 1
End Function

Private Function SZLastIndexOf(ByVal pSA As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal pfn As Long) As Long
    mIndexOfDelegator.pfn = pfn
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSA)
    
    Dim pvData As Long
    pvData = MemLong(pSA + PVDATA_OFFSET)
    
    Index = Index - SafeArrayGetLBound(pSA, 1)
    Do While Count > 0
        If mIndexOfComparer.Call(pvData + Index * ElemSize, pValue) Then
            SZLastIndexOf = Index + SafeArrayGetLBound(pSA, 1)
            Exit Function
        End If
        Count = Count - 1
        Index = Index - 1
    Loop

    SZLastIndexOf = SafeArrayGetLBound(pSA, 1) - 1
End Function

''
' Attempt to use a specialized search for a specific data type.
Private Function TrySZBinarySearch(ByVal pSA As Long, ByRef Value As Variant, ByVal startIndex As Long, ByVal Length As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    RetVal = SZBinarySearch(pSA, VarPtr(CLng(Value)), startIndex, Length, AddressOf CompareLongs)
        Case vbString:                  RetVal = SZBinarySearch(pSA, VarPtr(StrPtr(Value)), startIndex, Length, AddressOf CompareStrings)
        Case vbDouble:                  RetVal = SZBinarySearch(pSA, VarPtr(CDbl(Value)), startIndex, Length, AddressOf CompareDoubles)
        Case vbObject, vbDataObject:    RetVal = SZBinarySearch(pSA, VarPtr(Value), startIndex, Length, AddressOf CompareIComparable)
        Case vbVariant:                 RetVal = SZBinarySearch(pSA, VarPtr(Value), startIndex, Length, AddressOf CompareVariants)
        Case vbInteger:                 RetVal = SZBinarySearch(pSA, VarPtr(CInt(Value)), startIndex, Length, AddressOf CompareIntegers)
        Case vbSingle:                  RetVal = SZBinarySearch(pSA, VarPtr(CSng(Value)), startIndex, Length, AddressOf CompareSingles)
        Case vbCurrency:                RetVal = SZBinarySearch(pSA, VarPtr(CCur(Value)), startIndex, Length, AddressOf CompareCurrencies)
        Case vbDate:                    RetVal = SZBinarySearch(pSA, VarPtr(CDate(Value)), startIndex, Length, AddressOf CompareDates)
        Case vbBoolean:                 RetVal = SZBinarySearch(pSA, VarPtr(CBool(Value)), startIndex, Length, AddressOf CompareBooleans)
        Case vbByte:                    RetVal = SZBinarySearch(pSA, VarPtr(CByte(Value)), startIndex, Length, AddressOf CompareBytes)
        Case Else: Exit Function
    End Select
    TrySZBinarySearch = True
End Function

Private Function TrySZIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal startIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZIndexOf(pSA, VarPtr(CLng(Value)), startIndex, Count, AddressOf EqualsLong)
        Case vbString:                  RetVal = SZIndexOf(pSA, VarPtr(StrPtr(Value)), startIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  RetVal = SZIndexOf(pSA, VarPtr(CDbl(Value)), startIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    RetVal = SZIndexOf(pSA, VarPtr(CDate(Value)), startIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    RetVal = SZIndexOf(pSA, VarPtr(ObjPtr(Value)), startIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 RetVal = SZIndexOf(pSA, VarPtr(Value), startIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 RetVal = SZIndexOf(pSA, VarPtr(CInt(Value)), startIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  RetVal = SZIndexOf(pSA, VarPtr(CSng(Value)), startIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    RetVal = SZIndexOf(pSA, VarPtr(CSng(Value)), startIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 RetVal = SZIndexOf(pSA, VarPtr(CBool(Value)), startIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                RetVal = SZIndexOf(pSA, VarPtr(CCur(Value)), startIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZIndexOf = True
End Function

Private Function TrySZLastIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal startIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZLastIndexOf(pSA, VarPtr(CLng(Value)), startIndex, Count, AddressOf EqualsLong)
        Case vbString:                  RetVal = SZLastIndexOf(pSA, VarPtr(StrPtr(Value)), startIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  RetVal = SZLastIndexOf(pSA, VarPtr(CDbl(Value)), startIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    RetVal = SZLastIndexOf(pSA, VarPtr(CDate(Value)), startIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    RetVal = SZLastIndexOf(pSA, VarPtr(ObjPtr(Value)), startIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 RetVal = SZLastIndexOf(pSA, VarPtr(Value), startIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 RetVal = SZLastIndexOf(pSA, VarPtr(CInt(Value)), startIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  RetVal = SZLastIndexOf(pSA, VarPtr(CSng(Value)), startIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    RetVal = SZLastIndexOf(pSA, VarPtr(CSng(Value)), startIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 RetVal = SZLastIndexOf(pSA, VarPtr(CBool(Value)), startIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                RetVal = SZLastIndexOf(pSA, VarPtr(CCur(Value)), startIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZLastIndexOf = True
End Function

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
    Call mSorter.Call(pSA, Left, Right)
    TrySZSort = True
End Function


