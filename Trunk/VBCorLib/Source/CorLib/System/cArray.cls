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
' Module: cArray
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
    Guid    As VBGUID
    SA      As SafeArray1d
End Type

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
' @param arr The array to search.
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
' @include "..\Includes\cArray.Find.txt"
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
' @param arr The array to be searched.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindIndex(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = GetOptionalLong(Index, LBound(Arr))
    ElemCount = GetOptionalLong(Count, UBound(Arr) - ElemIndex + 1)
    If ElemIndex < LBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "Index")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Count), "Count")
    If ElemIndex + ElemCount - 1 > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_UBound))
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    
    mFindDelegator.pfn = AddressOfPredicate
    
    Dim i As Long
    For i = 1 To ElemCount
        If mFinder.Call(pElem) = True Then
            FindIndex = ElemIndex + i - 1
            Exit Function
        End If
        pElem = pElem + ElemSize
    Next
    FindIndex = LBound(Arr) - 1
End Function

''
' Finds all matching elements in an array, returning an array of the matched elements.
'
' @param arr The array to find all matching elements.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindAll(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Variant
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
        
    Dim LowerBound As Long
    LowerBound = LBound(Arr)
    
    Dim Count As Long
    Count = UBound(Arr) - LowerBound + 1
    
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET)
    
    
    mFindDelegator.pfn = AddressOfPredicate
    
    ' iterate each array element, passing the memory address
    ' to the user supplied AddressOf callback. This will appear
    ' as a ByRef parameter of the array's datatype in the callback.
    Dim i                   As Long
    Dim MatchCount          As Long
    Dim MatchedIndexes()    As Long
    
    ReDim MatchedIndexes(0 To 31)
    For i = LowerBound To LowerBound + Count - 1
        If mFinder.Call(pElem) = True Then
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
' @include "..\Includes\cArray.Find.txt"
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
' @param arr The array to search.
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
' @include "..\Includes\cArray.Find.txt"
Public Function FindLastIndex(ByRef Arr As Variant, ByVal AddressOfPredicate As Long, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))

    Dim ElemIndex As Long
    Dim ElemCount As Long
    ElemIndex = GetOptionalLong(Index, UBound(Arr))
    ElemCount = GetOptionalLong(Count, ElemIndex + 1)
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_ArrayBounds), "Index")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Count")
    If ElemIndex - ElemCount + 1 < LBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound))
        
    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET) + (ElemIndex - LBound(Arr)) * ElemSize
    mFindDelegator.pfn = AddressOfPredicate
    
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
    mFindDelegator.pfn = AddressOfPredicate
    
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
        
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
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))

    Dim ElemSize As Long
    ElemSize = SafeArrayGetElemsize(pSafeArray)
    
    Dim pElem As Long
    pElem = MemLong(pSafeArray + PVDATA_OFFSET)
    
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        mForEach.Call pElem
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
' @include "..\Includes\cArray_BinarySearch.txt"
Public Function BinarySearch(ByRef Arr As Variant, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    If IsNull(Arr) Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
        
    BinarySearch = BinarySearchEx(Arr, LBound(Arr), UBound(Arr) - LBound(Arr) + 1, Value, Comparer)
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
' @include "..\Includes\cArray_BinarySearch.txt"
Public Function BinarySearchEx(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If Index < LBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "Index")
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length")
    If Index + Length - 1 > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(Argument_InvalidOffLen))
        
    On Error GoTo errTrap
    If Comparer Is Nothing Then
        If TrySZBinarySearch(pSafeArray, Value, Index, Length, BinarySearchEx) Then Exit Function
        Set Comparer = VBCorLib.Comparer.Default
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
                BinarySearchEx = MiddleIndex
                Exit Function
            Case Is > 0
                HighIndex = MiddleIndex - 1
            Case Else
                LowIndex = MiddleIndex + 1
        End Select
    Loop
    
    BinarySearchEx = (Not (LowIndex - LBound(Arr))) + LBound(Arr)
    Exit Function
errTrap:
    Throw Cor.NewInvalidOperationException(Resources.GetString(InvalidOperation_Comparer_Arg, Err.Description))
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
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If Index < LBound(Arr) Then _
        Throw Cor.NewIndexOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index))
    If Length < 0 Then _
        Throw Cor.NewIndexOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Length - 1 > UBound(Arr) Then _
        Throw Cor.NewIndexOutOfRangeException(Resources.GetString(Argument_InvalidOffLen))
    
    If Length = 0 Then Exit Sub
    
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
    
    CopyEx SourceArray, LBound(SourceArray), DestinationArray, LBound(DestinationArray), Length, AddressOfCopier
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
    
    Dim SourceArrayType As vbVarType
    SourceArrayType = VarType(SourceArray) And &HFF    ' knock off the vbArray bit.
    
    Dim DestinationArrayType As vbVarType
    DestinationArrayType = VarType(DestinationArray) And &HFF   ' knock off the vbArray bit.
    
    If AddressOfCopier = 0 Then
        If SourceArrayType = DestinationArrayType Then
            Select Case SourceArrayType
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
                        Throw Cor.NewArrayTypeMismatchException(Resources.GetString(ArrayTypeMismatch_Incompatible))
                        
                    CopyReferenceTypes pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
                    
                Case Else
                    CopyValueTypes pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length
            End Select
        Else
            Dim Assigner As Long
            Assigner = GetAssigningFunction(SourceArrayType, DestinationArrayType)
            If Assigner = vbNullPtr Then _
                Throw Cor.NewArrayTypeMismatchException(Resources.GetString(ArrayTypeMismatch_Incompatible))
            
            CopyWithCast pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length, Assigner
        End If
    Else
        CopyWithCast pSourceSafeArray, SourceIndex, pDestinationSafeArray, DestinationIndex, Length, AddressOfCopier
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
Public Function CreateInstance(ByVal ElementType As vbVarType, Optional ByVal Length1 As Long = 0, Optional ByRef Length2 As Variant, Optional ByRef Length3 As Variant) As Variant
    Dim Bounds(2)   As SafeArrayBound
    Dim Rank        As Long
    
    Select Case ElementType
        Case vbByte, vbInteger, vbLong, vbDouble, vbSingle, vbBoolean, vbCurrency, vbDate, vbString, vbObject, vbVariant
        Case Else
            Throw Cor.NewArgumentException(Resources.GetString(Arg_MustBeVbVarType), "ElementType")
    End Select
    
    If Length1 < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length1", Length1)
    
    Rank = 1
    Bounds(0).cElements = Length1
    
    If Not IsMissing(Length2) Then
        If CLng(Length2) < 0 Then _
            Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length2", Length2)
        
        Rank = 2
        Bounds(1).cElements = Length2
    
        If Not IsMissing(Length3) Then
            If CLng(Length3) < 0 Then _
                Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length3", Length3)
            
            Rank = 3
            Bounds(2).cElements = Length3
        End If
    End If
        
    Dim Ret As Variant
    VariantType(Ret) = ElementType Or vbArray
    MemLong(VarPtr(Ret) + VARIANTDATA_OFFSET) = SafeArrayCreate(ElementType, Rank, Bounds(0))
    
    ' Elements of a size that are a multiple of 4 will not create a new array
    ' when being returned. All other datatype sizes will cause a new array
    ' to be created. The array must also be assigned at the end of the code
    ' so VB can optimize returning the Variant datatype and not copy it when possible.
    CreateInstance = Ret
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
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    
    GetLength = InternalGetLength(pSafeArray, Dimension)
End Function

Private Function InternalGetLength(ByVal pSafeArray As Long, ByVal Dimension As Long) As Long
    Dim Rank As Long
    Rank = SafeArrayGetDim(pSafeArray)
    If Dimension > Rank Then _
        Throw Cor.NewIndexOutOfRangeException(Resources.GetString(IndexOutOfRange_Dimension))
        
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
    Dim pSafeArray  As Long
    pSafeArray = GetArrayPointer(Arr)
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
        
    GetRank = SafeArrayGetDim(pSafeArray)
End Function

''
' Returns the index of the first occurrence of a value in a one-dimensional Array or in a portion of the Array.
'
' @param Arr The array to search.
' @param Value The value to search for in the array.
' @param StartIndex The index to the element in the array to begin search from.
' @param Count The number of elements to search, starting from startindex.
' @param Comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If a value of less that
' the lower-bounds is returned, then the value was not found in the array.
Public Function IndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    
    Dim ElemIndex As Long
    Dim ElemCount As Long
    ElemIndex = GetOptionalLong(StartIndex, LBound(Arr))
    ElemCount = GetOptionalLong(Count, UBound(Arr) - ElemIndex + 1)

    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_ArrayBounds), "StartIndex")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Count), "Count")
    If ElemIndex + ElemCount - 1 > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_UBound))
    
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
    Throw Cor.NewInvalidOperationException(Resources.GetString(InvalidOperation_Comparer_Arg, Err.Description))
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
' @param Arr The array to search.
' @param Value The value to search for in the array.
' @param StartIndex The index to the element in the array to begin search from.
' @param Count The number of elements to search, starting from startindex.
' @param Comparer An optional comparer, primarily used for User-Defined types.
' @return Value indicating the index the value was found. If no value was found,
' then one less than the lower bound of the array is returned.
'
Public Function LastIndexOf(ByRef Arr As Variant, ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = GetOptionalLong(StartIndex, UBound(Arr))
    ElemCount = GetOptionalLong(Count, ElemIndex - LBound(Arr) + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_ArrayBounds), "StartIndex")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Count), "Count")
    If ElemIndex - ElemCount + 1 < LBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound))
    
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
    Throw Cor.NewInvalidOperationException(Resources.GetString(InvalidOperation_Comparer_Arg, Err.Description))
End Function

''
' Reverses the elements in a subportion of an array.
'
' @param Arr The array to reverse.
' @param Index The starting element to begin reversing elements.
' @param Count The number of elements to reverse in the array.
'
Public Sub Reverse(ByRef Arr As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim pSafeArray As Long
    pSafeArray = GetArrayPointer(Arr)
    
    If pSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
    If SafeArrayGetDim(pSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = GetOptionalLong(Index, LBound(Arr))
    ElemCount = GetOptionalLong(Count, UBound(Arr) - ElemIndex + 1)
    
    If ElemIndex < LBound(Arr) Or ElemIndex > UBound(Arr) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "Index")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Count")
    If ElemIndex + ElemCount - 1 > UBound(Arr) Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
    
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
' @include "..\Includes\cArray_Sort.txt"
Public Sub Sort(ByRef Arr As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Arr) Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
        
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
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Arr")
        
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
' @include "..\Includes\cArray.SortKey.txt"
Public Sub SortKey(ByRef Keys As Variant, ByRef Items As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Keys) Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Keys")
    
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
    Dim pKeysSafeArray  As Long
    Dim pItemsSafeArray As Long
    
    pKeysSafeArray = GetArrayPointer(Keys)
    If pKeysSafeArray = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Keys")
    If SafeArrayGetDim(pKeysSafeArray) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
    If Index < LBound(Keys) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "Index")
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length")
    If Index + Length - 1 > UBound(Keys) Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
    
    If Not IsEmpty(Items) Then
        If Not IsArray(Items) Then _
            Throw Cor.NewArgumentException(Resources.GetString(Argument_ArrayRequired), "Items")
            
        pItemsSafeArray = GetArrayPointer(Items)
        If pItemsSafeArray = vbNullPtr Then _
            Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Items")
        If SafeArrayGetDim(pItemsSafeArray) > 1 Then _
            Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
            
        ' Items array must match the Keys array in size and bounds.
        If (LBound(Keys) <> LBound(Items)) Or _
           (UBound(Keys) <> UBound(Items)) Then _
            Throw Cor.NewArgumentException(Resources.GetString(Argument_MatchingBounds))
        
        SetSortItems pItemsSafeArray
    End If
    
    On Error GoTo errTrap
    ClearException
    Select Case VarType(Comparer)
        ' use an IComparer object to sort with.
        Case vbObject
            SetSortKeys pKeysSafeArray
            
            If Comparer Is Nothing Then
                Set SortComparer = VBCorLib.Comparer.Default
            Else
                Set SortComparer = Comparer
            End If
            
            QuickSortGeneral Keys, Index, Index + Length - 1
            
        ' Use a callback method to sort with.
        ' This allows a user to supply a highspeed comarison
        ' routine to compare two elements of the array.
        Case vbLong
            CallbackSorter.Sort pKeysSafeArray, pItemsSafeArray, Index, Index + Length - 1, Comparer
            
        ' try optimized routines first, then fallback to the default IComparer object.
        Case vbError
            If Not TrySZSort(pKeysSafeArray, Index, Index + Length - 1) Then
                ' If we get here, then no optimized routine existed,
                ' so we'll set our comparer to the default and loop
                ' around again, letting this case statement fall
                ' into the QuickSortGeneral routine.
                SetSortKeys pKeysSafeArray
                Set SortComparer = VBCorLib.Comparer.Default
                QuickSortGeneral Keys, Index, Index + Length - 1
            End If
        Case Else
            Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidComparer), "Comparer")
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
        mAssigner.Call dst, src
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
    pDstGuid = dstpSA - SIZEOF_GUID
    
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
                Throw Cor.NewArrayTypeMismatchException(Resources.GetString(ArrayTypeMismatch_Incompatible))
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
Private Function GetAssigningFunction(ByVal srcType As vbVarType, ByVal dstType As vbVarType) As Long
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
Private Function TrySZBinarySearch(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Length As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    RetVal = SZBinarySearch(pSA, VarPtr(CLng(Value)), StartIndex, Length, AddressOf CompareLongs)
        Case vbString:                  RetVal = SZBinarySearch(pSA, VarPtr(StrPtr(Value)), StartIndex, Length, AddressOf CompareStrings)
        Case vbDouble:                  RetVal = SZBinarySearch(pSA, VarPtr(CDbl(Value)), StartIndex, Length, AddressOf CompareDoubles)
        Case vbObject, vbDataObject:    RetVal = SZBinarySearch(pSA, VarPtr(Value), StartIndex, Length, AddressOf CompareIComparable)
        Case vbVariant:                 RetVal = SZBinarySearch(pSA, VarPtr(Value), StartIndex, Length, AddressOf CompareVariants)
        Case vbInteger:                 RetVal = SZBinarySearch(pSA, VarPtr(CInt(Value)), StartIndex, Length, AddressOf CompareIntegers)
        Case vbSingle:                  RetVal = SZBinarySearch(pSA, VarPtr(CSng(Value)), StartIndex, Length, AddressOf CompareSingles)
        Case vbCurrency:                RetVal = SZBinarySearch(pSA, VarPtr(CCur(Value)), StartIndex, Length, AddressOf CompareCurrencies)
        Case vbDate:                    RetVal = SZBinarySearch(pSA, VarPtr(CDate(Value)), StartIndex, Length, AddressOf CompareDates)
        Case vbBoolean:                 RetVal = SZBinarySearch(pSA, VarPtr(CBool(Value)), StartIndex, Length, AddressOf CompareBooleans)
        Case vbByte:                    RetVal = SZBinarySearch(pSA, VarPtr(CByte(Value)), StartIndex, Length, AddressOf CompareBytes)
        Case Else: Exit Function
    End Select
    TrySZBinarySearch = True
End Function

Private Function TrySZIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZIndexOf(pSA, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualsLong)
        Case vbString:                  RetVal = SZIndexOf(pSA, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  RetVal = SZIndexOf(pSA, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    RetVal = SZIndexOf(pSA, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    RetVal = SZIndexOf(pSA, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 RetVal = SZIndexOf(pSA, VarPtr(Value), StartIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 RetVal = SZIndexOf(pSA, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  RetVal = SZIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    RetVal = SZIndexOf(pSA, VarPtr(CByte(Value)), StartIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 RetVal = SZIndexOf(pSA, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                RetVal = SZIndexOf(pSA, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualsCurrency)
        Case Else: Exit Function
    End Select
    TrySZIndexOf = True
End Function

Private Function TrySZLastIndexOf(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Count As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    RetVal = SZLastIndexOf(pSA, VarPtr(CLng(Value)), StartIndex, Count, AddressOf EqualsLong)
        Case vbString:                  RetVal = SZLastIndexOf(pSA, VarPtr(StrPtr(Value)), StartIndex, Count, AddressOf EqualsString)
        Case vbDouble:                  RetVal = SZLastIndexOf(pSA, VarPtr(CDbl(Value)), StartIndex, Count, AddressOf EqualsDouble)
        Case vbDate:                    RetVal = SZLastIndexOf(pSA, VarPtr(CDate(Value)), StartIndex, Count, AddressOf EqualsDate)
        Case vbObject, vbDataObject:    RetVal = SZLastIndexOf(pSA, VarPtr(ObjPtr(Value)), StartIndex, Count, AddressOf EqualsObject)
        Case vbVariant:                 RetVal = SZLastIndexOf(pSA, VarPtr(Value), StartIndex, Count, AddressOf EqualsVariants)
        Case vbInteger:                 RetVal = SZLastIndexOf(pSA, VarPtr(CInt(Value)), StartIndex, Count, AddressOf EqualsInteger)
        Case vbSingle:                  RetVal = SZLastIndexOf(pSA, VarPtr(CSng(Value)), StartIndex, Count, AddressOf EqualsSingle)
        Case vbByte:                    RetVal = SZLastIndexOf(pSA, VarPtr(CByte(Value)), StartIndex, Count, AddressOf EqualsByte)
        Case vbBoolean:                 RetVal = SZLastIndexOf(pSA, VarPtr(CBool(Value)), StartIndex, Count, AddressOf EqualsBoolean)
        Case vbCurrency:                RetVal = SZLastIndexOf(pSA, VarPtr(CCur(Value)), StartIndex, Count, AddressOf EqualsCurrency)
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
    mSorter.Call pSA, Left, Right
    TrySZSort = True
End Function


