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
' Provides methods for creating, manipulating, searching, and sorting arrays.
'
' @remarks Arrays are a common data structure used when building applications. The <b>CorArray</b> class
' provides several methods to help manage arrays through the course of an applications lifetime.
' @see PublicFunctions
' @include "..\DocIncludes\CorArray.html"
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


''
' Determines whether the specified array contains elements that match the conditions defined by the specified predicate.
'
' @param Arr The one-dimensional array to search.
' @param AddressOfPredicate The <b>AddressOf</b> a callback method used to determine if an element matches the specified criteria.
' @return <b>True</b> if <i>Arr</i> contains one or more elements that match the conditions defined by the specified predicate callback; otherwise, <b>False</b>.
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' @see #Find
' @see #FindIndex
' @see #FindLast
' @see #FindLastIndex
' @see #FindAll
' @include "..\DocIncludes\CorArray.Exists.html"
Public Function Exists(ByRef Arr As Variant, ByVal AddressOfPredicate As Long) As Boolean
    Exists = (FindIndex(Arr, AddressOfPredicate) >= LBound(Arr))
End Function

''
' Uses a callback method to search an array for the first element that matches the criteria.
'
' @param arr The one-dimensional array to search.
' @param AddressOfPredicate The callback method address used to determine if an element matches the specified criteria.
' @return The element value if a match is found, otherwise a default for that element type.
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' @see #Exists
' @see #FindIndex
' @see #FindLast
' @see #FindLastIndex
' @see #FindAll
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
' Searches for an element that matches the conditions defined by the specified callback method,
' and returns the lower-bound based index of the first occurrence within the range of elements
' in the array that starts at the specified index and contains the specified number of elements.
'
' @param Arr The one-dimensional array to be searched.
' @param AddressOfPredicate The address of the callback function used to match the array elements.
' @param Index The starting index of the search.
' @param Count The number of elements in the section to search.
' @return The index of the first occurrence of an element that matches the conditions
' defined in the callback method, if found; otherwise the lower-bound of the array minus one.
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' <p>The search is based on the array lower-bound, not zero-based.</p>
' @see #Exists
' @see #Find
' @see #FindLast
' @see #FindLastIndex
' @see #FindAll
' @include "..\DocIncludes\CorArray.FindIndex.html"
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
' Retrieves all the elements that match the conditions defined by the specified callback method.
'
' @param Arr The one-dimensional array to find all matching elements.
' @param AddressOfPredicate The callback address of the matching criteria method.
' @return An array of all matched elements.
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' <p>The returned array is of the same type as <i>Arr</i>.</p>
' @see #Exists
' @see #Find
' @see #FindIndex
' @see #FindLast
' @see #FindLastIndex
' @include "..\DocIncludes\CorArray.FindAll.html"
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
' Searches for an element that matches the conditions defined by the specified callback method, and returns the last occurrence within the entire array.
'
' @param arr The array to search.
' @param AddressOfPredicate A callback address of the matching criteria method.
' @return The last element in the array to match the criteria, or the default for the array type if no elements matched.
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' @see #Exists
' @see #Find
' @see #FindIndex
' @see #FindLastIndex
' @see #FindAll
' @include "..\DocIncludes\CorArray.FindLast.html"
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
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' @see #Exists
' @see #Find
' @see #FindIndex
' @see #FindLast
' @see #FindAll
' @include "..\DocIncludes\CorArray.FindLastIndex.html"
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
' Determines whether every element in the array matches the conditions defined by the specified callback method.
'
' @param Arr The array to check against the conditions.
' @param AddressOfPredicate The callback address used to match the criteria.
' @Return <b>True</b> if every element in array matches the conditions defined by the specified predicate; otherwise, <b>False</b>.
' If there are no elements in the array, the return value is <b>True</b>.
' @remarks The <i>AddressOfPredicate</i> is an address to a method that returns <b>True</b> if the value passed to it matches the conditions defined in the callback method.
' The elements of <i>Arr</i> are individually passed to the callback method, and processing is stopped when the method returns <b>False</b> for any element.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Function CallbackMethod(ByRef Value As &lt;Type&gt;) As Boolean
'   '' Evaluate value
' End Function
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' @include "..\DocIncludes\CorArray.TrueForAll.html"
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
' Performs the specified action on each element of the specified array.
'
' @param Arr The one-dimensional array containing the elements to be processed.
' @param AddressOfAction The callback address of the function to process an array element.
' @remarks The <i>AddressOfAction</i> is an address to a method that will perform an action on the value passed to it.
' <p>The callback method should have a signature resembling the following:</p>
' <pre>
' Public Sub CallbackMethod(ByRef Value As &lt;Type&gt;)
'   '' Perform action using Value
' End Sub
' </pre>
' <p>It has one parameter that is defined as <b>ByRef</b> and should be the same type as the elements in the array to be searched.</p>
' <p>It is extremely important to define the callback method correctly. If the method is incorrect, the application may crash.</p>
' @include "..\DocIncludes\CorArray.ForEach.html"
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
' Searches the entire one-dimensional sorted array for a specific element, using an optionally specified <a href="IComparer.html">IComparer</a> interface.
'
' @param Arr The sorted one-dimensional array to search.
' @param Value The value to search for in the array.
' @param Comparer An object used to compare elements within the array. If this is <b>Nothing</b> then <a href="ComparerStatic.Get.Default.html">Comparer.Default</a> is used.
' @return The index of the specified <i>Value</i> in the specified array, if <i>Value</i> is found. If <i>Value</i> is not found and <i>Value</i> is less than one or more
' elements in <i>Arr</i>, a negative number which is the bitwise complement of the index of the first element that is larger than <i>Value</i>. If <i>Value</i> is not found
' and <i>Value</i> is greater than any of the elements in <i>Arr</i>, a negative number which is the bitwise complement of (the index of the last element plus 1).
' @remarks The <i>Arr</i> must be sorted before calling this method.
' <p>If the returned value is less than the lower bound of the array, the value is a bitwise <b>Not</b> indication of where the value would have been found in the array.
' The following shows how to convert the result for a value that is not found.
' <table style="width: auto">
'   <tr>
'       <th>Lower Bound of <i>Arr</i></th><th>Conversion Method</th>
'   </tr>
'   <tr>
'       <td>Zero</td><td>Index = Not Result</td>
'   </tr>
'   <tr>
'       <td>Non-zero</td><td>Index = (Not (Result - LBound(<i>Arr</i>))) + LBound(<i>Arr</i>)</td>
'   </tr>
' </table>
' @see #BinarySearchEx
' @include "..\DocIncludes\CorArray.BinarySearch.html"
Public Function BinarySearch(ByRef Arr As Variant, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ValidateArray Arr
    BinarySearch = InternalBinarySearch(Arr, LBound(Arr), Len1D(Arr), Value, Comparer)
End Function

''
' Searches a range of elements in a one-dimensional sorted <i>Arr</i> for a value, using an optionally specified <a href="IComparer.html">IComparer</a> interface.
'
' @param Arr The sorted one-dimensional array to search.
' @param Index The starting index in the range to search.
' @param Length The length of the range to search.
' @param Value The value to search for in the array.
' @param Comparer The <a href="IComparer.html">IComparer</a> implementation to use when comparing elements.<br/>
' -or-<br/>
' <b>Nothing</b> to use the default <a href="ComparerStatic.Get.Default.html">Comparer</a> implementation.
' @return The index of the specified <i>Value</i> in the specified array, if <i>Value</i> is found. If <i>Value</i> is not found and <i>Value</i> is less than one or more
' elements in <i>Arr</i>, a negative number which is the bitwise complement of the index of the first element that is larger than <i>Value</i>. If <i>Value</i> is not found
' and <i>Value</i> is greater than any of the elements in <i>Arr</i>, a negative number which is the bitwise complement of (the index of the last element plus 1).
' @remarks The <i>Arr</i> must be sorted before calling this method.
' @remarks If the returned value is less than the lower bound of the array, then
' <p>If the returned value is less than the lower bound of the array, the value is a bitwise <b>Not</b> indication of where the value would have been found in the array.
' The following shows how to convert the result for a value that is not found.
' <table style="width: auto">
'   <tr>
'       <th>Lower Bound of <i>Arr</i></th><th>Conversion Method</th>
'   </tr>
'   <tr>
'       <td>Zero</td><td>Index = Not Result</td>
'   </tr>
'   <tr>
'       <td>Non-zero</td><td>Index = (Not (Result - LBound(<i>Arr</i>))) + LBound(<i>Arr</i>)</td>
'   </tr>
' </table>
' @see #BinarySearch
' @include "..\DocIncludes\CorArray.BinarySearchEx.html"
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Arr</i> is uninitialized.</td>
'   </tr>
'    <tr>
'        <td><a href="IndexOutOfRangeException.html">IndexOutOfRangeException</a></td>
'        <td>
'           <i>Index</i> is less than the lower-bound of <i>Arr</i>.<br/>
'           -or-<br/>
'           <i>Length</i> is less than zero.<br/>
'           -or-<br/>
'           The sum of <i>Index</i> and <i>Length</i> is greater than the upper-bound of <i>Arr</i>.
'        </td>
'    </tr>
' </table>
Public Sub Clear(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long)
    Dim pSafeArray  As Long
    Dim SubArray    As GuidSafeArray1d
    Dim IIDOffset   As Long
    
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
    
    Select Case VarType(Arr) And &HFF   ' knock off the vbArray bit
        Case vbObject
            IIDOffset = vbSizeOfGuid
        Case vbString, vbUserDefinedType
            IIDOffset = 4
        Case Else
            IIDOffset = 0 ' We don't need to set this, but we like explicit defaults here.
    End Select
    
    CopyMemory ByVal VarPtr(SubArray.SA) - IIDOffset, ByVal pSafeArray - IIDOffset, Len(SubArray.SA) + IIDOffset
    
    With SubArray.SA
        .fFeatures = .fFeatures Or FADF_FIXEDSIZE Or FADF_STATIC Or FADF_AUTO
        .PVData = .PVData + (Index - LBound(Arr)) * .cbElements
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td>
'        <td>
'           <i>SourceArray</i> is <b>null</b>.<br/>
'           -or-<br/>
'           <i>DestinationArray</i> is <b>null</b>.
'        </td>
'   </tr>
'    <tr>
'        <td><a href="RankException.html">RankException</a></td><td><i>SourceArray</i> and <i>DestinationArray</i> have different ranks.</td>
'    </tr>
'    <tr>
'        <td><a href="ArrayTypeMismatchException.html">ArrayTypeMismatchException</a></td><td><i>SourceArray</i> and <i>DestinationArray</i> are incompatible types.</td>
'    </tr>
'    <tr>
'        <td><a href="InvalidCastException.html">InvalidCastException</a></td><td>At least one element in <i>SourceArray</i> cannot be cast to the type of <i>DestinationArray</i>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Length</i> is less than zero.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td>
'        <td>
'           <i>Length</i> is greater than the number of elements from <i>SourceIndex</i> to the end of <i>SourceArray</i>.<br/>
'           -or-<br/>
'           <i>Length</i> is greater than the number of elements from <i>DestinationIndex</i> to the end of <i>DestinationArray</i>.
'        </td>
'    </tr>
' </table>
' @remarks The <i>SourceArray</i> and <i>DestinationArray</i> parameters must have the same number of dimensions.
' In addition, <i>DestinationArray</i> must already have been dimensioned and must have a sufficient number of elements to accommodate the copied data.
' <p>When copying between multidimensional arrays, the array behaves like a long one-dimensional array, where the rows (or columns) are conceptually
' laid end to end. For example, if an array has three rows (or columns) with four elements each, copying six elements from the beginning of the array
' would copy all four elements of the first row (or column) and the first two elements of the second row (or column).</p>
' @include "..\DocIncludes\CorArray.html"
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td>
'        <td>
'           <i>SourceArray</i> is <b>null</b>.<br/>
'           -or-<br/>
'           <i>DestinationArray</i> is <b>null</b>.
'        </td>
'   </tr>
'    <tr>
'        <td><a href="RankException.html">RankException</a></td><td><i>SourceArray</i> and <i>DestinationArray</i> have different ranks.</td>
'    </tr>
'    <tr>
'        <td><a href="ArrayTypeMismatchException.html">ArrayTypeMismatchException</a></td><td><i>SourceArray</i> and <i>DestinationArray</i> are incompatible types.</td>
'    </tr>
'    <tr>
'        <td><a href="InvalidCastException.html">InvalidCastException</a></td><td>At least one element in <i>SourceArray</i> cannot be cast to the type of <i>DestinationArray</i>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>SourceIndex</i> is less than the lower-bound of the first dimension of <i>SourceArray</i>.<br/>
'           -or-<br/>
'           <i>DestinationIndex</i> is less than the lower-bound of the first dimension of <i>DestinationArray</i>.<br/>
'           -or-<br/>
'           <i>Length</i> is less than zero.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td>
'        <td>
'           <i>Length</i> is greater than the number of elements from <i>SourceIndex</i> to the end of <i>SourceArray</i>.<br/>
'           -or-<br/>
'           <i>Length</i> is greater than the number of elements from <i>DestinationIndex</i> to the end of <i>DestinationArray</i>.
'        </td>
'    </tr>
' </table>
' @remarks The <i>SourceArray</i> and <i>DestinationArray</i> parameters must have the same number of dimensions.
' In addition, <i>DestinationArray</i> must already have been dimensioned and must have a sufficient number of elements to accommodate the copied data.
' <p>When copying between multidimensional arrays, the array behaves like a long one-dimensional array, where the rows (or columns) are conceptually
' laid end to end. For example, if an array has three rows (or columns) with four elements each, copying six elements from the beginning of the array
' would copy all four elements of the first row (or column) and the first two elements of the second row (or column).</p>
' @include "..\DocIncludes\CorArray.html"
Public Sub CopyEx(ByRef SourceArray As Variant, ByVal SourceIndex As Long, ByRef DestinationArray As Variant, ByVal DestinationIndex As Long, ByVal Length As Long)
    Dim pSourceSafeArray        As Long
    Dim pDestinationSafeArray   As Long
    Dim SourceLength            As Long
    Dim DestinationLength       As Long
    Dim SourceArrayType         As VbVarType
    Dim DestinationArrayType    As VbVarType
    Dim Assigner                As Long

    pSourceSafeArray = SAPtrV(SourceArray)
    pDestinationSafeArray = SAPtrV(DestinationArray)
    
    If pSourceSafeArray = vbNullPtr Then _
        Error.ArgumentNull "SourceArray", ArgumentNull_Array
    If pDestinationSafeArray = vbNullPtr Then _
        Error.ArgumentNull "DestinationArray", ArgumentNull_Array
    
    SourceLength = LengthCore(pSourceSafeArray)
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
    If SafeArrayGetDim(pSourceSafeArray) <> SafeArrayGetDim(pDestinationSafeArray) Then _
        Error.Rank Rank_MustMatch
        
    SourceArrayType = VarType(SourceArray) And &HFF    ' knock off the vbArray bit.
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>ElementType</i> is not a valid type.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>Length1</i> is less than zero.<br/>
'           -or-<br/>
'           <i>Length2</i> is less than zero.<br/>
'           -or-<br/>
'           <i>Length3</i> is less than zero.
'        </td>
'    </tr>
' </table>
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
' @return A 32-bit integer that represents the total number of elements in all the dimensions of <i>Arr</i>; zero if there are no elements in the array.
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
    
    Ptr = GetValidSAPtr(Arr)
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
    Dim Range   As ListRange
        
    ValidateArray Arr
    Range = MakeDefaultStepRange(StartIndex, LBound(Arr), Count, Len1D(Arr), Parameter_StartIndex)
    
    If Range.Index < LBound(Arr) Or Range.Index > UBound(Arr) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_ArrayBounds
    If Range.Count < 0 Or Range.Index + Range.Count - 1 > UBound(Arr) Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    
    On Error GoTo Catch
    
    If TrySZIndexOf(Arr, Value, Range.Index, Range.Count, IndexOf) Then
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Arr</i> is not an array.</td>
'   </tr>
' </table>
Public Function IsNull(ByRef Arr As Variant) As Boolean
    If Not IsArray(Arr) Then _
        Error.Argument Argument_ArrayRequired, "Arr"
    
    IsNull = SAPtrV(Arr) = vbNullPtr
End Function

''
' Returns if the array variable is unitialized(null) or empty(no elements).
'
' @param Arr The array to inspectd.
' @return Returns <b>True</b> if the array is unitialized or has zero elements, <b>False</b> otherwise.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Arr</i> is not an array.</td>
'   </tr>
' </table>
Public Function IsNullOrEmpty(ByRef Arr As Variant) As Boolean
    Dim Ptr As Long
    
    If Not IsArray(Arr) Then _
        Error.Argument Argument_ArrayRequired, "Arr"
    
    Ptr = SAPtrV(Arr)
    
    If Ptr = vbNullPtr Then
        IsNullOrEmpty = True
    ElseIf LengthCore(Ptr) = 0 Then
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
    
    If TrySZLastIndexOf(Arr, Value, ElemIndex, ElemCount, LastIndexOf) Then
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
' @remarks The <i>Comparer</i> parameter can be an <a href="IComparer.html">IComparer</a> interface or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As &lt;Type&gt;, ByRef y As &lt;Type&gt;) As Long
'     '' return less than zero if x is less than y.
'     '' return greater than zero if x is greater than y.
'     '' return zero if x equals y.
' End Function
' </pre>
' The &lt;Type&gt; the type of element of the array. If the
' array is an array of <b>Variants</b>, then &lt;Type&gt; would be a <b>Variant</b>, not
' any specific sub-type within the variant elements of the array.</p>
' @see IComparer
' @see #SortEx
' @see #SortKey
' @see #SortKeyEx
' @include "..\DocIncludes\CorArray.Sort.html"
Public Sub Sort(ByRef Arr As Variant, Optional ByRef Comparer As Variant)
    If IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    SortKeyEx Arr, Missing, LBound(Arr), Len1D(Arr), Comparer
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
' @remarks The <i>Comparer</i> parameter can be an <a href="IComparer.html">IComparer</a> interface or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As &lt;Type&gt;, ByRef y As &lt;Type&gt;) As Long
'     '' return less than zero if x is less than y.
'     '' return greater than zero if x is greater than y.
'     '' return zero if x equals y.
' End Function
' </pre>
' The &lt;Type&gt; the type of element of the array. If the
' array is an array of <b>Variants</>, then &lt;Type&gt; would be a <b>Variant</b>, not
' any specific sub-type within the variant elements of the array.</p>
' @see IComparer
' @see #Sort
' @see #SortKey
' @see #SortKeyEx
Public Sub SortEx(ByRef Arr As Variant, ByVal Index As Long, ByVal Length As Long, Optional ByRef Comparer As Variant)
    If IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    SortKeyEx Arr, Missing, Index, Length, Comparer
End Sub

''
' Sorts an entire array based on an array of keys. An optionally supplied
' comparer object can be used to compare special elements, such as userdefined values.
'
' @param keys An array the sorting is based on.
' @param items An array that is sorted based on the sorting of keys.
' @param Comparer (Optional) An IComparer object or address of a comparer callback function.
' @remarks The <i>Comparer</i> parameter can be an <a href="IComparer.html">IComparer</a> interface or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As &lt;Type&gt;, ByRef y As &lt;Type&gt;) As Long
'     '' return less than zero if x is less than y.
'     '' return greater than zero if x is greater than y.
'     '' return zero if x equals y.
' End Function
' </pre>
' The &lt;Type&gt; the type of element of the array. If the
' array is an array of <b>Variants</b>, then &lt;Type&gt; would be a <b>Variant</b>, not
' any specific sub-type within the variant elements of the array.</p>
' @see IComparer
' @see #Sort
' @see #SortEx
' @see #SortKeyEx
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
' @remarks The <i>Comparer</i> parameter can be an <a href="IComparer.html">IComparer</a> interface or a
' callback address to a compare function using the <b>AddressOf</b> method. The
' callback method signature is defined as follows:
' <pre>
' Public Function SortCallback(ByRef x As &lt;Type&gt;, ByRef y As &lt;Type&gt;) As Long
'     '' return less than zero if x is less than y.
'     '' return greater than zero if x is greater than y.
'     '' return zero if x equals y.
' End Function
' </pre>
' The &lt;Type&gt; the type of element of the array. If the
' array is an array of <b>Variants</b>, then &lt;Type&gt; would be a <b>Variant</b>, not
' any specific sub-type within the variant elements of the array.</p>
' @see IComparer
' @see #Sort
' @see #SortEx
' @see #SortKey
Public Sub SortKeyEx(ByRef Keys As Variant, ByRef Items As Variant, ByVal Index As Long, ByVal Length As Long, Optional ByRef Comparer As Variant)
    Dim KeysPtr     As Long
    Dim ItemsPtr    As Long
    Dim Left        As Long
    Dim Right       As Long
    
    ValidateArrayRange Keys, Index, Length, Parameter_Keys, CountName:=NameOfLength
    KeysPtr = SAPtrV(Keys)
    Left = Index
    Right = Index + Length - 1
    
    If Not IsMissing(Items) Then
        ValidateArray Items, Parameter_Items
        
        ' Items array must match the Keys array in size and bounds.
        If (LBound(Keys) <> LBound(Items)) Or _
           (UBound(Keys) <> UBound(Items)) Then _
            Error.Argument Argument_MatchingBounds
                
        ' If a callback address was passed in for the Comparer then
        ' a callback sort routine will be used and it will
        ' allocate its own sort-item swap buffer.
        If Not IsInteger(Comparer) Then
            ItemsPtr = SAPtrV(Items)
            SetSortItems ItemsPtr
        End If
    End If
    
    On Error GoTo Catch
    Select Case VarType(Comparer)
        ' use an IComparer object to sort with.
        Case vbObject
            ' use specialized sorting for different types of string comparisons.
            If (VarType(Keys) And &HFF) = vbString Then
                SortStringsWithComparer Keys, Left, Right, Comparer
                GoSub Finally
                Exit Sub
            ElseIf Comparer Is Nothing Or Comparer Is Statics.Comparer.Default Then
                ' try specialized sorting using default comparisons.
                If TrySZSort(Keys, Left, Right) Then
                    GoSub Finally
                    Exit Sub
                End If
            End If
            
            If Comparer Is Nothing Then _
                Error.ArgumentNull "Comparer"
            If Not TypeOf Comparer Is IComparer Then _
                Throw Cor.NewInvalidCastException(Environment.GetResourceString(InvalidCast_IComparer))
            
            SetSortKeys KeysPtr
            QuickSortGeneral Keys, Left, Right, CUnk(Comparer)
            
        ' Use a callback method to sort with.
        ' This allows a user to supply a highspeed comparison
        ' routine to compare two elements of the array.
        Case vbLong
            Dim Sorter As CallbackSorter
            Set Sorter = Cor.NewCallbackSorter(Comparer)
            Sorter.Sort Keys, Items, Left, Right
            
        ' No compare given so try optimized routines first, then fallback to general sort.
        Case vbMissing
            If Not TrySZSort(Keys, Left, Right) Then
                SetSortKeys KeysPtr
                QuickSortGeneral Keys, Left, Right, Statics.Comparer.Default
            End If
        Case Else
            Error.Argument Argument_InvalidComparer, "Comparer"
    End Select
    
    GoSub Finally
    Exit Sub
    
Catch:
    GoSub Finally
    ThrowOrErr Err
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
    Dim MiddleIndex As Long
    Dim LowIndex    As Long
    Dim HighIndex   As Long
        
    On Error GoTo Catch
    
    If Comparer Is Nothing Then
        If TrySZBinarySearch(Arr, Value, Index, Length, InternalBinarySearch) Then
            Exit Function
        End If
        
        Set Comparer = Statics.Comparer.Default
    End If
        
    ' perform a search using comparer.
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
                .PVData = .PVData + (SourceIndex + Length) * .cbElements
                .cElements = DestinationIndex - SourceIndex
            End With
            SafeArrayDestroyData VarPtr(ClearSA.SA)
            
            With srcSA.SA
                CopyMemory ByVal .PVData + (DestinationIndex - .lLbound) * .cbElements, ByVal .PVData + (SourceIndex - .lLbound) * .cbElements, Length * .cbElements
    
                ' If we have moved reference type values, then we need to
                ' fill in the uncovered portion of the original values
                ' with copies. First we have to zero out the uncovered portion
                ' so the copy won't release any references, since we merely
                ' moved the references without making copies to objects or strings.
                ' Once the uncovered portion is set up, we can copy backwards using
                ' the standard reference copy method without stepping on ourselves.
                ZeroMemory ByVal .PVData + (SourceIndex - .lLbound) * .cbElements, (DestinationIndex - SourceIndex) * .cbElements
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
    
    On Error GoTo Catch
    
    Do While Length > 0
        Assigner.Invoke ByVal Dst, ByVal Src
        Src = Src + SrcSize
        Dst = Dst + DstSize
        Length = Length - 1
    Loop
    
    Exit Sub
    
Catch:
    Throw Cor.NewInvalidCastException(Environment.GetResourceString(InvalidCast_DownCastArrayElement))
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
                Throw Cor.NewInvalidCastException(Environment.GetResourceString(InvalidCast_DownCastArrayElement))
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
        .PVData = .PVData + (Index - .lLbound) * .cbElements
        .cElements = Length
    End With
End Sub

Private Sub GetArrayInfo(ByRef Arr As Variant, ByRef Info As ArrayInfo)
    Info.SAPtr = SAPtrV(Arr)
    Info.DataPtr = MemLong(Info.SAPtr + PVDATA_OFFSET)
    Info.ElementSize = SafeArrayGetElemsize(Info.SAPtr)
End Sub


