Attribute VB_Name = "Information"
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: Information
'

''
' Functions used to retrieve information about variables.
'
Option Explicit


' Retrieves the return value of the AddressOf method.
'
' Example:
'   pfn = FuncAddr(AddressOf MyFunction)
Public Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

Public Function SizeOf(ByVal DataType As VbVarType) As Long
    Select Case DataType
        Case vbBoolean
            SizeOf = vbSizeOfBoolean
        Case vbByte
            SizeOf = vbSizeOfByte
        Case vbInteger
            SizeOf = vbSizeOfInteger
        Case vbLong
            SizeOf = vbSizeOfLong
        Case vbCurrency
            SizeOf = vbSizeOfCurrency
        Case vbSingle
            SizeOf = vbSizeOfSingle
        Case vbDouble
            SizeOf = vbSizeOfDouble
        Case vbDate
            SizeOf = vbSizeOfDate
        Case vbDecimal
            SizeOf = vbSizeOfDecimal
        Case Else
            Error.ArgumentOutOfRange "DataType", ArgumentOutOfRange_Enum
    End Select
End Function

' simply compares the UUID of the passed in UDT against
' the UUID of the VBCorLib.Int64 UDT. Any UDT that can be
' passed as a variant will always have a UUID defined.
Public Function IsInt64(ByRef Value As Variant) As Boolean
    Dim OtherGuid As VBGUID
    
    If VarType(Value) = vbUserDefinedType Then
        OtherGuid = GetGuid(Value)
        IsInt64 = IsEqualGUID(Statics.Int64.Int64Guid, OtherGuid)
    End If
End Function

Public Function IsInt64Array(ByRef Value As Variant) As Boolean
    Dim Info As IRecordInfo
    
    Set Info = SafeArrayGetRecordInfo(SAPtrV(Value))
    IsInt64Array = IsEqualGUID(Info.GetGuid, Statics.Int64.Int64Guid)
End Function

Public Function GetGuid(ByRef Value As Variant) As VBGUID
    Dim Record As IRecordInfo
    
    Set Record = GetRecordInfo(Value)
    GetGuid = Record.GetGuid
End Function

Public Function IsPicture(ByRef Value As Variant) As Boolean
    Dim Pic As IPicture
    
    On Error GoTo Catch
    Set Pic = Value
    IsPicture = Not Pic Is Nothing
    
Catch:
End Function

Public Function IsPictureResourceGroup(ByRef Value As Variant) As Boolean
    Dim Group As PictureResourceGroup
    
    On Error GoTo Catch
    Set Group = Value
    IsPictureResourceGroup = Not Group Is Nothing
    
Catch:
End Function

Public Function IsInteger(ByRef Value As Variant) As Boolean
    Select Case CorVarType(Value)
        Case vbLong, vbInteger, vbByte
            IsInteger = True
    End Select
End Function

Public Function IsEnumerable(ByRef Value As Variant) As Boolean
    If IsArray(Value) Then
        IsEnumerable = True
    ElseIf IsObject(Value) Then
        If Not Value Is Nothing Then
            If TypeOf Value Is Collection Then
                IsEnumerable = True
            ElseIf TypeOf Value Is IEnumerable Then
                IsEnumerable = True
            End If
        End If
    End If
End Function

' This method returns the VbVarType of a variable like the VarType
' method. However, this will not cause an object's default property
' to be invoked like VarType does. This method will also return
' the datatype of an array.
Public Function CorVarType(ByRef Value As Variant) As VbVarType
    CorVarType = VariantType(Value) And &HFF ' strip off the vbArray and BY_REF (&h4000) flag if it exists.
End Function

' Dereferences the 32-bits within a variant that represent
' a pointer to some other memory location.
'
' The address is held in bytes 8-11 of a 16-byte variant.
Public Function DataPtr(ByRef Value As Variant) As Long
    DataPtr = MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)
End Function

' Returns an IRecordInfo for a UDT.
'
' This method returns Nothing if the value is not a UDT.
Public Function GetRecordInfo(ByRef Value As Variant) As IRecordInfo
    Dim Record As IUnknown
    
    If VarType(Value) = vbUserDefinedType Then
        ObjectPtr(Record) = MemLong(VarPtr(Value) + VARIANTRECORD_OFFSET)
        Set GetRecordInfo = Record
        ObjectPtr(Record) = vbNullPtr
    End If
End Function

' Retrieves a pointer to the first element of an array.
'
' This extracts the pvData value from an array's SafeArray structure.
Public Function ArrDataPtr(ByRef Value As Variant) As Long
    ArrDataPtr = MemLong(SAPtrV(Value) + PVDATA_OFFSET)
End Function

' Retrieves a pointer to an element in an array.
'
' Since this is raw memory access, the size of the datatype within the
' array must be known in order to correctly apply the index. If the
' SizeOfType parameter is zero, then the size of an element is extracted
' from within the array's SafeArray structure.
Public Function ElementPtr(ByRef Arr As Variant, ByVal Index As Long, Optional ByVal SizeOfType As Long) As Long
    Dim Ptr As Long
    
    Ptr = SAPtrV(Arr)
    
    If SizeOfType = 0 Then
        SizeOfType = MemLong(Ptr + CBELEMENTS_OFFSET)
    End If
    
    ElementPtr = MemLong(Ptr + PVDATA_OFFSET) + Index * SizeOfType
End Function

' We use this function for arrays that are passed around as
' variants. The original SAPtr function accepts an array
' variable. However, many methods accept different arrays
' of different datatypes so we need a way to dereference
' the array regardless of datatype.
Public Function SAPtrV(ByRef Value As Variant) As Long
    SAPtrV = MemLong(vbaRefVarAry(Value))
End Function

Public Function GetValidSAPtr(ByRef Value As Variant, Optional ByVal ParamName As ParameterName = NameOfArr) As Long
    If Not IsArray(Value) Then _
        Error.Argument Argument_ArrayRequired, Environment.GetParameterName(ParamName)
    
    GetValidSAPtr = MemLong(vbaRefVarAry(Value))
    
    If GetValidSAPtr = vbNullPtr Then _
        Error.ArgumentNull Environment.GetParameterName(ParamName), ArgumentNull_Array
End Function

' This is for high-speed array length length retrieval without validation.
' Only one dimensional arrays are assumed. If there are more than one dimension
' possible, this method should not be used. Instead use CorArray.Length or CorArray.GetLength.
Public Function Len1D(ByRef Arr As Variant) As Long
    Len1D = UBound(Arr) - LBound(Arr) + 1
End Function

' Returns a consistant pointer to an object regardless of which interface is used.
Public Function WeakPtr(ByVal Obj As IUnknown) As Long
    WeakPtr = ObjPtr(Obj)
End Function

' Returns a strong reference based on an object pointer.
Public Function StrongPtr(ByVal Ptr As Long) As IUnknown
    Dim Obj As IUnknown
    ObjectPtr(Obj) = Ptr
    Set StrongPtr = Obj
    ObjectPtr(Obj) = vbNullPtr
End Function
