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

Option Explicit

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
    ' we use VBVM6.VariantType because VarType will cause objects
    ' that have a default property to return the default value.
    ' If the default property happens to return a numeric value
    ' it will give a false positive here.
    Select Case CorVarType(Value)
        Case vbLong, vbInteger, vbByte
            IsInteger = True
    End Select
End Function

' This method returns the VbVarType of a variable like the VarType
' method. However, this will not cause an object's default property
' to be invoked like VarType does. This method will also return
' the datatype of an array.
Public Function CorVarType(ByRef Value As Variant) As VbVarType
    CorVarType = VariantType(Value) And &HFF ' strip of the vbArray and BY_REF (&h4000) flag if it exists.
End Function

' Dereferences the 32-bits within a variant that represent
' a pointer to some other memory location.
'
' The address is held in bytes 8-11 of a 16-byte variant.
Public Function DataPtr(ByRef Value As Variant) As Long
    DataPtr = MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)
End Function

' Retrieves the 32-bits within a variant that represent
' a pointer to an IRecordInfo object.
'
' The IRecordInfo address is held in bytes 12-15 of a 16-byte variant.
Public Function RecordPtr(ByRef Value As Variant) As Long
    RecordPtr = MemLong(VarPtr(Value) + VARIANTRECORD_OFFSET)
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
