VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: Buffer
'

''
' This class provides methods for copying manipulating arrays as byte arrays.
'
' @remarks<br>
' Any array that is used in these methods are treated as byte arrays. No attention
' is paid to the type of the original array. They are manipulated as bytes.
'
' <br><H4>Usage:</H4><br>
' This class cannot be instantiated. It represents a set of static methods that
' are accessed through the name of the class; Buffer.*<br><br>
' <pre>
' Buffer.BlockCopy <i>srCorArray</i>, <i>sourceindex</i>, <i>dstArray</i>, <i>destinationindex</i>, <i>Count</i><br>
' </pre>
'
Option Explicit

''
' Copies the bytes from the source array to the destination array.
'
' @param Source The array to copy bytes from.
' @param SourceOffset The starting byte position to begin copying.
' @param Destination The array to receive the bytes.
' @param DestinationIndex The starting byte position to begin setting bytes.
' @param Count The number of bytes to be copied.
' @remarks <p>The arrays are treated as byte arrays, so no datatype conversion
' takes place. All indexes are in bytes, not in the datatype of the specific
' array.</p>
' <p>The lowerbound of the arrays is ignored. The first byte is at index 0.</p>
'
Public Sub BlockCopy(ByRef Source As Variant, ByVal SourceOffset As Long, ByRef Destination As Variant, ByVal DestinationOffset As Long, ByVal Count As Long)
    Dim SourcePtr As Long
    Dim DestinationPtr As Long
    
    SourcePtr = CorArray.ArrayPointer(Source)
    DestinationPtr = CorArray.ArrayPointer(Destination)
    
    If SourcePtr = vbNullPtr Then _
        Throw Cor.NewArgumentNullException("Source", Environment.GetResourceString(ArgumentNull_Array))
    If DestinationPtr = vbNullPtr Then _
        Throw Cor.NewArgumentNullException("Destination", Environment.GetResourceString(ArgumentNull_Array))
    If SourceOffset < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("SourceOffset", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If DestinationOffset < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("DestinationOffset", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
    Dim SourceLength        As Long
    Dim DestinationLength   As Long
    
    SourceLength = CorArray.Length(Source) * SafeArrayGetElemsize(SourcePtr)
    DestinationLength = CorArray.Length(Destination) * SafeArrayGetElemsize(DestinationPtr)
    
    If SourceOffset + Count > SourceLength Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    If DestinationOffset + Count > DestinationLength Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    CopyMemory ByVal MemLong(DestinationPtr + PVDATA_OFFSET) + DestinationOffset, ByVal MemLong(SourcePtr + PVDATA_OFFSET) + SourceOffset, Count
End Sub

''
' Returns the length of the array in bytes.
'
' @param Arr The array to retrieve the size in bytes from.
' @return The size of the array in bytes.
' @remarks The size is calculated by taking the number of elements
' in the array and multiply it by the number of bytes each element
' is in length.
Public Function ByteLength(ByRef Arr As Variant) As Long
    Dim ArrayPtr As Long
    ArrayPtr = CorArray.ArrayPointer(Arr)
    
    If ArrayPtr = vbNullPtr Then _
        Throw Cor.NewArgumentNullException("Arr", Environment.GetResourceString(ArgumentNull_Array))
    
    ByteLength = SafeArrayGetElemsize(ArrayPtr) * CorArray.Length(Arr)
End Function

''
' Returns a the byte value at the given byte index within an array.
'
' @param Arr The array to retrieve the byte value from.
' @param Index The byte index to read the byte at.
' @return The byte value from within the array.
' @remarks The index is based on bytes, not the size of the actual array
' elements. For example: if getting the byte from a Long array at index
' 3, then the upper 8bits of the first element will be returned.
' <p>The lowerbound of the array is ignored. Index begins at 0.</p>
'
Public Function GetByte(ByRef Arr As Variant, ByVal Index As Long) As Byte
    Dim ArrayPtr As Long
    ArrayPtr = CorArray.ArrayPointer(Arr)
    
    If ArrayPtr = vbNullPtr Then _
        Throw Cor.NewArgumentNullException("Arr", Environment.GetResourceString(ArgumentNull_Array))
    If Index < 0 Or Index >= CorArray.Length(Arr) * SafeArrayGetElemsize(ArrayPtr) Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Message:=Environment.GetResourceString(ArgumentOutOfRange_Index))

    GetByte = MemByte(MemLong(ArrayPtr + PVDATA_OFFSET) + Index)
End Function

''
' Sets the byte value at a given byte index within the array.
'
' @param arr The array to set the byte value in.
' @param index The byte index into the array to set the byte at.
' @param value The byte value to set.
' @remarks The index is based on bytes, not the size of the actual array
' elements. For example: if setting the byte of a Long array at index 3
' , then the upper 8bits of the first element will be set.
' <p>The lowerbound of the array is ignored. Index begins at 0.</p>
'
Public Sub SetByte(ByRef Arr As Variant, ByVal Index As Long, ByVal Value As Byte)
    Dim ArrayPtr As Long
    ArrayPtr = CorArray.ArrayPointer(Arr)
    
    If ArrayPtr = vbNullPtr Then _
        Throw Cor.NewArgumentNullException("Arr", Environment.GetResourceString(ArgumentNull_Array))
    If Index < 0 Or Index >= CorArray.Length(Arr) * SafeArrayGetElemsize(ArrayPtr) Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Message:=Environment.GetResourceString(ArgumentOutOfRange_Index))

    MemByte(MemLong(ArrayPtr + PVDATA_OFFSET) + Index) = Value
End Sub

