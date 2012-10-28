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
'    Module: Buffer
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
' Buffer.BlockCopy <i>srcArray</i>, <i>sourceindex</i>, <i>dstArray</i>, <i>destinationindex</i>, <i>Count</i><br>
' </pre>
'
Option Explicit

''
' Copies the bytes from the source array to the destination array.
'
' @param SrcArray The array to copy bytes from.
' @param SourceIndex The starting byte position to begin copying.
' @param DstArray The array to receive the bytes.
' @param DestinationIndex The starting byte position to begin setting bytes.
' @param Count The number of bytes to be copied.
' @remarks <p>The arrays are treated as byte arrays, so no datatype conversion
' takes place. All indexes are in bytes, not in the datatype of the specific
' array.</p>
' <p>The lowerbound of the arrays is ignored. The first byte is at index 0.</p>
'
Public Sub BlockCopy(ByRef SrcArray As Variant, ByVal SourceIndex As Long, ByRef DstArray As Variant, ByVal DestinationIndex As Long, ByVal Count As Long)
    Dim srcpSA As Long
    srcpSA = GetArrayPointer(SrcArray, True)
    
    Dim dstpSA As Long
    dstpSA = GetArrayPointer(DstArray, True)
    
    Dim Result As Long
    Result = VerifyListRange((UBound(SrcArray) - LBound(SrcArray) + 1) * SafeArrayGetElemsize(srcpSA), SourceIndex, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, SourceIndex, "SourceIndex", Count, "Count")
    
    Result = VerifyListRange((UBound(DstArray) - LBound(DstArray) + 1) * SafeArrayGetElemsize(dstpSA), DestinationIndex, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, DestinationIndex, "DestinationIndex", Count, "Count")
    
    Call CopyMemory(ByVal MemLong(dstpSA + PVDATA_OFFSET) + DestinationIndex, ByVal MemLong(srcpSA + PVDATA_OFFSET) + SourceIndex, Count)
End Sub

''
' Returns the length of the array in bytes.
'
' @param arr The array to retrieve the size in bytes from.
' @return The size of the array in bytes.
' @remarks The size is calculated by taking the number of elements
' in the array and multiply it by the number of bytes each element
' is in length.
Public Function ByteLength(ByRef Arr As Variant) As Long
    ByteLength = SafeArrayGetElemsize(GetArrayPointer(Arr, True)) * cArray.GetLength(Arr)
End Function

''
' Returns a the byte value at the given byte index within an array.
'
' @param arr The array to retrieve the byte value from.
' @param index The byte index to read the byte at.
' @return The byte value from within the array.
' @remarks The index is based on bytes, not the size of the actual array
' elements. For example: if getting the byte from a Long array at index
' 3, then the upper 8bits of the first element will be returned.
' <p>The lowerbound of the array is ignored. Index begins at 0.</p>
'
Public Function GetByte(ByRef Arr As Variant, ByVal Index As Long) As Byte
    Dim pSA As Long
    pSA = GetArrayPointer(Arr, True)
    If Index < 0 Or Index >= (UBound(Arr) - LBound(Arr) + 1) * SafeArrayGetElemsize(pSA) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "Index", Index)

    GetByte = MemByte(MemLong(pSA + PVDATA_OFFSET) + Index)
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
    Dim pSA As Long
    pSA = GetArrayPointer(Arr, True)
    If Index < 0 Or Index >= (UBound(Arr) - LBound(Arr) + 1) * SafeArrayGetElemsize(pSA) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "Index", Index)

    MemByte(MemLong(pSA + PVDATA_OFFSET) + Index) = Value
End Sub

