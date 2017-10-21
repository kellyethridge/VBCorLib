VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToBase64Transform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: ToBase64Transform
'

''
' A class used to transform byte data to Base 64 data.
'
' @remarks <b>ToBase64Transform</b> is used to transfrom sections of byte arrays to a Base 64 array.
' <p>This class is primarily used with a CryptoStream object to perform transformations of data
' in a streaming mechanism. It is not intended to be used directly.</p>
' @see CryptoStream
'
Option Explicit
Implements IObject
Implements ICryptoTransform


''
' Returns if this transform object can be reused.
'
' @return Returns <b>True</b> if this transform can be used multiple times, <b>False</b> otherwise.
' @remarks This always returns <b>True</b>.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = True
End Property

''
' Returns if this transform can process multiple data blocks at once.
'
' @return Returns <b>True</b> if this transform can process multiple data blocks at once, <b>False</b> otherwise.
' @remarks This always returns <b>False</b>.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = False
End Property

''
' Returns the block size of data that is operated on at a time.
'
' @return The number of bytes operated on at a time.
' @remarks This always returns 3.
' <p>When calling the TransformBlock method, the InputCount must be zero or a multiple of this value.</p>
' <p><b>ToBase64Transform</b> works on 24bits at a time using standard Base 64 conversion methods. The
' resulting output is 32bits of transformed data.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = 3
End Property

''
' Returns the block size of data that is returned from a transformation operation.
'
' @return The number of bytes that will result from a transformation.
' @remarks This always returns 4.
' <p><b>ToBase64Transform</b> works on 24bits at a time using standard Base 64 conversion methods. The
' resulting output is 32bits of transformed data.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = 4
End Property

''
' Releases all resources being used by the transform.
'
' @remarks Once all resources have been called, this transform instance will no longer work.
'
Public Sub Clear()
End Sub

''
' Converts a specified region of byte array data to Base 64 byte array data.
'
' @param InputBuffer A byte array containing the data to be transformed.
' @param InputOffset The starting position in the InputBuffer to begin transforming data.
' @param InputCount The number of bytes to be transformed.
' @param OutputBuffer A byte array that received the transformed data.
' @param OutputOffset The starting position in the OutputBuffer to begin placing transformed data.
' @return The number of bytes added to the OutputBuffer after the transformation.
' @remarks <b>ToBase64Transform</b> only operated on 3 bytes at a time, generating 4 bytes of output.
' If more than 3 bytes is needed to be transformed, then this method needs to be called multiple times.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ValidateArray InputBuffer, NameOfInputBuffer
    ValidateArray OutputBuffer, NameOfOutputBuffer
    
    If InputOffset < LBound(InputBuffer) Then _
        Error.ArgumentOutOfRange "InputOffset", ArgumentOutOfRange_ArrayLB
    If InputCount < 0 Then _
        Error.ArgumentOutOfRange "InputCount", ArgumentOutOfRange_NeedNonNegNum
    If InputOffset + 2 > UBound(InputBuffer) Then _
        Error.Argument Argument_InvalidOffLen
    If OutputOffset < LBound(OutputBuffer) Then _
        Error.ArgumentOutOfRange "OutputOffset", ArgumentOutOfRange_ArrayLB
    If OutputOffset + 3 > UBound(OutputBuffer) Then _
        Error.Argument Argument_InvalidOffLen, "OutputOffset"
    
    TransformBlockCore InputBuffer, InputOffset, OutputBuffer, OutputOffset
    
    TransformBlock = OutputBlockSize
End Function

''
' Converts a specified region of byte array data to Base 64 byte array data.
'
' @param InputBuffer A byte array containing the data to be transformed.
' @param InputOffset The starting position in the InputBuffer to begin transforming data.
' @param InputCount The number of bytes to be transformed.
' @return A byte array containing data transformed to Base 64.
' @remarks The purpose of this method is to transform a block of data that is 0 to 3 bytes in length. This is to allow
' the end of an array to be transformed if the array length is not a multiple of 3.
' <p>The resulting array will always be 4 bytes in length. It is padded if less than 3 bytes is transformed. If zero
' bytes are transformed, then a zero-length array is returned, not a null array.</p>
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    ValidateArrayRange InputBuffer, InputOffset, InputCount, NameOfInputBuffer, NameOfInputOffset, NameOfInputCount
    
    Dim OutputBuffer() As Byte
    ReDim OutputBuffer(3)
    
    Select Case InputCount
        Case 0
            OutputBuffer = Cor.NewBytes()
            
        Case 1
            OutputBuffer(0) = Base64Bytes((InputBuffer(InputOffset) And &HFC) \ 4)
            OutputBuffer(1) = Base64Bytes((InputBuffer(InputOffset) And &H3) * &H10)
            OutputBuffer(2) = vbEqualChar
            OutputBuffer(3) = vbEqualChar
        
        Case 2
            OutputBuffer(0) = Base64Bytes((InputBuffer(InputOffset) And &HFC) \ 4)
            OutputBuffer(1) = Base64Bytes(((InputBuffer(InputOffset) And &H3) * &H10) Or ((InputBuffer(InputOffset + 1) And &HF0) \ &H10))
            OutputBuffer(2) = Base64Bytes((InputBuffer(InputOffset + 1) And &HF) * 4)
            OutputBuffer(3) = vbEqualChar
    
        Case 3
            TransformBlockCore InputBuffer, InputOffset, OutputBuffer, 0
        
        Case Else
            Throw Cor.NewArgumentOutOfRangeException("InputCount", "Cannot process more than 3 byte block sizes.")
    End Select
    
    TransformFinalBlock = OutputBuffer
End Function

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compare against this instance.
' @return Returns if the the Value and this instance are equal.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return Pseudo-unique number identifying this instance.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return Returns a string representation of this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TransformBlockCore(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long)
    OutputBuffer(OutputOffset) = Base64Bytes((InputBuffer(InputOffset) And &HFC) \ 4)
    OutputBuffer(OutputOffset + 1) = Base64Bytes(((InputBuffer(InputOffset) And &H3) * &H10) Or ((InputBuffer(InputOffset + 1) And &HF0) \ &H10))
    OutputBuffer(OutputOffset + 2) = Base64Bytes(((InputBuffer(InputOffset + 1) And &HF) * 4) Or ((InputBuffer(InputOffset + 2) And &HC0) \ &H40))
    OutputBuffer(OutputOffset + 3) = Base64Bytes(InputBuffer(InputOffset + 2) And &H3F)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICryptoTransform
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ICryptoTransform_CanReuseTransform() As Boolean
    ICryptoTransform_CanReuseTransform = CanReuseTransform
End Property

Private Property Get ICryptoTransform_CanTransformMultipleBlocks() As Boolean
    ICryptoTransform_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Property Get ICryptoTransform_InputBlockSize() As Long
    ICryptoTransform_InputBlockSize = InputBlockSize
End Property

Private Property Get ICryptoTransform_OutputBlockSize() As Long
    ICryptoTransform_OutputBlockSize = OutputBlockSize
End Property

Private Function ICryptoTransform_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ICryptoTransform_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function ICryptoTransform_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    ICryptoTransform_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
