VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashAlgorithmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: HashAlgorithmBase
'

''
' Provides a base implementation for implementing hash algorithms.
'
' @remarks The <b>HashAlgorithmBase</b> class provides base functionality for common
' task when computing a hash value on data within the <b>VBCorLib</b> library. This
' allows for an easier and more consistent implementation of the hash algorithms.
' <p>By implementing the <b>IHashAlgorithm</b> interface, an instance of <b>HashAlgorithmBase</b>
' will manage standard functions and validation of data. When needed, the base class will
' call back into the <b>IHashAlgorithm</b> implementation for required data processing.</p>
'
' @see Constructors
'
Option Explicit

Private mCore               As IHashAlgorithm
Private mDisposed           As Boolean
Private mHash()             As Byte
Private mHashFinished       As Boolean
Private mKeepWeakReference  As Boolean
Private mBuffer()           As Byte
Private mBlockSize          As Long
Private mBytesBuffered      As Long
Private mUseBuffer          As Boolean
Private mMessageLength      As Currency
Private mExecutingCall      As Boolean
Private mInternalCall       As Boolean



''
' Returns the computed hash.
'
' @return The compute hash value.
' @remarks Calling this property is only valid after the <b>TransformFinalBlock</b> has been called.
'
Public Property Get Hash() As Byte()
    Call VerifyNotDisposed

    If Not mHashFinished Then _
        Throw Cor.NewInvalidOperationException("Cannot get Hash until TransformFinalBlock is called.")

    Hash = mHash
End Property

''
' Computes the hash on a source in its entirety.
'
' @param Source The source of data compute the hash for. This can be a byte array or object implementing the Stream interface.
' @param Index The starting index in a byte array source. If the source is not a byte array, this argument is ignored.
' @param Count The number of bytes to compute the hash on. If the source is not a byte array, this argument is ignored.
' @return The hash value computed from the source provided.
' @remarks If the source is a Stream based object, then the stream is read from and the hash continuously computed until
' the stream no longer returns data.
' <p>If data has already been processed using TransformBlock, then that data will be included in the hashing process.</p>
'
Public Function ComputeHash(ByRef Source As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Call VerifyNotDisposed
    
    Select Case VarType(Source)
        Case vbByteArray:   Call ComputeHashOnBytes(Source, Index, Count)
        Case vbObject:      Call ComputeHashOnStream(Source)
        Case Else:          Throw Cor.NewArgumentException("Invalid source. Must be a byte array or Stream object.", "Source")
    End Select
    
    Dim Ret() As Byte
    Ret = mCore.HashFinal(mBuffer, mBytesBuffered, mMessageLength)
    Call InternalInitialize
    ComputeHash = Ret
End Function

''
' Continues to compute the hash value for blocks of data.
'
' @param InputBuffer The bytes to continue computing the hash from.
' @param InputOffset The index into the byte array to begin computing from.
' @param InputCount The number of bytes to be included in the hash computation.
' @param OutputBuffer The data after it has been processed. This will be the same as the input data, no changes are made.
' @param OutputOffset The starting index in the output buffer to place the processed data.
' @return The number of bytes that were processed.
' @remarks The OutputBuffer will contain the same plain text data as the input buffer. No transformation of the data
' is applied. The OutputBuffer array can be a Null array or the same array as the InputBuffer. If the InputBuffer is
' used as the OutputBuffer and the InputOffset equals the OutputOffset, no data is copied.
' <p>Once this method is called, the Hash property cannot be called until the TransformFinalBlock is called, finishing
' the hash computation.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    Call VerifyNotDisposed
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(InputBuffer), InputOffset, InputCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "InputBuffer", InputOffset, "InputOffset", InputCount, "InputCount")
    
    Call HashCore(InputBuffer, InputOffset, InputCount)
    
    If Not cArray.IsNull(OutputBuffer) Then
        ' We won't bother to copy if the OutputBuffer and offset are the same as
        ' the InputBuffer and offset since this would must be copying the same
        ' data over the top of itself.
        If (SAPtr(InputBuffer) <> SAPtr(OutputBuffer)) Or (InputOffset <> OutputOffset) Then
            Result = VerifyArrayRange(SAPtr(OutputBuffer), OutputOffset, InputCount)
            If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "OutputBuffer", OutputOffset, "OutputBuffer", InputCount, "OutputCount")
                
            Call CopyMemory(OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount)
        End If
    End If
    
    mHashFinished = False
    TransformBlock = InputCount
End Function

''
' Finalizes the hash computation by processing the last block of data.
'
' @param InputBuffer The bytes to finish the hash computation with.
' @param InputOffset The index into the byte array to begin hash computations.
' @param InputCount The number of bytes to be included in the final hash computation.
' @return A copy of the portion of the InputBuffer that was processed.
' @remarks The hash value is not returned by this method. To retrieve the final
' hash value, call the Hash property.
' <p>The Hash property can only be called after this method is called. If the Hash property
' is called before this method, then an exception is thrown.</p>
' <p>If the implementation can be reused (CanReuseTransform) then the implementation is reset
' to allow computing of a new hash value.</p>
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    Call VerifyNotDisposed
    
    If cArray.IsNull(InputBuffer) Then _
        Throw Cor.NewArgumentNullException("InputBuffer")
    If InputOffset < LBound(InputBuffer) Then _
        Throw Cor.NewArgumentOutOfRangeException("InputOffset", InputOffset, Environment.GetResourceString(ArgumentOutOfRange_LBound))
    If (cArray.GetLength(InputBuffer) - InputCount < InputOffset) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    Dim Ret() As Byte
    If InputCount > 0 Then
        Call HashCore(InputBuffer, InputOffset, InputCount)
        
        ReDim Ret(0 To InputCount - 1)
        Call CopyMemory(Ret(0), InputBuffer(InputOffset), InputCount)
    Else
        Ret = Cor.NewBytes
    End If
    
    mHash = mCore.HashFinal(mBuffer, mBytesBuffered, mMessageLength)
    mHashFinished = True
    Call InternalInitialize
    TransformFinalBlock = Ret
End Function

''
' Releases any resources being held by the hash algorithm.
'
' @remarks Once this is called, the object is disposed and cannot be reused.
'
Public Sub Clear()
    Erase mHash
    mDisposed = True
End Sub

''
' Resets the base class to an initial state to begin computing the
' hash for a new set of values.
'
' @remarks If this method is called by the base class itself, then the <b>IHashAlgorithm.Initialize</b>
' method will also be called. If this method is called from an external source, then the
' <b>IHashAlgorithm.Initialize</b> method will not be called. It is expected that the external
' source to be responsible for calling that method.
'
Public Sub Initialize()
    If mExecutingCall Then Exit Sub
    mExecutingCall = True
    
    If mInternalCall Then Call mCore.Initialize
    If mUseBuffer Then ReDim mBuffer(0 To mBlockSize - 1)
    mBytesBuffered = 0
    mMessageLength = 0@
    
    mExecutingCall = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Callback As IHashAlgorithm, ByVal KeepWeakReference As Boolean, ByVal UseBuffer As Boolean, ByVal BlockSize As Long)
    mUseBuffer = UseBuffer
    mBlockSize = BlockSize
    mKeepWeakReference = KeepWeakReference
    
    If KeepWeakReference Then
        ObjectPtr(mCore) = ObjectPtr(Callback)
    Else
        Set mCore = Callback
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InternalInitialize()
    mInternalCall = True
    Call Initialize
    mInternalCall = False
End Sub

Private Sub HashCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    If Count = 0 Then Exit Sub
    
    mMessageLength = mMessageLength + Count
    
    If Not mUseBuffer Then
        Call mCore.HashCore(Bytes, Index, Count)
        Exit Sub
    End If
    
    ' deal with any previously buffered partial data.
    If mBytesBuffered > 0 Then
        ' we have a partially filled buffer, so we need to deal with that first.
        Dim AmountToCopy As Long
        
        ' we only want to copy enough into the buffer to just fill it, if possible.
        AmountToCopy = MathExt.Min(mBlockSize - mBytesBuffered, Count)
        Call CopyMemory(mBuffer(mBytesBuffered), Bytes(Index), AmountToCopy)
        mBytesBuffered = mBytesBuffered + AmountToCopy
        Index = Index + AmountToCopy
        Count = Count - AmountToCopy
        
        ' if we filled the buffer then compute the hash and start over.
        If mBytesBuffered = mBlockSize Then
            Call mCore.HashCore(mBuffer, 0, mBlockSize)
            mBytesBuffered = 0
        End If
    End If
    
    ' process as many full blocks as we can.
    Do While Count >= mBlockSize
        Call mCore.HashCore(Bytes, Index, mBlockSize)
        Index = Index + mBlockSize
        Count = Count - mBlockSize
    Loop
    
    ' and store any remaining partial blocks.
    If Count > 0 Then
        Call CopyMemory(mBuffer(0), Bytes(Index), Count)
        mBytesBuffered = Count
    End If
End Sub

Private Sub ComputeHashOnStream(ByRef Source As Variant)
    If Source Is Nothing Then _
        Throw Cor.NewArgumentNullException("Source")
    If Not TypeOf Source Is Stream Then _
        Throw Cor.NewArgumentException("Source must implement the Stream interface.", "Source")
    
    Dim Stream As Stream
    Set Stream = Source
    
    ReDim InputBuffer(4095) As Byte
    Do
        Dim BytesRead As Long
        BytesRead = Stream.ReadBlock(InputBuffer, 0, 4096)
        If BytesRead = 0 Then Exit Do
        Call HashCore(InputBuffer, 0, BytesRead)
    Loop
End Sub

Private Sub ComputeHashOnBytes(ByRef Source As Variant, ByRef Index As Variant, ByRef Count As Variant)
    Dim pSA As Long
    pSA = GetArrayPointer(Source)
    
    Dim Result      As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Result = GetOptionalArrayRange(pSA, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Source", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    On Error GoTo errTrap
    Call ClearException
    
    Dim Bytes() As Byte
    SAPtr(Bytes) = pSA
    Call HashCore(Bytes, ElemIndex, ElemCount)

errTrap:
    SAPtr(Bytes) = vbNullPtr
    
    Dim Ex As Exception
    If Catch(Ex) Then Throw Ex
End Sub

Private Sub VerifyNotDisposed()
    If mDisposed Then
        Throw Cor.NewObjectDisposedException("Hash Object")
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    If mKeepWeakReference Then ObjectPtr(mCore) = vbNullPtr
End Sub
