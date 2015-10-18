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
' Module: HashAlgorithmBase
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
    VerifyNotDisposed
    
    If Not mHashFinished Then _
        Error.InvalidOperation Cryptography_HashNotYetFinalized

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
    VerifyNotDisposed
    
    Select Case VarType(Source)
        Case corByteArray
            ComputeHashOnBytes Source, Index, Count
        Case vbObject
            ComputeHashOnStream Source
        Case Else
            Error.Argument Argument_ByteArrayOrStreamRequired, "Source"
    End Select
    
    Dim Result() As Byte
    Result = mCore.HashFinal(mBuffer, mBytesBuffered, mMessageLength)
    InternalInitialize
    ComputeHash = Result
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
    VerifyNotDisposed
    ValidateByteArrayRange InputBuffer, InputOffset, InputCount, Parameter_InputBuffer, Parameter_InputOffset, Parameter_InputCount
    
    HashCore InputBuffer, InputOffset, InputCount
    
    If SAPtr(OutputBuffer) <> vbNullPtr Then
        ' We won't bother to copy if the OutputBuffer and offset are the same as
        ' the InputBuffer and offset since this would be copying the same
        ' data over the top of itself.
        If (SAPtr(InputBuffer) <> SAPtr(OutputBuffer)) Or (InputOffset <> OutputOffset) Then
            If OutputOffset < LBound(OutputBuffer) Then _
                Error.ArgumentOutOfRange "OutputOffset", ArgumentOutOfRange_ArrayLB
            If OutputOffset + InputCount - 1 > UBound(OutputBuffer) Then _
                Error.Argument Argument_InvalidOffLen
                
            CopyMemory OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount
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
    VerifyNotDisposed
    ValidateByteArrayRange InputBuffer, InputOffset, InputCount, Parameter_InputBuffer, Parameter_InputOffset, Parameter_InputCount
    
    Dim Result() As Byte
    If InputCount > 0 Then
        HashCore InputBuffer, InputOffset, InputCount
        
        ReDim Result(0 To InputCount - 1)
        CopyMemory Result(0), InputBuffer(InputOffset), InputCount
    Else
        Result = Cor.NewBytes
    End If
    
    mHash = mCore.HashFinal(mBuffer, mBytesBuffered, mMessageLength)
    mHashFinished = True
    InternalInitialize
    TransformFinalBlock = Result
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
    If mExecutingCall Then
        Exit Sub
    End If
    
    mExecutingCall = True
    
    If mInternalCall Then
        mCore.Initialize
    End If
    
    If mUseBuffer Then
        ReDim mBuffer(0 To mBlockSize - 1)
    End If
    
    mBytesBuffered = 0
    mMessageLength = 0
    mExecutingCall = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Callback As IHashAlgorithm, ByVal KeepWeakReference As Boolean, ByVal UseBuffer As Boolean, ByVal BlockSize As Long)
    mUseBuffer = UseBuffer
    mBlockSize = BlockSize
    
    If KeepWeakReference Then
        ObjectPtr(mCore) = ObjectPtr(Callback)
        mKeepWeakReference = True
    Else
        Set mCore = Callback
    End If
End Sub

Private Sub Class_Terminate()
    If mKeepWeakReference Then
        ObjectPtr(mCore) = vbNullPtr
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InternalInitialize()
    mInternalCall = True
    Initialize
    mInternalCall = False
End Sub

Private Sub HashCore(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    If Count = 0 Then
        Exit Sub
    End If
    
    mMessageLength = mMessageLength + Count
    
    If Not mUseBuffer Then
        mCore.HashCore Bytes, Index, Count
        Exit Sub
    End If
    
    ' deal with any previously buffered partial data.
    If mBytesBuffered > 0 Then
        ' we have a partially filled buffer, so we need to deal with that first.
        Dim AmountToCopy As Long
        
        ' we only want to copy enough into the buffer to just fill it, if possible.
        AmountToCopy = CorMath.Min(mBlockSize - mBytesBuffered, Count)
        CopyMemory mBuffer(mBytesBuffered), Bytes(Index), AmountToCopy
        mBytesBuffered = mBytesBuffered + AmountToCopy
        Index = Index + AmountToCopy
        Count = Count - AmountToCopy
        
        ' if we filled the buffer then compute the hash and start over.
        If mBytesBuffered = mBlockSize Then
            mCore.HashCore mBuffer, 0, mBlockSize
            mBytesBuffered = 0
        End If
    End If
    
    ' process as many full blocks as we can.
    Do While Count >= mBlockSize
        mCore.HashCore Bytes, Index, mBlockSize
        Index = Index + mBlockSize
        Count = Count - mBlockSize
    Loop
    
    ' and store any remaining partial blocks.
    If Count > 0 Then
        CopyMemory mBuffer(0), Bytes(Index), Count
        mBytesBuffered = Count
    End If
End Sub

Private Sub ComputeHashOnStream(ByRef Source As Variant)
    If Source Is Nothing Then _
        Error.ArgumentNull "Source", ArgumentNull_Generic
    If Not TypeOf Source Is Stream Then _
        Error.Argument Argument_StreamRequired, "Source"
    
    Dim Stream As Stream
    Set Stream = Source
    
    Dim InputBuffer(0 To 4095)  As Byte
    Dim BytesRead               As Long
    Do
        BytesRead = Stream.ReadBlock(InputBuffer, 0, 4096)
        If BytesRead = 0 Then
            Exit Do
        End If
        HashCore InputBuffer, 0, BytesRead
    Loop
End Sub

Private Sub ComputeHashOnBytes(ByRef Source As Variant, ByRef Index As Variant, ByRef Count As Variant)
    On Error GoTo Catch
    
    Dim Buffer() As Byte
    Dim Range As ListRange
    SAPtr(Buffer) = CorArray.ArrayPointer(Source)
    Range = ValidateByteArrayOptionalRange(Buffer, Index, Count, Parameter_Buffer)
    HashCore Buffer, Range.Index, Range.Count
    
    GoSub Finally
    Exit Sub
    
Catch:
    GoSub Finally
    Throw
Finally:
    SAPtr(Buffer) = vbNullPtr
    Return
End Sub

Private Sub VerifyNotDisposed()
    If mDisposed Then
        Throw Cor.NewObjectDisposedException("Hash Object")
    End If
End Sub

