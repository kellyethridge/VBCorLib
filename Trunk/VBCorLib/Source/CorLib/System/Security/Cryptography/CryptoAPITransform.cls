VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoAPITransform"
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
' Module: CryptoAPITransform
'

''
' Utilizes the Windows CryptoAPI cryptography services to perform data transformations.
'
' @remarks This class is used to support 3 cryptographic service providers within the
' CryptoAPI: DES, TripleDES, and RC2.
'
' @see DESCryptoServiceProvider
' @see TripleDESCryptoServiceProvider
' @see RC2CryptoServiceProvider
'
Option Explicit
Implements IObject
Implements ICryptoTransform

Private mBlockSize          As Long
Private mCipherKey          As Long
Private mPadding            As PaddingMode
Private mIsEncrypting       As Boolean
Private mLastBlock()        As Byte
Private mKeepLastBlock      As Boolean
Private mHasLastBlock       As Boolean



''
' Returns if the current transfrom can be reused after calling TransformFinalBlock.
'
' @return Returns True if it can be reused, False otherwise.
' @remarks This always returns True.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = True
End Property

''
' Returns if this transform can process more than one block per TransformBlock call.
'
' @return This always returns True.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = True
End Property

''
' Returns the size of one block of data that is to be processed.
'
' @return The size of one block of data to be processed.
' @remarks When calling the TransformBlock method, the InputCount must be zero or a multiple of this value.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = mBlockSize
End Property

''
' Returns the size of the resulting block of data after an input block of data has been processed.
'
' @return The size of the resulting block of data after an input block of data has been processed.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = mBlockSize
End Property

''
' This returns the actual handle to a key created by the CryptoAPI service.
'
' @return A handle to a CryptoAPI key.
'
Public Property Get KeyHandle() As Long
    KeyHandle = mCipherKey
End Property

''
' This will release the key and context of the CryptoAPI provider.
'
Public Sub Clear()
    Call CryptoAPI.DestroyKey(mCipherKey)
    Erase mLastBlock
End Sub

''
' Resets the transform so a new set of data can begin to be processed.
'
Public Sub Reset()
    mHasLastBlock = False
    Call CryptEncrypt(mCipherKey, 0, True, 0, ByVal 0&, 0, 0)

    ReDim mLastBlock(0 To mBlockSize - 1)
    Call CryptDecrypt(mCipherKey, 0, True, 0, mLastBlock(0), InputBlockSize)
    ReDim mLastBlock(0 To mBlockSize - 1)
End Sub

''
' Processes a set of data, encrypting or decrypting it.
'
' @param InputBuffer The data to be processed.
' @param InputOffset The index to begin processing from.
' @param InputCount The amount of data to be processed.
' @param OutputBuffer The buffer to place the processed data.
' @param OutputOffset The index to begin placing the processed data.
' @return The number of bytes that were processed.
' @remarks The InputCount must be an exact multiple of the InputBlockSize or an exception will be thrown.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ' During normal transformations, the number of bytes to be worked on must be
    ' an exact multiple of mBlockSize. There can be no partial blocks.
    If (InputCount <= 0) Or (InputCount Mod mBlockSize <> 0) Then _
        Throw Cor.NewArgumentException("Invalid value.", "InputCount")

    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(InputBuffer), InputOffset, InputCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "InputBuffer", InputOffset, "InputOffset", InputCount, "InputCount", False)

    ' We can use the InputCount for the number of bytes to be outputted because
    ' this cipher keeps the original number of bytes in the encrypted output, except
    ' for any padding which is not added in this process.
    Result = VerifyArrayRange(SAPtr(OutputBuffer), OutputOffset, InputCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "OutputBuffer", OutputOffset, "OutputOffset", InputCount, "", False)

    If mIsEncrypting Then
        TransformBlock = EncryptBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
    Else
        TransformBlock = DecryptBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
    End If
End Function

''
' Processes a set of data adding any padding needed to complete the process.
'
' @param InputBuffer The final data to be processed.
' @param InputOffset The index to begin processing from.
' @param InputCount The amount of data to be processed.
' @return The final block of processed data.
' @remarks Once a final block has been processed, the transform is reset to begin transforming new data.
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    If cArray.IsNull(InputBuffer) Then _
        Throw Cor.NewArgumentNullException("InputBuffer")
    If InputOffset < LBound(InputBuffer) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "InputOffset", InputOffset)
    If (cArray.GetLength(InputBuffer) - InputCount < InputOffset) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))

    If mIsEncrypting Then
        TransformFinalBlock = EncryptFinalBlock(InputBuffer, InputOffset, InputCount)
    Else
        TransformFinalBlock = DecryptFinalBlock(InputBuffer, InputOffset, InputCount)
    End If
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Algorithm As Long, ByRef RgbKey() As Byte, ByRef RgbIV() As Byte, ByVal Base As SymmetricAlgorithmBase, ByVal IsEncrypting As Boolean, Optional ByVal EffectiveKeySize As Long = 0, Optional ByVal UseSalt As Boolean = False)
    With Base
        If .Mode = CipherMode.OFB Then _
            Throw Cor.NewCryptographicException("Output Feedback (OFB) mode is not supported.")
        
        mCipherKey = CryptoAPI.ImportPlainTextKey(CryptoAPI.DefaultProvider, Algorithm, RgbKey, UseSalt)

        Call CryptoAPI.SetKeyParam(mCipherKey, KP_MODE, .Mode)

        If .Mode <> CipherMode.ECB Then Call CryptoAPI.SetKeyParam(mCipherKey, KP_IV, RgbIV)
        If .Mode = CipherMode.CFB Then Call CryptoAPI.SetKeyParam(mCipherKey, KP_MODE_BITS, .FeedbackSize)
        If EffectiveKeySize > 0 Then Call CryptoAPI.SetKeyParam(mCipherKey, KP_EFFECTIVE_KEYLEN, EffectiveKeySize)

        mBlockSize = .BlockSize \ 8
        mPadding = .Padding

        If (mPadding <> PaddingMode.None) And (mPadding <> PaddingMode.Zeros) Then
            mKeepLastBlock = True
            ReDim mLastBlock(0 To mBlockSize - 1)
        End If
    End With

    mIsEncrypting = IsEncrypting
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' The way we handle custom padding is to pad the block ourselves then let the CryptoAPI
' add one additional padded block. Afterwards, we strip that additional block off.
'
Private Function EncryptFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    Dim FullBlocks As Long
    FullBlocks = InputCount \ mBlockSize

    ' If we have a number of bytes to encrypt equal to a multiple of mBlockSize then
    ' we will have zero remaining bytes. Any padding will create a new full block instead
    ' of just filling in the last partial block.
    Dim RemainingBytes As Long
    RemainingBytes = InputCount - (FullBlocks * mBlockSize)

    If (mPadding = PaddingMode.None) And (RemainingBytes <> 0) Then _
        Throw Cor.NewCryptographicException("Length of data to be encrypted is incorrect.")

    ' We need enough bytes to hold all the full blocks, 1 padded block produced by us
    ' and 1 padded block that the CryptoAPI must create.
    Dim TotalBytes As Long
    TotalBytes = (FullBlocks + 2) * mBlockSize

    Dim OutputBuffer() As Byte
    ReDim OutputBuffer(0 To TotalBytes - 1)

    ' We have to copy the data to be encrypted to our output buffer because CryptoAPI
    ' puts the encrypted data back into the working array.
    If InputCount > 0 Then Call CopyMemory(OutputBuffer(0), InputBuffer(InputOffset), InputCount)

    ' We subtract the number of padding bytes that were not added.
    TotalBytes = TotalBytes - CryptoHelper.PadBlock(OutputBuffer, FullBlocks * mBlockSize + RemainingBytes, mPadding, mBlockSize - RemainingBytes, mBlockSize)

    Dim Size As Long
    Size = TotalBytes - mBlockSize ' We don't want to include the CryptoAPI padding block as part of the data to encrypt.
    If CryptEncrypt(mCipherKey, 0, True, 0, OutputBuffer(0), Size, TotalBytes) = 0 Then _
        Throw Cor.NewCryptographicException("Encryption failed: " & GetErrorMessage(Err.LastDllError))

    ' Size will be set to the total bytes encrypted, including the additional block of padding.
    ' We don't want that additional block of padding from the CryptoAPI since we've added our own,
    ' so we will just strip of the last block of bytes.
    Dim NewSize As Long
    NewSize = Size - mBlockSize
    If NewSize > 0 Then
        ReDim Preserve OutputBuffer(0 To NewSize - 1)
    Else
        OutputBuffer = Cor.NewBytes()
    End If

    Call Reset
    EncryptFinalBlock = OutputBuffer
End Function

Private Function EncryptBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ' Since the cipher changes the data within the same array, we need to copy the
    ' plain text data to the output buffer and let the cipher work on that buffer.
    Call CopyMemory(OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount)

    If CryptEncrypt(mCipherKey, 0, False, 0, OutputBuffer(OutputOffset), InputCount, InputCount) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Failed to encrypt data. " & GetErrorMessage(Err.LastDllError))

    EncryptBlock = InputCount
End Function

''
' This function is usually called repeatedly. The last block may be buffered and not processed
' on the same call. It may be processed during the next call or the call to TransformFinalBlock.
'
Private Function DecryptBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    Dim NextOutputOffset As Long
    NextOutputOffset = OutputOffset

    ' Start out attempting to process the original requested number of bytes.
    Dim TotalBytes As Long
    TotalBytes = InputCount

    ' If we buffered the last block from the previous call, then we now need to
    ' include it in this decryption process as the first block of the InputBuffer.
    If mHasLastBlock Then
        Call CopyMemory(OutputBuffer(NextOutputOffset), mLastBlock(0), mBlockSize)
        NextOutputOffset = NextOutputOffset + mBlockSize
        TotalBytes = TotalBytes + mBlockSize
        mHasLastBlock = False
    End If

    ' If we are still needing to keep the last block after each call, then we need
    ' to copy it now and not include it in the current operation by decrementing the
    ' number of bytes to be processed.
    If mKeepLastBlock Then
        TotalBytes = TotalBytes - mBlockSize
        InputCount = InputCount - mBlockSize
        Call CopyMemory(mLastBlock(0), InputBuffer(InputOffset + InputCount), mBlockSize)
        mHasLastBlock = True
    End If

    ' If there is anything left in the InputBuffer to process we need to copy it to the
    ' output buffer which the decryption routine will work on.
    If InputCount > 0 Then Call CopyMemory(OutputBuffer(NextOutputOffset), InputBuffer(InputOffset), InputCount)

    ' Have the decryption routine work on the bytes in the output buffer. The decrypted bytes
    ' will be placed back into the same buffer.
    If CryptDecrypt(mCipherKey, 0, False, 0, OutputBuffer(OutputOffset), TotalBytes) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not decrypt data. " & GetErrorMessage(Err.LastDllError))

    ' Return how many bytes we actually decrypted.
    DecryptBlock = TotalBytes
End Function

Private Function DecryptFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    If (InputCount Mod mBlockSize) <> 0 Then _
        Throw Cor.NewCryptographicException("Invalid data length.")

    Dim TotalBytes As Long
    TotalBytes = InputCount

    ' If we kept the last block from the previous decrypt call then we need
    ' to include it in this final decryption call, so up the total bytes to process.
    If mHasLastBlock Then
        TotalBytes = TotalBytes + mBlockSize
    End If

    ' If we still don't have anything to process (no previous block and nothing in this block)
    ' then simply reset everything and return an empty byte array representing the final block.
    If TotalBytes = 0 Then
        Call Reset
        DecryptFinalBlock = Cor.NewBytes
        Exit Function
    End If

    ' Allocate the buffer that all the work will be performed on.
    Dim OutputBuffer() As Byte
    ReDim OutputBuffer(0 To TotalBytes - 1)

    ' If we kept the last block from the previous decrypt call then we need
    ' to now copy that block into the working buffer.
    Dim OutputOffset As Long
    If mHasLastBlock Then
        Call CopyMemory(OutputBuffer(0), mLastBlock(0), mBlockSize)
        OutputOffset = mBlockSize
        mHasLastBlock = False
    End If

    ' If there is something in the InputBuffer to be processed, then
    ' add it to the working buffer, too.
    If InputCount > 0 Then Call CopyMemory(OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount)

    ' Decrypt the working buffer. We don't tell the CryptDecrypt method that this is the
    ' final block because we don't want the CryptoAPI to deal with any padding. We will
    ' deal with that next ourselves so we can support different padding types.
    If CryptDecrypt(mCipherKey, 0, False, 0, OutputBuffer(0), TotalBytes) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not decrypt data. " & GetErrorMessage(Err.LastDllError))

    ' Remove the padding based on the type of padding expected.
    Dim DepadCount As Long
    DepadCount = CryptoHelper.DepadBlock(OutputBuffer, mPadding, mBlockSize)

    ' Calculate how many plain text bytes we really have.
    Dim NewSize As Long
    NewSize = TotalBytes - DepadCount
    If NewSize > 0 Then
        ' Chop of the end padding bytes now. This is pretty efficient
        ' because the data is not actually moved since we are shrinking the array.
        ReDim Preserve OutputBuffer(0 To NewSize - 1)
    Else
        ' Nothing is left, so return an empty array.
        OutputBuffer = Cor.NewBytes
    End If

    Call Reset
    DecryptFinalBlock = OutputBuffer
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call Clear
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICryptoTransform Interface
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
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
