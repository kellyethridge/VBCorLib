VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RijndaelManagedTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: RijndaelManagedTransform
'

''
' Performs cryptographic operations on data using the Rijndael algorithm.
'
' @remarks Details about the Rijndael algorithm can be found at http://csrc.nist.gov/CryptoToolkit/aes/rijndael
' <p>Use the RijndaelManaged class to create an instance of this class.</p>
' @see RijndaelManaged
'
Option Explicit
Implements IObject
Implements ICryptoTransform

' We use a structure to ensure the layout of the temp variables
' remain constant so we can manipulate them.
Private Type Temps
    ' we want these to remain in this layout so we can use CopyMemory.
    T0 As Long
    t1 As Long
    t2 As Long
    T3 As Long
    T4 As Long
    T5 As Long
    T6 As Long
    T7 As Long
    
    ' defining a fixed-size array in a struct is like declaring
    ' each element as its own variable in the struct, so accessing
    ' an element requires very little overhead compared to normal arrays.
    s(31) As Byte
End Type

Private mNr                 As Long ' Number of times a block will be mangled.
Private mNb                 As Long ' Number of 4-byte columns in a block.
Private mNk                 As Long ' Number of 4-byte columns in the key.
Private mExpKey()           As Long
Private mBlockSizeBits      As Long
Private mBlockSizeBytes     As Long
Private mInputBlockSize     As Long
Private mOutputBlockSize    As Long
Private mPadding            As PaddingMode
Private mMode               As CipherMode
Private mRgbIV()            As Byte
Private mWorkingIV()        As Byte
Private mTempIV()           As Byte
Private mIsEncrypting       As Boolean
Private mHasLastBlock       As Boolean
Private mLastBlock()        As Byte
Private mKeepLastBlock      As Boolean


' Lookup tables. We use fixed-length arrays to increase the access speed.
Private mT0(0 To 255)   As Long
Private mT1(0 To 255)   As Long
Private mT2(0 To 255)   As Long
Private mT3(0 To 255)   As Long
Private mT4(0 To 255)   As Long
Private mRCon() As Long
Private mSBox() As Byte



' predefined pointers used for quick array access.
Private mTmp            As Temps
Private mSPtr           As Long
Private mTmpPtr         As Long
Private mTempIVPtr      As Long
Private mWorkingIVPtr   As Long


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
' Returns the blocksize in bits.
'
' @return The blocksize in bits.
'
Public Property Get BlockSizeValue() As Long
    BlockSizeValue = mBlockSizeBits
End Property

''
' Returns the size of one block of data that is to be processed.
'
' @return The size of one block of data to be processed.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = mInputBlockSize
End Property

''
' Returns the size of the resulting block of data after an input block of data has been processed.
'
' @return The size of the resulting block of data after an input block of data has been processed.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = mOutputBlockSize
End Property

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
    ' an exact multiple of mInputBlockSize. There can be no partial blocks.
    If (InputCount <= 0) Or (InputCount Mod mInputBlockSize <> 0) Then _
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
    If CorArray.IsNull(InputBuffer) Then _
        Throw Cor.NewArgumentNullException("InputBuffer")
    If InputOffset < LBound(InputBuffer) Then _
        Throw Cor.NewArgumentOutOfRangeException("InputOffset", Environment.GetResourceString(ArgumentOutOfRange_LBound))
    If (CorArray.GetLength(InputBuffer) - InputCount < InputOffset) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    If mIsEncrypting Then
        TransformFinalBlock = EncryptFinalBlock(InputBuffer, InputOffset, InputCount)
    Else
        TransformFinalBlock = DecryptFinalBlock(InputBuffer, InputOffset, InputCount)
    End If
End Function

''
' Resets the transform so a new set of data can begin to be processed.
'
Public Sub Reset()
    If mMode <> CipherMode.ECB Then
        mWorkingIV = mRgbIV
        mTempIV = mRgbIV
        
        mWorkingIVPtr = VarPtr(mWorkingIV(0))
        mTempIVPtr = VarPtr(mTempIV(0))
    End If
    
    If Not mIsEncrypting Then
        mHasLastBlock = False
        ReDim mLastBlock(0 To mInputBlockSize)
    End If
End Sub

''
' Releases all resources used by the class.
'
Public Sub Clear()
    Erase mRgbIV
    Erase mWorkingIV
    Erase mTempIV
    Erase mLastBlock
    Erase mExpKey
    Erase mTmp.s
End Sub

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compare against this instance.
' @return Returns True if the values are the same.
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
' @return Returns a number identifing this instance.
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef RgbKey() As Byte, ByRef RgbIV() As Byte, ByVal Base As SymmetricAlgorithmBase, ByVal IsEncrypting As Boolean)
    mIsEncrypting = IsEncrypting
    
    With Base
        If (.Mode = CipherMode.CTS) Or (.Mode = CipherMode.OFB) Then _
            Throw Cor.NewCryptographicException("Invalid cipher mode.")
        
        mMode = .Mode
        mPadding = .Padding
        mBlockSizeBits = .BlockSize
        mBlockSizeBytes = mBlockSizeBits \ 8
        
        If mMode = CipherMode.CFB Then
            mInputBlockSize = .FeedbackSize \ 8
        Else
            mInputBlockSize = mBlockSizeBytes
        End If
            
        mOutputBlockSize = mInputBlockSize
    End With
    
    mNb = mBlockSizeBits \ 32
    mNk = CorArray.GetLength(RgbKey) \ 4
    
    Call InitLookupTables
    Call CalculateNumberOfRounds
    Call CreateExpandedKey(RgbKey)
    
    mRgbIV = RgbIV
    Call Reset
    
    If (mPadding <> PaddingMode.None) And (mPadding <> PaddingMode.Zeros) Then
        mKeepLastBlock = True
        ReDim mLastBlock(0 To mInputBlockSize - 1)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EncryptFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    Dim FullBlocks As Long
    FullBlocks = InputCount \ mInputBlockSize

    ' If we have a number of bytes to encrypt equal to a multiple of mBlockSize then
    ' we will have zero remaining bytes. Any padding will create a new full block instead
    ' of just filling in the last partial block.
    Dim RemainingBytes As Long
    RemainingBytes = InputCount - (FullBlocks * mOutputBlockSize)

    If (mPadding = PaddingMode.None) And (RemainingBytes <> 0) Then _
        Throw Cor.NewCryptographicException("Length of data to be encrypted is incorrect.")

    ' We need enough bytes to hold all the full blocks, 1 padded block produced by us.
    Dim TotalBytes As Long
    TotalBytes = (FullBlocks + 1) * mOutputBlockSize
    
    Dim OutputBuffer() As Byte
    ReDim OutputBuffer(0 To TotalBytes - 1)

    ' We have to copy the data to be encrypted to our output buffer because CryptoAPI
    ' puts the encrypted data back into the working array.
    If InputCount > 0 Then Call CopyMemory(OutputBuffer(0), InputBuffer(InputOffset), InputCount)

    TotalBytes = TotalBytes - CryptoHelper.PadBlock(OutputBuffer, FullBlocks * mOutputBlockSize + RemainingBytes, mPadding, mOutputBlockSize - RemainingBytes, mOutputBlockSize)
   
    If TotalBytes > 0 Then
        If TotalBytes < UBound(OutputBuffer) Then ReDim Preserve OutputBuffer(0 To TotalBytes - 1)
        
        Select Case mMode
            Case CipherMode.ECB: Call EncryptECB(OutputBuffer, 0, TotalBytes)
            Case CipherMode.CBC: Call EncryptCBC(OutputBuffer, 0, TotalBytes)
            Case CipherMode.CFB: Call EncryptCFB(OutputBuffer, 0, TotalBytes)
        End Select
    Else
        OutputBuffer = Cor.NewBytes()
    End If
    
    Call Reset
    EncryptFinalBlock = OutputBuffer
End Function

Private Function DecryptFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    If (InputCount Mod mInputBlockSize) <> 0 Then _
        Throw Cor.NewCryptographicException("Invalid data length.")

    Dim TotalBytes As Long
    TotalBytes = InputCount

    ' If we kept the last block from the previous decrypt call then we need
    ' to include it in this final decryption call, so up the total bytes to process.
    If mHasLastBlock Then
        TotalBytes = TotalBytes + mInputBlockSize
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
        Call CopyMemory(OutputBuffer(0), mLastBlock(0), mInputBlockSize)
        OutputOffset = mInputBlockSize
        mHasLastBlock = False
    End If

    ' If there is something in the InputBuffer to be processed, then
    ' add it to the working buffer, too.
    If InputCount > 0 Then Call CopyMemory(OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount)

    ' Have the decryption routine work on the bytes in the output buffer. The decrypted bytes
    ' will be placed back into the same buffer.
    Select Case mMode
        Case CipherMode.ECB: Call DecryptECB(OutputBuffer, 0, TotalBytes)
        Case CipherMode.CBC: Call DecryptCBC(OutputBuffer, 0, TotalBytes)
        Case CipherMode.CFB: Call DecryptCFB(OutputBuffer, 0, TotalBytes)
    End Select

    ' Remove the padding based on the type of padding expected.
    Dim DepadCount As Long
    DepadCount = CryptoHelper.DepadBlock(OutputBuffer, mPadding, mInputBlockSize)

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
        Call CopyMemory(OutputBuffer(NextOutputOffset), mLastBlock(0), mInputBlockSize)
        NextOutputOffset = NextOutputOffset + mInputBlockSize
        TotalBytes = TotalBytes + mInputBlockSize
        mHasLastBlock = False
    End If

    ' If we are still needing to keep the last block after each call, then we need
    ' to copy it now and not include it in the current operation by decrementing the
    ' number of bytes to be processed.
    If mKeepLastBlock Then
        TotalBytes = TotalBytes - mInputBlockSize
        InputCount = InputCount - mInputBlockSize
        Call CopyMemory(mLastBlock(0), InputBuffer(InputOffset + InputCount), mInputBlockSize)
        mHasLastBlock = True
    End If

    ' If there is anything left in the InputBuffer to process we need to copy it to the
    ' output buffer which the decryption routine will work on.
    If InputCount > 0 Then Call CopyMemory(OutputBuffer(NextOutputOffset), InputBuffer(InputOffset), InputCount)

    ' Have the decryption routine work on the bytes in the output buffer. The decrypted bytes
    ' will be placed back into the same buffer.
    Select Case mMode
        Case CipherMode.ECB: Call DecryptECB(OutputBuffer, OutputOffset, TotalBytes)
        Case CipherMode.CBC: Call DecryptCBC(OutputBuffer, OutputOffset, TotalBytes)
        Case CipherMode.CFB: Call DecryptCFB(OutputBuffer, OutputOffset, TotalBytes)
    End Select

    ' Return how many bytes we actually decrypted.
    DecryptBlock = TotalBytes
End Function

Private Function EncryptBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ' Since the cipher changes the data within the same array, we need to copy the
    ' plain text data to the output buffer and let the cipher work on that buffer.
    Call CopyMemory(OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount)

    Select Case mMode
        Case CipherMode.ECB: Call EncryptECB(OutputBuffer, OutputOffset, InputCount)
        Case CipherMode.CBC: Call EncryptCBC(OutputBuffer, OutputOffset, InputCount)
        Case CipherMode.CFB: Call EncryptCFB(OutputBuffer, OutputOffset, InputCount)
    End Select

    EncryptBlock = InputCount
End Function

''
' ECB feedback doesn't actually use any feedback or IV values. Each block is individually processed.
'
' This performs inplace encryption. The array should be the output array so as not to
' mangle the original input array.
'
Private Sub EncryptECB(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    ' encrypt a single blocks worth of data at a time.
    Dim i As Long
    For i = 0 To Count - 1 Step mInputBlockSize
        Select Case mBlockSizeBits
            Case 128: Call Encrypt128(Bytes, Index + i)
            Case 192: Call Encrypt192(Bytes, Index + i)
            Case 256: Call Encrypt256(Bytes, Index + i)
        End Select
    Next i
End Sub

''
' CBC feedback uses the IV array to kind of pre-mangle the data to be encrypted. Once a block
' has been encrypted, the IV is updated for the next block encryption.
'
Private Sub EncryptCBC(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Dim i As Long
    For i = 0 To Count - 1 Step mInputBlockSize
        ' apply the CBC feedback before encrypting the text.
        Dim j As Long
        For j = 0 To mInputBlockSize - 1
            Bytes(Index + i + j) = Bytes(Index + i + j) Xor mWorkingIV(j)
        Next j
    
        Select Case mBlockSizeBits
            Case 128: Call Encrypt128(Bytes, Index + i)
            Case 192: Call Encrypt192(Bytes, Index + i)
            Case 256: Call Encrypt256(Bytes, Index + i)
        End Select
        
        ' update the IV we are using for the next block to be mangled.
        Call CopyMemory(ByVal mWorkingIVPtr, Bytes(Index + i), mInputBlockSize)
    Next i
End Sub

''
' CFB will encrypt the current IV array instead of the plain text. Then the encrypted
' IV array will be Xor'd with the plain text data.
'
Private Sub EncryptCFB(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Dim i As Long
    For i = 0 To Count - 1 Step mInputBlockSize
        ' encrypt the IV array first.
        Call CopyMemory(ByVal mTempIVPtr, ByVal mWorkingIVPtr, mBlockSizeBytes)
        
        Select Case mBlockSizeBits
            Case 128: Call Encrypt128(mTempIV, 0)
            Case 192: Call Encrypt192(mTempIV, 0)
            Case 256: Call Encrypt256(mTempIV, 0)
        End Select
        
        ' Xor the IV array with the plain text.
        Dim j As Long
        For j = 0 To mInputBlockSize - 1
            Bytes(Index + i + j) = Bytes(Index + i + j) Xor mTempIV(j)
        Next j
        
        ' move the unused portion down to be used next time around. The unused
        ' portion will be encrypted again, as well.
        If (mBlockSizeBytes - mInputBlockSize) > 0 Then
            Call CopyMemory(ByVal mWorkingIVPtr, mWorkingIV(mInputBlockSize), mBlockSizeBytes - mInputBlockSize)
        End If
        
        ' Copy the cipher text back to the IV to be used in the next block.
        Call CopyMemory(mWorkingIV(mBlockSizeBytes - mInputBlockSize), Bytes(Index + i), mInputBlockSize)
    Next i
End Sub

''
' Performs a straight decryption on each individual inputblocksize of bytes. No feedback
' is used on the data.
'
Private Sub DecryptECB(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Dim i As Long
    For i = 0 To Count - 1 Step mInputBlockSize
        Select Case mBlockSizeBits
            Case 128: Call Decrypt128(Bytes, Index + i)
            Case 192: Call Decrypt192(Bytes, Index + i)
            Case 256: Call Decrypt256(Bytes, Index + i)
        End Select
    Next i
End Sub

''
' CBC feedback uses the IV array to kind of pre-mangle the data to be encrypted. Once a block
' has been encrypted, the IV is updated for the next block encryption.
'
Private Sub DecryptCBC(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Dim i As Long
    For i = 0 To Count - 1 Step mInputBlockSize
        Call CopyMemory(ByVal mTempIVPtr, Bytes(Index + i), mBlockSizeBytes)
        
        Select Case mBlockSizeBits
            Case 128: Call Decrypt128(Bytes, Index + i)
            Case 192: Call Decrypt192(Bytes, Index + i)
            Case 256: Call Decrypt256(Bytes, Index + i)
        End Select
        
        Dim j As Long
        For j = 0 To mInputBlockSize - 1
            Bytes(Index + i + j) = Bytes(Index + i + j) Xor mWorkingIV(j)
        Next j
        
        Call CopyMemory(ByVal mWorkingIVPtr, ByVal mTempIVPtr, mBlockSizeBytes)
    Next i
End Sub

''
' CFB will encrypt the current IV array instead of the plain text. Then the encrypted
' IV array will be Xor'd with the plain text data.
'
Private Sub DecryptCFB(ByRef Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Dim i As Long
    For i = 0 To Count - 1 Step mInputBlockSize
        Call CopyMemory(ByVal mTempIVPtr, ByVal mWorkingIVPtr, mBlockSizeBytes)
        
        Select Case mBlockSizeBits
            Case 128: Call Encrypt128(mTempIV, 0)
            Case 192: Call Encrypt192(mTempIV, 0)
            Case 256: Call Encrypt256(mTempIV, 0)
        End Select
        
        If (mBlockSizeBytes - mInputBlockSize) > 0 Then
            Call CopyMemory(ByVal mWorkingIVPtr, mWorkingIV(mInputBlockSize), mBlockSizeBytes - mInputBlockSize)
        End If
        
        Call CopyMemory(mWorkingIV(mBlockSizeBytes - mInputBlockSize), Bytes(Index + i), mInputBlockSize)
        
        Dim j As Long
        For j = 0 To mInputBlockSize - 1
            Bytes(Index + i + j) = Bytes(Index + i + j) Xor mTempIV(j)
        Next j
    Next i
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Encryption/Decryption blocks are unrolled to increase performance,
' therefore, there is an encryption and decryption routine for
' each block size.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Encrypt128(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim bPtr As Long
    bPtr = VarPtr(Bytes(Index))
    
    MemLong(mSPtr) = MemLong(bPtr) Xor mExpKey(0)
    MemLong(mSPtr + 4) = MemLong(bPtr + 4) Xor mExpKey(1)
    MemLong(mSPtr + 8) = MemLong(bPtr + 8) Xor mExpKey(2)
    MemLong(mSPtr + 12) = MemLong(bPtr + 12) Xor mExpKey(3)
    
    With mTmp
        Dim j As Long
        j = 4
        
        Dim i As Long
        For i = 1 To mNr - 1
            .T0 = mT0(.s(0)) Xor mT1(.s(5)) Xor mT2(.s(10)) Xor mT3(.s(15)) Xor mExpKey(j)
            .t1 = mT0(.s(4)) Xor mT1(.s(9)) Xor mT2(.s(14)) Xor mT3(.s(3)) Xor mExpKey(j + 1)
            .t2 = mT0(.s(8)) Xor mT1(.s(13)) Xor mT2(.s(2)) Xor mT3(.s(7)) Xor mExpKey(j + 2)
            .T3 = mT0(.s(12)) Xor mT1(.s(1)) Xor mT2(.s(6)) Xor mT3(.s(11)) Xor mExpKey(j + 3)
            
            j = j + 4

            Call CopyMemory(ByVal mSPtr, ByVal mTmpPtr, 16)
        Next i

        MemLong(bPtr) = (mT4(.s(0)) And &HFF&) Xor (mT4(.s(5)) And &HFF00&) Xor (mT4(.s(10)) And &HFF0000) Xor (mT4(.s(15)) And &HFF000000) Xor mExpKey(j)
        MemLong(bPtr + 4) = (mT4(.s(4)) And &HFF&) Xor (mT4(.s(9)) And &HFF00&) Xor (mT4(.s(14)) And &HFF0000) Xor (mT4(.s(3)) And &HFF000000) Xor mExpKey(j + 1)
        MemLong(bPtr + 8) = (mT4(.s(8)) And &HFF&) Xor (mT4(.s(13)) And &HFF00&) Xor (mT4(.s(2)) And &HFF0000) Xor (mT4(.s(7)) And &HFF000000) Xor mExpKey(j + 2)
        MemLong(bPtr + 12) = (mT4(.s(12)) And &HFF&) Xor (mT4(.s(1)) And &HFF00&) Xor (mT4(.s(6)) And &HFF0000) Xor (mT4(.s(11)) And &HFF000000) Xor mExpKey(j + 3)
    End With
End Sub

Private Sub Encrypt192(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim bPtr As Long
    bPtr = VarPtr(Bytes(Index))
    
    MemLong(mSPtr) = MemLong(bPtr) Xor mExpKey(0)
    MemLong(mSPtr + 4) = MemLong(bPtr + 4) Xor mExpKey(1)
    MemLong(mSPtr + 8) = MemLong(bPtr + 8) Xor mExpKey(2)
    MemLong(mSPtr + 12) = MemLong(bPtr + 12) Xor mExpKey(3)
    MemLong(mSPtr + 16) = MemLong(bPtr + 16) Xor mExpKey(4)
    MemLong(mSPtr + 20) = MemLong(bPtr + 20) Xor mExpKey(5)
    
    With mTmp
        Dim j As Long
        j = 6
        
        Dim i As Long
        For i = 1 To mNr - 1
            .T0 = mT0(.s(0)) Xor mT1(.s(5)) Xor mT2(.s(10)) Xor mT3(.s(15)) Xor mExpKey(j)
            .t1 = mT0(.s(4)) Xor mT1(.s(9)) Xor mT2(.s(14)) Xor mT3(.s(19)) Xor mExpKey(j + 1)
            .t2 = mT0(.s(8)) Xor mT1(.s(13)) Xor mT2(.s(18)) Xor mT3(.s(23)) Xor mExpKey(j + 2)
            .T3 = mT0(.s(12)) Xor mT1(.s(17)) Xor mT2(.s(22)) Xor mT3(.s(3)) Xor mExpKey(j + 3)
            .T4 = mT0(.s(16)) Xor mT1(.s(21)) Xor mT2(.s(2)) Xor mT3(.s(7)) Xor mExpKey(j + 4)
            .T5 = mT0(.s(20)) Xor mT1(.s(1)) Xor mT2(.s(6)) Xor mT3(.s(11)) Xor mExpKey(j + 5)
            j = j + 6

            Call CopyMemory(ByVal mSPtr, ByVal mTmpPtr, 24)
        Next i

        MemLong(bPtr) = (mT4(.s(0)) And &HFF&) Xor (mT4(.s(5)) And &HFF00&) Xor (mT4(.s(10)) And &HFF0000) Xor (mT4(.s(15)) And &HFF000000) Xor mExpKey(j)
        MemLong(bPtr + 4) = (mT4(.s(4)) And &HFF&) Xor (mT4(.s(9)) And &HFF00&) Xor (mT4(.s(14)) And &HFF0000) Xor (mT4(.s(19)) And &HFF000000) Xor mExpKey(j + 1)
        MemLong(bPtr + 8) = (mT4(.s(8)) And &HFF&) Xor (mT4(.s(13)) And &HFF00&) Xor (mT4(.s(18)) And &HFF0000) Xor (mT4(.s(23)) And &HFF000000) Xor mExpKey(j + 2)
        MemLong(bPtr + 12) = (mT4(.s(12)) And &HFF&) Xor (mT4(.s(17)) And &HFF00&) Xor (mT4(.s(22)) And &HFF0000) Xor (mT4(.s(3)) And &HFF000000) Xor mExpKey(j + 3)
        MemLong(bPtr + 16) = (mT4(.s(16)) And &HFF&) Xor (mT4(.s(21)) And &HFF00&) Xor (mT4(.s(2)) And &HFF0000) Xor (mT4(.s(7)) And &HFF000000) Xor mExpKey(j + 4)
        MemLong(bPtr + 20) = (mT4(.s(20)) And &HFF&) Xor (mT4(.s(1)) And &HFF00&) Xor (mT4(.s(6)) And &HFF0000) Xor (mT4(.s(11)) And &HFF000000) Xor mExpKey(j + 5)
    End With
End Sub

Private Sub Encrypt256(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim bPtr As Long
    bPtr = VarPtr(Bytes(Index))
    
    MemLong(mSPtr) = MemLong(bPtr) Xor mExpKey(0)
    MemLong(mSPtr + 4) = MemLong(bPtr + 4) Xor mExpKey(1)
    MemLong(mSPtr + 8) = MemLong(bPtr + 8) Xor mExpKey(2)
    MemLong(mSPtr + 12) = MemLong(bPtr + 12) Xor mExpKey(3)
    MemLong(mSPtr + 16) = MemLong(bPtr + 16) Xor mExpKey(4)
    MemLong(mSPtr + 20) = MemLong(bPtr + 20) Xor mExpKey(5)
    MemLong(mSPtr + 24) = MemLong(bPtr + 24) Xor mExpKey(6)
    MemLong(mSPtr + 28) = MemLong(bPtr + 28) Xor mExpKey(7)
    
    With mTmp
        Dim j As Long
        j = 8
        
        Dim i As Long
        For i = 1 To mNr - 1
            .T0 = mT0(.s(0)) Xor mT1(.s(5)) Xor mT2(.s(14)) Xor mT3(.s(19)) Xor mExpKey(j)
            .t1 = mT0(.s(4)) Xor mT1(.s(9)) Xor mT2(.s(18)) Xor mT3(.s(23)) Xor mExpKey(j + 1)
            .t2 = mT0(.s(8)) Xor mT1(.s(13)) Xor mT2(.s(22)) Xor mT3(.s(27)) Xor mExpKey(j + 2)
            .T3 = mT0(.s(12)) Xor mT1(.s(17)) Xor mT2(.s(26)) Xor mT3(.s(31)) Xor mExpKey(j + 3)
            .T4 = mT0(.s(16)) Xor mT1(.s(21)) Xor mT2(.s(30)) Xor mT3(.s(3)) Xor mExpKey(j + 4)
            .T5 = mT0(.s(20)) Xor mT1(.s(25)) Xor mT2(.s(2)) Xor mT3(.s(7)) Xor mExpKey(j + 5)
            .T6 = mT0(.s(24)) Xor mT1(.s(29)) Xor mT2(.s(6)) Xor mT3(.s(11)) Xor mExpKey(j + 6)
            .T7 = mT0(.s(28)) Xor mT1(.s(1)) Xor mT2(.s(10)) Xor mT3(.s(15)) Xor mExpKey(j + 7)
            
            j = j + 8
            
            Call CopyMemory(ByVal mSPtr, ByVal mTmpPtr, 32)
        Next i
    
        MemLong(bPtr) = (mT4(.s(0)) And &HFF&) Xor (mT4(.s(5)) And &HFF00&) Xor (mT4(.s(14)) And &HFF0000) Xor (mT4(.s(19)) And &HFF000000) Xor mExpKey(j)
        MemLong(bPtr + 4) = (mT4(.s(4)) And &HFF&) Xor (mT4(.s(9)) And &HFF00&) Xor (mT4(.s(18)) And &HFF0000) Xor (mT4(.s(23)) And &HFF000000) Xor mExpKey(j + 1)
        MemLong(bPtr + 8) = (mT4(.s(8)) And &HFF&) Xor (mT4(.s(13)) And &HFF00&) Xor (mT4(.s(22)) And &HFF0000) Xor (mT4(.s(27)) And &HFF000000) Xor mExpKey(j + 2)
        MemLong(bPtr + 12) = (mT4(.s(12)) And &HFF&) Xor (mT4(.s(17)) And &HFF00&) Xor (mT4(.s(26)) And &HFF0000) Xor (mT4(.s(31)) And &HFF000000) Xor mExpKey(j + 3)
        MemLong(bPtr + 16) = (mT4(.s(16)) And &HFF&) Xor (mT4(.s(21)) And &HFF00&) Xor (mT4(.s(30)) And &HFF0000) Xor (mT4(.s(3)) And &HFF000000) Xor mExpKey(j + 4)
        MemLong(bPtr + 20) = (mT4(.s(20)) And &HFF&) Xor (mT4(.s(25)) And &HFF00&) Xor (mT4(.s(2)) And &HFF0000) Xor (mT4(.s(7)) And &HFF000000) Xor mExpKey(j + 5)
        MemLong(bPtr + 24) = (mT4(.s(24)) And &HFF&) Xor (mT4(.s(29)) And &HFF00&) Xor (mT4(.s(6)) And &HFF0000) Xor (mT4(.s(11)) And &HFF000000) Xor mExpKey(j + 6)
        MemLong(bPtr + 28) = (mT4(.s(28)) And &HFF&) Xor (mT4(.s(1)) And &HFF00&) Xor (mT4(.s(10)) And &HFF0000) Xor (mT4(.s(15)) And &HFF000000) Xor mExpKey(j + 7)
    End With
End Sub

Private Sub Decrypt128(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim bPtr As Long
    bPtr = VarPtr(Bytes(Index))
    
    MemLong(mSPtr) = MemLong(bPtr) Xor mExpKey(0)
    MemLong(mSPtr + 4) = MemLong(bPtr + 4) Xor mExpKey(1)
    MemLong(mSPtr + 8) = MemLong(bPtr + 8) Xor mExpKey(2)
    MemLong(mSPtr + 12) = MemLong(bPtr + 12) Xor mExpKey(3)

    With mTmp
        Dim j As Long
        j = 4
        
        Dim i As Long
        For i = 1 To mNr - 1
            .T0 = mT0(.s(0)) Xor mT1(.s(13)) Xor mT2(.s(10)) Xor mT3(.s(7)) Xor mExpKey(j)
            .t1 = mT0(.s(4)) Xor mT1(.s(1)) Xor mT2(.s(14)) Xor mT3(.s(11)) Xor mExpKey(j + 1)
            .t2 = mT0(.s(8)) Xor mT1(.s(5)) Xor mT2(.s(2)) Xor mT3(.s(15)) Xor mExpKey(j + 2)
            .T3 = mT0(.s(12)) Xor mT1(.s(9)) Xor mT2(.s(6)) Xor mT3(.s(3)) Xor mExpKey(j + 3)
            j = j + 4

            Call CopyMemory(ByVal mSPtr, ByVal mTmpPtr, 16)
        Next i

        MemLong(bPtr) = (mT4(.s(0)) And &HFF&) Xor (mT4(.s(13)) And &HFF00&) Xor (mT4(.s(10)) And &HFF0000) Xor (mT4(.s(7)) And &HFF000000) Xor mExpKey(j)
        MemLong(bPtr + 4) = (mT4(.s(4)) And &HFF&) Xor (mT4(.s(1)) And &HFF00&) Xor (mT4(.s(14)) And &HFF0000) Xor (mT4(.s(11)) And &HFF000000) Xor mExpKey(j + 1)
        MemLong(bPtr + 8) = (mT4(.s(8)) And &HFF&) Xor (mT4(.s(5)) And &HFF00&) Xor (mT4(.s(2)) And &HFF0000) Xor (mT4(.s(15)) And &HFF000000) Xor mExpKey(j + 2)
        MemLong(bPtr + 12) = (mT4(.s(12)) And &HFF&) Xor (mT4(.s(9)) And &HFF00&) Xor (mT4(.s(6)) And &HFF0000) Xor (mT4(.s(3)) And &HFF000000) Xor mExpKey(j + 3)
    End With
End Sub

Private Sub Decrypt192(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim bPtr As Long
    bPtr = VarPtr(Bytes(Index))
    
    MemLong(mSPtr) = MemLong(bPtr) Xor mExpKey(0)
    MemLong(mSPtr + 4) = MemLong(bPtr + 4) Xor mExpKey(1)
    MemLong(mSPtr + 8) = MemLong(bPtr + 8) Xor mExpKey(2)
    MemLong(mSPtr + 12) = MemLong(bPtr + 12) Xor mExpKey(3)
    MemLong(mSPtr + 16) = MemLong(bPtr + 16) Xor mExpKey(4)
    MemLong(mSPtr + 20) = MemLong(bPtr + 20) Xor mExpKey(5)

    With mTmp
        Dim j As Long
        j = 6
        
        Dim i As Long
        For i = 1 To mNr - 1
            .T0 = mT0(.s(0)) Xor mT1(.s(21)) Xor mT2(.s(18)) Xor mT3(.s(15)) Xor mExpKey(j)
            .t1 = mT0(.s(4)) Xor mT1(.s(1)) Xor mT2(.s(22)) Xor mT3(.s(19)) Xor mExpKey(j + 1)
            .t2 = mT0(.s(8)) Xor mT1(.s(5)) Xor mT2(.s(2)) Xor mT3(.s(23)) Xor mExpKey(j + 2)
            .T3 = mT0(.s(12)) Xor mT1(.s(9)) Xor mT2(.s(6)) Xor mT3(.s(3)) Xor mExpKey(j + 3)
            .T4 = mT0(.s(16)) Xor mT1(.s(13)) Xor mT2(.s(10)) Xor mT3(.s(7)) Xor mExpKey(j + 4)
            .T5 = mT0(.s(20)) Xor mT1(.s(17)) Xor mT2(.s(14)) Xor mT3(.s(11)) Xor mExpKey(j + 5)
            j = j + 6

            Call CopyMemory(ByVal mSPtr, ByVal mTmpPtr, 24)
        Next i

        MemLong(bPtr) = (mT4(.s(0)) And &HFF&) Xor (mT4(.s(21)) And &HFF00&) Xor (mT4(.s(18)) And &HFF0000) Xor (mT4(.s(15)) And &HFF000000) Xor mExpKey(j)
        MemLong(bPtr + 4) = (mT4(.s(4)) And &HFF&) Xor (mT4(.s(1)) And &HFF00&) Xor (mT4(.s(22)) And &HFF0000) Xor (mT4(.s(19)) And &HFF000000) Xor mExpKey(j + 1)
        MemLong(bPtr + 8) = (mT4(.s(8)) And &HFF&) Xor (mT4(.s(5)) And &HFF00&) Xor (mT4(.s(2)) And &HFF0000) Xor (mT4(.s(23)) And &HFF000000) Xor mExpKey(j + 2)
        MemLong(bPtr + 12) = (mT4(.s(12)) And &HFF&) Xor (mT4(.s(9)) And &HFF00&) Xor (mT4(.s(6)) And &HFF0000) Xor (mT4(.s(3)) And &HFF000000) Xor mExpKey(j + 3)
        MemLong(bPtr + 16) = (mT4(.s(16)) And &HFF&) Xor (mT4(.s(13)) And &HFF00&) Xor (mT4(.s(10)) And &HFF0000) Xor (mT4(.s(7)) And &HFF000000) Xor mExpKey(j + 4)
        MemLong(bPtr + 20) = (mT4(.s(20)) And &HFF&) Xor (mT4(.s(17)) And &HFF00&) Xor (mT4(.s(14)) And &HFF0000) Xor (mT4(.s(11)) And &HFF000000) Xor mExpKey(j + 5)
    End With
End Sub

Private Sub Decrypt256(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim bPtr As Long
    bPtr = VarPtr(Bytes(Index))
    
    MemLong(mSPtr) = MemLong(bPtr) Xor mExpKey(0)
    MemLong(mSPtr + 4) = MemLong(bPtr + 4) Xor mExpKey(1)
    MemLong(mSPtr + 8) = MemLong(bPtr + 8) Xor mExpKey(2)
    MemLong(mSPtr + 12) = MemLong(bPtr + 12) Xor mExpKey(3)
    MemLong(mSPtr + 16) = MemLong(bPtr + 16) Xor mExpKey(4)
    MemLong(mSPtr + 20) = MemLong(bPtr + 20) Xor mExpKey(5)
    MemLong(mSPtr + 24) = MemLong(bPtr + 24) Xor mExpKey(6)
    MemLong(mSPtr + 28) = MemLong(bPtr + 28) Xor mExpKey(7)
    
    With mTmp
        Dim j As Long
        j = 8
        
        Dim i As Long
        For i = 1 To mNr - 1
            .T0 = mT0(.s(0)) Xor mT1(.s(29)) Xor mT2(.s(22)) Xor mT3(.s(19)) Xor mExpKey(j)
            .t1 = mT0(.s(4)) Xor mT1(.s(1)) Xor mT2(.s(26)) Xor mT3(.s(23)) Xor mExpKey(j + 1)
            .t2 = mT0(.s(8)) Xor mT1(.s(5)) Xor mT2(.s(30)) Xor mT3(.s(27)) Xor mExpKey(j + 2)
            .T3 = mT0(.s(12)) Xor mT1(.s(9)) Xor mT2(.s(2)) Xor mT3(.s(31)) Xor mExpKey(j + 3)
            .T4 = mT0(.s(16)) Xor mT1(.s(13)) Xor mT2(.s(6)) Xor mT3(.s(3)) Xor mExpKey(j + 4)
            .T5 = mT0(.s(20)) Xor mT1(.s(17)) Xor mT2(.s(10)) Xor mT3(.s(7)) Xor mExpKey(j + 5)
            .T6 = mT0(.s(24)) Xor mT1(.s(21)) Xor mT2(.s(14)) Xor mT3(.s(11)) Xor mExpKey(j + 6)
            .T7 = mT0(.s(28)) Xor mT1(.s(25)) Xor mT2(.s(18)) Xor mT3(.s(15)) Xor mExpKey(j + 7)
            j = j + 8
            
            Call CopyMemory(ByVal mSPtr, ByVal mTmpPtr, 32)
        Next i
    
        MemLong(bPtr) = (mT4(.s(0)) And &HFF&) Xor (mT4(.s(29)) And &HFF00&) Xor (mT4(.s(22)) And &HFF0000) Xor (mT4(.s(19)) And &HFF000000) Xor mExpKey(j)
        MemLong(bPtr + 4) = (mT4(.s(4)) And &HFF&) Xor (mT4(.s(1)) And &HFF00&) Xor (mT4(.s(26)) And &HFF0000) Xor (mT4(.s(23)) And &HFF000000) Xor mExpKey(j + 1)
        MemLong(bPtr + 8) = (mT4(.s(8)) And &HFF&) Xor (mT4(.s(5)) And &HFF00&) Xor (mT4(.s(30)) And &HFF0000) Xor (mT4(.s(27)) And &HFF000000) Xor mExpKey(j + 2)
        MemLong(bPtr + 12) = (mT4(.s(12)) And &HFF&) Xor (mT4(.s(9)) And &HFF00&) Xor (mT4(.s(2)) And &HFF0000) Xor (mT4(.s(31)) And &HFF000000) Xor mExpKey(j + 3)
        MemLong(bPtr + 16) = (mT4(.s(16)) And &HFF&) Xor (mT4(.s(13)) And &HFF00&) Xor (mT4(.s(6)) And &HFF0000) Xor (mT4(.s(3)) And &HFF000000) Xor mExpKey(j + 4)
        MemLong(bPtr + 20) = (mT4(.s(20)) And &HFF&) Xor (mT4(.s(17)) And &HFF00&) Xor (mT4(.s(10)) And &HFF0000) Xor (mT4(.s(7)) And &HFF000000) Xor mExpKey(j + 5)
        MemLong(bPtr + 24) = (mT4(.s(24)) And &HFF&) Xor (mT4(.s(21)) And &HFF00&) Xor (mT4(.s(14)) And &HFF0000) Xor (mT4(.s(11)) And &HFF000000) Xor mExpKey(j + 6)
        MemLong(bPtr + 28) = (mT4(.s(28)) And &HFF&) Xor (mT4(.s(25)) And &HFF00&) Xor (mT4(.s(18)) And &HFF0000) Xor (mT4(.s(15)) And &HFF000000) Xor mExpKey(j + 7)
    End With
End Sub

''
' Creates the expanded key used for encrypting data. If the CipherMode is CFB,
' then the expanded key is used for decryption as well.
'
Private Sub CreateExpandedKey(ByRef RgbKey() As Byte)
    ReDim mExpKey(0 To (mNb * (mNr + 1)) - 1)
    
    Call CopyMemory(mExpKey(0), RgbKey(0), CorArray.GetLength(RgbKey))
    
    Dim w As Long
    Dim i As Long
    
    For i = mNk To UBound(mExpKey)
        w = mExpKey(i - 1)
        If (i Mod mNk) = 0 Then
            w = SubWord(RRotate(w, 8))
            w = w Xor mRCon(i \ mNk - 1)
        ElseIf (mNk > 6) And ((i Mod mNk) = 4) Then
            w = SubWord(w)
        End If
        mExpKey(i) = mExpKey(i - mNk) Xor w
    Next i
    
    If (Not mIsEncrypting) And (mMode <> CipherMode.CFB) Then
        mExpKey = CreateDecryptionKey(mExpKey)
    End If
End Sub

Private Sub CalculateNumberOfRounds()
    If (mNk = 8) Or (mNb = 8) Then
        mNr = 14
    ElseIf (mNk = 6) Or (mNb = 6) Then
        mNr = 12
    Else
        mNr = 10
    End If
End Sub

Private Function SubWord(ByVal Word As Long) As Long
    With AsQByteArr(Word)
        .Bytes(0) = mSBox(.Bytes(0))
        .Bytes(1) = mSBox(.Bytes(1))
        .Bytes(2) = mSBox(.Bytes(2))
        .Bytes(3) = mSBox(.Bytes(3))
        
        SubWord = AsLong(.Bytes(0))
    End With
End Function

''
' Creates a decryption version of the expanded key. This will not be
' used if the CipherMode is CFB.
'
Private Function CreateDecryptionKey(ByRef ExpKey() As Long) As Long()
    Dim Ret() As Long
    ReDim Ret(0 To UBound(ExpKey))

    Dim i    As Long
    Dim j    As Long
    Dim k    As Long
    k = mNb * mNr
    For i = 0 To mNr
        Call CopyMemory(Ret(j), ExpKey(k), mBlockSizeBytes)
        j = j + mNb
        k = k - mNb
    Next i
    
    Dim Te4(0 To 255) As Long
    Dim Td0(0 To 255) As Long
    Dim Td1(0 To 255) As Long
    Dim Td2(0 To 255) As Long
    Dim Td3(0 To 255) As Long
    
    Call CopyMemory(Td0(0), ByVal Rijndael.Td0Ptr, 1024)
    Call CopyMemory(Td1(0), ByVal Rijndael.Td1Ptr, 1024)
    Call CopyMemory(Td2(0), ByVal Rijndael.Td2Ptr, 1024)
    Call CopyMemory(Td3(0), ByVal Rijndael.Td3Ptr, 1024)
    Call CopyMemory(Te4(0), ByVal Rijndael.Te4Ptr, 1024)
    
    For i = mNb To mNb * mNr - 1
        With AsQByteArr(Ret(i))
            Ret(i) = Td0(Te4(.Bytes(0)) And &HFF&) Xor _
                     Td1(Te4(.Bytes(1)) And &HFF&) Xor _
                     Td2(Te4(.Bytes(2)) And &HFF&) Xor _
                     Td3(Te4(.Bytes(3)) And &HFF&)
        End With
    Next i
    
    CreateDecryptionKey = Ret
End Function

Private Sub InitLookupTables()
    If (mIsEncrypting) Or (mMode = CipherMode.CFB) Then
        Call CopyMemory(mT0(0), ByVal Rijndael.Te0Ptr, 1024)
        Call CopyMemory(mT1(0), ByVal Rijndael.Te1Ptr, 1024)
        Call CopyMemory(mT2(0), ByVal Rijndael.Te2Ptr, 1024)
        Call CopyMemory(mT3(0), ByVal Rijndael.Te3Ptr, 1024)
        Call CopyMemory(mT4(0), ByVal Rijndael.Te4Ptr, 1024)
    Else
        Call CopyMemory(mT0(0), ByVal Rijndael.Td0Ptr, 1024)
        Call CopyMemory(mT1(0), ByVal Rijndael.Td1Ptr, 1024)
        Call CopyMemory(mT2(0), ByVal Rijndael.Td2Ptr, 1024)
        Call CopyMemory(mT3(0), ByVal Rijndael.Td3Ptr, 1024)
        Call CopyMemory(mT4(0), ByVal Rijndael.Td4Ptr, 1024)
    End If
    
    SAPtr(mSBox) = Rijndael.SBoxPtr
    SAPtr(mRCon) = Rijndael.RConPtr
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mTmpPtr = VarPtr(mTmp.T0)
    mSPtr = VarPtr(mTmp.s(0))
End Sub

Private Sub Class_Terminate()
    SAPtr(mSBox) = vbNullPtr
    SAPtr(mRCon) = vbNullPtr
End Sub


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


