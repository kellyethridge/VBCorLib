VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACTripleDES"
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
'    Module: MACTripleDES
'

''
' Represents a TripleDES implementation of an HMAC algorithm.
'
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements KeyedHashAlgorithm
Implements ICryptoTransform

Private mTripleDES      As TripleDES
Private mStream         As CryptoStream
Private mHashStream     As New SingleBlockStream
Private mDisposed       As Boolean
Private mHashStarted    As Boolean



''
' Returns the padding being used by the algorithm.
'
' @return The padding being used.
'
Public Property Get Padding() As PaddingMode
    Call VerifyNotDisposed
    Padding = mTripleDES.Padding
End Property

''
' Sets the padding to be used by the hash algorithm.
'
' @param RHS The new padding value.
'
Public Property Let Padding(ByVal RHS As PaddingMode)
    Call VerifyNotDisposed
    mTripleDES.Padding = RHS
End Property

''
' Returns the key being used by the hash algorithm.
'
' @return The key being used.
'
Public Property Get Key() As Byte()
    Call VerifyNotDisposed
    Key = mTripleDES.Key
End Property

''
' Sets the key to be used by the hash algorithm.
'
' @param RHS The new key to be used.
'
Public Property Let Key(ByRef RHS() As Byte)
    Call VerifyNotDisposed
    
    If mHashStarted Then _
        Throw Cor.NewCryptographicException("Cannot change key after hash has started.")
    
    Call SetKey(RHS)
    Set mStream = Nothing
End Property

''
' Returns if the hash algorithm implementation can be reused once a hash has been computed.
'
' @returns True if the implementation can be reused, False otherwise.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = True
End Property

''
' Returns if the hash algorithm implementation can processd multiple blocks of data at a time.
'
' @return True if multiple blocks of data can be processed at a time, False otherwise.
' @remarks The size of 1 block of data is defined by the InputBlockSize property.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = True
End Property

''
' Returns the computed hash after all data has been processed.
'
' @return A byte array containing the computed hash.
' @remarks The size of the array will vary depending on the hash algorithm being performed.
'
Public Property Get Hash() As Byte()
    Call VerifyNotDisposed
    
    If mHashStarted Then _
        Throw Cor.NewInvalidOperationException("Cannot get Hash until TransformFinalBlock is called.")

    Hash = mHashStream.Hash
End Property

''
' Returns the size of the hash value in bits.
'
' @return The size of the hash value in bits.
'
Public Property Get HashSize() As Long
    HashSize = 64
End Property

''
' Returns the size of the input block size allowed.
'
' @returns The number of bytes a single input block is.
' @remarks When calling the TransformBlock method, the InputCount must equal this value.
' If the hash's CanTransformMultipleBlocks returns True, then the InputCount can be
' larger than the InputBlockSize, but must be an even multiple of InputBlockSize.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = 1
End Property

''
' Returns the size of the output block.
'
' @return The size of the output block.
' @remarks When calling the TransformBlock method, the return value will be 0 or a multiple of this value.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = 1
End Property

''
' Releases any resources being held by the hash algorithm.
'
' @remarks Once this is called, the object is disposed and cannot be reused.
'
Public Sub Clear()
    Set mStream = Nothing
    Set mHashStream = Nothing
    Set mTripleDES = Nothing
    mDisposed = True
End Sub

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
    Call StartHash
    
    Select Case VarType(Source)
        Case vbByteArray:   Call ComputeHashOnBytes(Source, Index, Count)
        Case vbObject:      Call ComputeHashOnStream(Source)
        Case Else:          Throw Cor.NewArgumentException("Invalid source. Must be a byte array or Stream object.", "Source")
    End Select
    
    Call EndHash
    ComputeHash = mHashStream.Hash
End Function

''
' Returns if the Value is equal to this object instance.
'
' @param Value The value to compare against this object instance.
' @return True if they are considered equal, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Resets a hashing object to begin hashing a new value.
'
' @remarks If Clear has been called, then the object cannot be re-initialized.
'
Public Sub Initialize()
    Set mStream = Nothing
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
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
    Call StartHash
    
    If Not cArray.IsNull(OutputBuffer) Then
        ' We won't bother to copy if the OutputBuffer and offset are the same as
        ' the InputBuffer and offset since this would be copying the same
        ' data over the top of itself.
        If (SAPtr(InputBuffer) <> SAPtr(OutputBuffer)) Or (InputOffset <> OutputOffset) Then
            Dim Result As Long
            Result = VerifyArrayRange(SAPtr(OutputBuffer), OutputOffset, InputCount)
            If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "OutputBuffer", OutputOffset, "OutputBuffer", InputCount, "OutputCount")
                
            Call CopyMemory(OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount)
        End If
    End If
    
    Call mStream.WriteBlock(InputBuffer, InputOffset, InputCount)
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
' <p>If the implementation can be reused (CanReuseTransform) then the implemntation is reset
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

    Call StartHash
    
    Dim Ret() As Byte
    If InputCount > 0 Then
        Call mStream.WriteBlock(InputBuffer, InputOffset, InputCount)
        
        ReDim Ret(0 To InputCount - 1)
        Call CopyMemory(Ret(0), InputBuffer(InputOffset), InputCount)
    Else
        Ret = Cor.NewBytes
    End If
    
    Call EndHash
    TransformFinalBlock = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef RgbKey() As Byte, ByVal StrTripleDES As String)
    Set mTripleDES = CryptoConfig.CreateFromName(StrTripleDES)
    Call SetKey(RgbKey)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetKey(ByRef Key() As Byte)
    ' we have to do it this way to work around the Byte array assignment bug in VB6.
    Dim b() As Byte
    b = Key
    mTripleDES.Key = b
End Sub

Private Sub StartHash()
    If mHashStarted Then Exit Sub
    Set mStream = Cor.NewCryptoStream(mHashStream, mTripleDES.CreateEncryptor, WriteMode)
    mHashStarted = True
End Sub

Private Sub EndHash()
    Call mStream.FlushFinalBlock
    Call Initialize
    mHashStarted = False
End Sub

Private Sub ComputeHashOnStream(ByRef Source As Variant)
    If Source Is Nothing Then _
        Throw Cor.NewArgumentNullException("Source")
    If Not TypeOf Source Is Stream Then _
        Throw Cor.NewArgumentException("Source must implement the Stream interface.", "Source")
    
    Dim Stream As Stream
    Set Stream = Source
    
    ReDim InputBuffer(4095) As Byte
    Dim BytesRead As Long
    Do
        BytesRead = Stream.ReadBlock(InputBuffer, 0, 4096)
        If BytesRead = 0 Then Exit Do
        Call mStream.WriteBlock(InputBuffer, 0, BytesRead)
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
    Dim Bytes() As Byte
    SAPtr(Bytes) = pSA
    Call mStream.WriteBlock(Bytes, ElemIndex, ElemCount)

errTrap:
    SAPtr(Bytes) = vbNullPtr
    Throw
End Sub

Private Sub VerifyNotDisposed()
    If mDisposed Then
        Throw Cor.NewObjectDisposedException("SHA1CryptoServiceProvider", "The service provider has been disposed.")
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mTripleDES = New TripleDESCryptoServiceProvider
    mTripleDES.Mode = CipherMode.CBC
    mTripleDES.IV = Cor.NewBytes(0, 0, 0, 0, 0, 0, 0, 0)
    Call mHashStream.Init(8)
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
'   HashAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get HashAlgorithm_CanReuseTransform() As Boolean
    HashAlgorithm_CanReuseTransform = CanReuseTransform
End Property

Private Property Get HashAlgorithm_CanTransformMultipleBlocks() As Boolean
    HashAlgorithm_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub HashAlgorithm_Clear()
    Call Clear
End Sub

Private Function HashAlgorithm_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    HashAlgorithm_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function HashAlgorithm_Equals(Value As Variant) As Boolean
    HashAlgorithm_Equals = Equals(Value)
End Function

Private Function HashAlgorithm_GetHashCode() As Long
    HashAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Get HashAlgorithm_Hash() As Byte()
    HashAlgorithm_Hash = Hash
End Property

Private Property Get HashAlgorithm_HashSize() As Long
    HashAlgorithm_HashSize = HashSize
End Property

Private Sub HashAlgorithm_Initialize()
    Call Initialize
End Sub

Private Property Get HashAlgorithm_InputBlockSize() As Long
    HashAlgorithm_InputBlockSize = InputBlockSize
End Property

Private Property Get HashAlgorithm_OutputBlockSize() As Long
    HashAlgorithm_OutputBlockSize = OutputBlockSize
End Property

Private Function HashAlgorithm_ToString() As String
    HashAlgorithm_ToString = ToString
End Function

Private Function HashAlgorithm_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    HashAlgorithm_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function HashAlgorithm_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    HashAlgorithm_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   KeyedHashAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let KeyedHashAlgorithm_Key(RHS() As Byte)
    Call SetKey(RHS)
End Property

Private Property Get KeyedHashAlgorithm_Key() As Byte()
    KeyedHashAlgorithm = Key
End Property

Private Property Get KeyedHashAlgorithm_CanReuseTransform() As Boolean
    KeyedHashAlgorithm_CanReuseTransform = CanReuseTransform
End Property

Private Property Get KeyedHashAlgorithm_CanTransformMultipleBlocks() As Boolean
    KeyedHashAlgorithm_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub KeyedHashAlgorithm_Clear()
    Call Clear
End Sub

Private Function KeyedHashAlgorithm_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    KeyedHashAlgorithm_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function KeyedHashAlgorithm_Equals(Value As Variant) As Boolean
    KeyedHashAlgorithm_Equals = Equals(Value)
End Function

Private Function KeyedHashAlgorithm_GetHashCode() As Long
    KeyedHashAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Get KeyedHashAlgorithm_Hash() As Byte()
    KeyedHashAlgorithm_Hash = Hash
End Property

Private Property Get KeyedHashAlgorithm_HashSize() As Long
    KeyedHashAlgorithm_HashSize = HashSize
End Property

Private Sub KeyedHashAlgorithm_Initialize()
    Call Initialize
End Sub

Private Property Get KeyedHashAlgorithm_InputBlockSize() As Long
    KeyedHashAlgorithm_InputBlockSize = InputBlockSize
End Property

Private Property Get KeyedHashAlgorithm_OutputBlockSize() As Long
    KeyedHashAlgorithm_OutputBlockSize = OutputBlockSize
End Property

Private Function KeyedHashAlgorithm_ToString() As String
    KeyedHashAlgorithm_ToString = ToString
End Function

Private Function KeyedHashAlgorithm_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    KeyedHashAlgorithm_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function KeyedHashAlgorithm_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    KeyedHashAlgorithm_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function




