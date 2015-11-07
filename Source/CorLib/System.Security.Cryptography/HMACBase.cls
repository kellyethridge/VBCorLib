VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACBase"
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
' Module: HMACBase
'

''
' Provides the base functionality to perfrom Hash Message Authentication Code operations.
'
' @remarks This base implementation manages many of the aspects utilized with HMAC implementations.
' Usually, any class that implements the <b>HashAlgorithm</b> interface can easily be used to
' provide HMAC capabilities by simply initializing an instance of <b>HMACBase</b> with an instance
' of the hash object.
' <p>The HMAC set of classes provided in <b>VBCorLib</b> utilize the <b>HMACBase</b> class to provide
' consistent HMAC operations.</p>
'
' @see HMACMD5
' @see HMACRIPEMD160
' @see HMACSHA1
' @see HMACSHA256
' @see HMACSHA384
' @see HMACSHA512
' @see Constructors
'
Option Explicit

Private mHashAlgorithm  As HashAlgorithm
Private mKey()          As Byte
Private mKeyIpad()      As Byte
Private mKeyOpad()      As Byte
Private mKeySize        As Long
Private mHashStarted    As Boolean
Private mHashName       As String
Private mIsDisposed     As Boolean


''
' Returns the name of the hash algorithm used.
'
' @return The hash algorithm name.
'
Public Property Get HashName() As String
    HashName = mHashName
End Property

''
' Sets the name of the hash algorithm used.
'
' @param Value The new hash name.
'
Public Property Let HashName(ByRef Value As String)
    mHashName = Value
End Property

''
' Returns the computed hash.
'
' @return The computed hash.
'
Public Property Get Hash() As Byte()
    VerifyNotDisposed
    Hash = mHashAlgorithm.Hash
End Property

''
' Sets the key the hash will use.
'
' @param Key The key to be used during the hash.
' @remarks This is required as a work-around for the byte array
' property assignment bug in VB6.
'
Public Sub SetKey(ByRef Key() As Byte)
    VerifyNotDisposed
    If mHashStarted Then _
        Throw Cor.NewCryptographicException("Cannot change key after hash has started.")
    If SAPtr(Key) = vbNullPtr Then _
        Error.ArgumentNull "Key", ArgumentNull_Array
    
    ' We copy the array to ensure a zero-based array.
    mKey = CorArray.CreateInstance(vbByte, CorArray.LengthFirstDim(Key))
    Buffer.BlockCopy Key, 0, mKey, 0, CorArray.LengthFirstDim(Key)
    Initialize
End Sub

''
' Returns the current key used by the hash.
'
' @return The current key used.
'
Public Function GetKey() As Byte()
    VerifyNotDisposed
    InitNullKey
    GetKey = mKey
End Function

''
' Computes the hash code for the specific source.
'
' @param Source The source to compute the hash for. This can be a byte array or Stream based object.
' @param Index The starting index in the byte array to begin compute the hash. This is ignored for Stream objects.
' @param Count The number of bytes to used to compute the hash in a byte array. This is ignored for Stream objects.
' @return The computed hash.
' @remarks For Stream objects, the stream is continuously read from and the hash computed until the stream contains no data.
'
Public Function ComputeHash(ByRef Source As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    VerifyNotDisposed
    StartHash
    EndHash mHashAlgorithm.ComputeHash(Source, Index, Count)
    ComputeHash = mHashAlgorithm.Hash
End Function

''
' Resets an instance to begin computing a hash from the begining.
'
Public Sub Initialize()
    VerifyNotDisposed
    Erase mKeyIpad
    Erase mKeyOpad
    mHashStarted = False
    mHashAlgorithm.Initialize
End Sub

''
' Processes a block of data towards the computed hash.
'
' @param InputBuffer The data to be processed.
' @param InputOffset The starting index to begin processing data.
' @param InputCount The number of bytes to be processed.
' @param OutputBuffer The array to place the processed data.
' @param OutputOffset The starting index to begin placing the processed data.
' @return The number of bytes processed.
' @remarks The OutputBuffer will not contain a hashed value. The data processed in the
' InputBuffer will be copied to the OutputBuffer. The data will not be copied if the
' OutputBuffer is a null array or the OutputBuffer and InputBuffer are the same array and
' the InputOffset and OutputOffset are the same.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    VerifyNotDisposed
    StartHash
    TransformBlock = mHashAlgorithm.TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

''
' Processes the final block of data, generating the computed hash.
'
' @param InputBuffer The final block of data to process.
' @param InputOffset The starting index of the data to be processed.
' @param InputCount The number of bytes to be processed.
' @return The processed data. This is a copy of the InputBuffer data.
' @remarks The returned array is a copy of the InputBuffer with a length the same as InputCount. To retrieve the
' hash value, call the Hash property.
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    VerifyNotDisposed
    StartHash
    
    Dim Result() As Byte
    Result = mHashAlgorithm.TransformFinalBlock(InputBuffer, InputOffset, InputCount)
    
    EndHash mHashAlgorithm.Hash
    TransformFinalBlock = Result
End Function

''
' Releases any resources and marks the object as disposed.
'
Public Sub Clear()
    Erase mKey
    Erase mKeyIpad
    Erase mKeyOpad
    Set mHashAlgorithm = Nothing
    mIsDisposed = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal HashAlgorithm As HashAlgorithm, ByVal KeySize As Long, ByVal HashName As String, ByRef Key As Variant)
    If HashAlgorithm Is Nothing Then _
        Throw Cor.NewArgumentNullException("HashAlgorithm")
    
    Set mHashAlgorithm = HashAlgorithm
    mKeySize = KeySize
    mHashName = HashName
    
    If Not IsMissing(Key) Then
        If VarType(Key) <> corByteArray Then
            Throw Cor.NewArgumentException("Key must be a Byte array.", "Key")
        ElseIf CorArray.IsNull(Key) Then
            Error.ArgumentNull "Key", ArgumentNull_Array
        Else
            mKey = Key
        End If
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyNotDisposed()
    If mIsDisposed Then
        Throw New ObjectDisposedException
    End If
End Sub

Private Sub InitNullKey()
    If SAPtr(mKey) = vbNullPtr Then
        mKey = CryptoHelper.GetRandomBytes(mKeySize)
    End If
End Sub

Private Sub EndHash(ByRef Hash() As Byte)
    mHashAlgorithm.TransformBlock mKeyOpad, 0, mKeySize, mKeyOpad, 0
    mHashAlgorithm.TransformFinalBlock Hash, 0, CorArray.LengthFirstDim(Hash)
    mHashStarted = False
End Sub

Private Sub StartHash()
    If Not mHashStarted Then
        ExpandKey
        mHashAlgorithm.TransformBlock mKeyIpad, 0, mKeySize, mKeyIpad, 0
        mHashStarted = True
    End If
End Sub

Private Sub ExpandKey()
    If SAPtr(mKeyIpad) <> vbNullPtr Then
        Exit Sub
    End If
    
    InitNullKey
    
    Dim K0() As Byte
    Select Case CorArray.LengthFirstDim(mKey)
        Case Is < mKeySize
            K0 = mKey
            ReDim Preserve K0(0 To mKeySize - 1)
        Case Is > mKeySize
            K0 = mHashAlgorithm.ComputeHash(mKey)
            ReDim Preserve K0(0 To mKeySize - 1)
        Case Else
            K0 = mKey
    End Select
    
    ' Generate (key Xor ipad) and (key Xor opad)
    ReDim mKeyIpad(0 To mKeySize - 1)
    ReDim mKeyOpad(0 To mKeySize - 1)
    
    Dim i As Long
    For i = 0 To mKeySize - 1
        mKeyIpad(i) = K0(i) Xor &H36
        mKeyOpad(i) = K0(i) Xor &H5C
    Next i
End Sub

