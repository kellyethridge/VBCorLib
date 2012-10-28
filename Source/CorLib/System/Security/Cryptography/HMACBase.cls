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
'    Module: HMACBase
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
' @param RHS The new hash name.
'
Public Property Let HashName(ByVal RHS As String)
    mHashName = HashName
End Property

''
' Returns the computed hash.
'
' @return The computed hash.
'
Public Property Get Hash() As Byte()
    Call VerifyNotDisposed
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
    Call VerifyNotDisposed
    If mHashStarted Then _
        Throw Cor.NewCryptographicException("Cannot change key after hash has started.")
    If cArray.IsNull(Key) Then _
        Throw Cor.NewArgumentNullException("Invalid HMAC Key.", "Key")
    
    mKey = cArray.CreateInstance(ciByte, cArray.GetLength(Key))
    Call Buffer.BlockCopy(Key, 0, mKey, 0, cArray.GetLength(Key))
    Call Initialize
End Sub

''
' Returns the current key used by the hash.
'
' @return The current key used.
'
Public Function GetKey() As Byte()
    Call VerifyNotDisposed
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
    Call VerifyNotDisposed
    Call StartHash
    Call EndHash(mHashAlgorithm.ComputeHash(Source, Index, Count))
    ComputeHash = mHashAlgorithm.Hash
End Function

''
' Resets an instance to begin computing a hash from the begining.
'
Public Sub Initialize()
    Call VerifyNotDisposed
    Erase mKeyIpad
    Erase mKeyOpad
    mHashStarted = False
    Call mHashAlgorithm.Initialize
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
    Call VerifyNotDisposed
    Call StartHash
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
    Call VerifyNotDisposed
    Call StartHash
    
    Dim Ret() As Byte
    Ret = mHashAlgorithm.TransformFinalBlock(InputBuffer, InputOffset, InputCount)
    
    Call EndHash(mHashAlgorithm.Hash)
    TransformFinalBlock = Ret
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal HashAlgorithm As HashAlgorithm, ByVal KeySize As Long, ByVal HashName As String, ByRef Key As Variant)
    If HashAlgorithm Is Nothing Then _
        Throw Cor.NewArgumentNullException("HashAlgorithm cannot be Nothing.", "HashAlgorithm")
    
    Set mHashAlgorithm = HashAlgorithm
    mKeySize = KeySize
    mHashName = HashName
    
    If IsMissing(Key) Then
        mKey = CryptoHelper.GetRandomBytes(mKeySize)
    ElseIf VarType(Key) <> vbByteArray Then
        Throw Cor.NewArgumentException("Key must be a Byte array.", "Key")
    ElseIf cArray.IsNull(Key) Then
        Throw Cor.NewArgumentNullException("Key cannot be Null.", "Key")
    Else
        mKey = Key
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyNotDisposed()
    If mIsDisposed Then Throw New ObjectDisposedException
End Sub

Private Sub EndHash(ByRef Hash() As Byte)
    Call mHashAlgorithm.TransformBlock(mKeyOpad, 0, mKeySize, mKeyOpad, 0)
    Call mHashAlgorithm.TransformFinalBlock(Hash, 0, cArray.GetLength(Hash))
    
    mHashStarted = False
End Sub

Private Sub StartHash()
    If mHashStarted Then Exit Sub
    
    Call ExpandKey
    Call mHashAlgorithm.TransformBlock(mKeyIpad, 0, mKeySize, mKeyIpad, 0)
    
    mHashStarted = True
End Sub

Private Sub ExpandKey()
    If Not cArray.IsNull(mKeyIpad) Then Exit Sub
    
    Dim K0() As Byte
    Select Case cArray.GetLength(mKey)
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

