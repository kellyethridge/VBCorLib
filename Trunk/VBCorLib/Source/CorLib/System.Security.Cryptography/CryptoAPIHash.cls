VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoAPIHash"
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
' Module: CryptoAPIHash
'

''
' Wraps the CryptoAPI hash functions to be used in the HashAlgorithmBase callback methods.
'
Option Explicit
Implements IHashAlgorithm

Private mProvider       As Long
Private mHasher         As Long
Private mHashAlg        As Long
Private mHashSize       As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal HashAlg As Long, ByVal HashSize As Long)
    mHashSize = HashSize \ 8
    mProvider = CryptoAPI.DefaultProvider
    mHashAlg = HashAlg
End Sub

Private Sub Class_Terminate()
    DestroyHasher
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DestroyHasher()
    If mHasher <> vbNullPtr Then
        CryptDestroyHash mHasher
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IHashAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IHashAlgorithm_HashCore(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    If CryptHashData(mHasher, Bytes(Index), Count, 0) = BOOL_FALSE Then
        Throw Cor.NewCryptographicException("Could not perform hash on data->" & GetMessage(Err.LastDllError))
    End If
End Sub

Private Function IHashAlgorithm_HashFinal(ByRef Buffer() As Byte, ByVal ByteCount As Long, ByVal MessageLength As Currency) As Byte()
    Dim Result() As Byte
    ReDim Result(0 To mHashSize - 1)
    
    If CryptGetHashParam(mHasher, HP_HASHVAL, Result(0), mHashSize, 0) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not retrieve hash value from service provider->" & GetMessage(Err.LastDllError))
    
    IHashAlgorithm_HashFinal = Result
End Function

Private Sub IHashAlgorithm_Initialize()
    DestroyHasher
    
    If CryptCreateHash(mProvider, mHashAlg, 0, 0, mHasher) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not create hash algorithm->" & GetMessage(Err.LastDllError))
End Sub
