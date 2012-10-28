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
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Friend License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Friend License for more details.
'
'    You should have received a copy of the GNU Library General Friend License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: CryptoAPIHash
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal HashAlg As Long, ByVal HashSize As Long)
    mHashSize = HashSize \ 8
    mProvider = CryptoAPI.DefaultProvider
    mHashAlg = HashAlg
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    If mHasher <> vbNullPtr Then Call CryptDestroyHash(mHasher)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IHashAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IHashAlgorithm_HashCore(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    If CryptHashData(mHasher, Bytes(Index), Count, 0) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not perform hash on data.")
End Sub

Private Function IHashAlgorithm_HashFinal(ByRef Buffer() As Byte, ByVal ByteCount As Long, ByVal MessageLength As Currency) As Byte()
    Dim Ret() As Byte
    ReDim Ret(0 To mHashSize - 1)
    
    If CryptGetHashParam(mHasher, HP_HASHVAL, Ret(0), mHashSize, 0) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not retrieve hash value from service provider. " & GetErrorMessage(Err.LastDllError))
    
    IHashAlgorithm_HashFinal = Ret
End Function

Private Sub IHashAlgorithm_Initialize()
    If mHasher <> 0 Then Call CryptDestroyHash(mHasher)
    
    If CryptCreateHash(mProvider, mHashAlg, 0, 0, mHasher) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException("Could not create hash algorithm.")
End Sub
