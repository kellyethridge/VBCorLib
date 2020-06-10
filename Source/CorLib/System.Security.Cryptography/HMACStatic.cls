VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: HMACStatic
'

''
' Provides static methods associated to the HMAC class.
'
' @see HMACMD5
' @see HMACSHA1
' @see HMACSHA256
' @see HMACSHA384
' @see HMACSHA512
' @see HMACRIPEMD160
'
Option Explicit

''
' Returns an HMAC implementation based on the algorithm name provided.
'
' @param AlgName The name of the HMAC implementation to create.
' @return Returns an HMAC implementation based on the algorithm.
' @remarks The default implementation is HMACSHA1 based on the SHA1 hash algorithm.
' <p>Valid algorithm names are:
' <pre>
' HMACSHA1
' HMACSHA256
' HMACSHA384
' HMACSHA512
' HMACMD5
' HMACRIPEMD160
' </pre></p>
' <p> If a different implementation is desired, then use the standard <server>.<class> naming convention.</p>
'
Public Function Create(Optional ByVal AlgName As String = "HMACSHA1") As HMAC
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function
