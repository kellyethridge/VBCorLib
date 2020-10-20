VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA384Static"
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
' Module: SHA384Static
'

''
' Provides methods associated with the SHA-384 class.
'
'@Folder("CorLib.System.Security.Cryptography")
Option Explicit

Private mStateH() As Long
Private mStateL() As Long
Private mStatePtr As DLong


''
' Creates a new instance of an SHA-384 implementation.
'
' @param AlgName The name of the implementation to create an instance of.
' @return A new SHA-384 compatible implementation.
' @remarks If not name is given, then a new SHA384Managed is returned.
'
Public Function Create(Optional ByRef AlgName As String = "CorLib.SHA384") As SHA384
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function

Friend Property Get StatePtr() As DLong
    If mStatePtr.LoDWord = vbNullPtr Then
        mStateH = Cor.NewLongs(&HCBBB9D5D, &H629A292A, &H9159015A, &H152FECD8, &H67332667, &H8EB44A87, &HDB0C2E0D, &H47B5481D)
        mStateL = Cor.NewLongs(&HC1059ED8, &H367CD507, &H3070DD17, &HF70E5939, &HFFC00B31, &H68581511, &H64F98FA7, &HBEFA4FA4)
        
        mStatePtr.HiDWord = VarPtr(mStateH(0))
        mStatePtr.LoDWord = VarPtr(mStateL(0))
    End If
    
    StatePtr = mStatePtr
End Property
