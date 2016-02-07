VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilderCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: StringBuilderCache
'
Option Explicit

Private mCache As New Stack


Private Sub Class_Initialize()
    mCache.Push New StringBuilder
    mCache.Push New StringBuilder
End Sub

Friend Function Acquire() As StringBuilder
    If mCache.Count = 0 Then
        Set Acquire = New StringBuilder
    Else
        Set Acquire = mCache.Pop
        Acquire.Length = 0
    End If
End Function

Friend Function GetStringAndRelease(ByVal Builder As StringBuilder) As String
    Debug.Assert Not Builder Is Nothing
    
    If Builder.Length > 0 Then
        GetStringAndRelease = Builder.ToString
    End If
    
    Release Builder
End Function

Friend Sub Release(ByVal Builder As StringBuilder)
    Debug.Assert Not Builder Is Nothing
    
    mCache.Push Builder
End Sub


