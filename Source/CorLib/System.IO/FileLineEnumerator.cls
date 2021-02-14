VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileLineEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2021 Kelly Ethridge
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
' Module: FileLineEnumerator
'
Option Explicit
Implements IObject
Implements IEnumerable
Implements IEnumerator

Private mPath       As String
Private mEncoding   As Encoding
Private mReader     As StreamReader
Private mLine       As String
Private mFinished   As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Path As String, ByVal Encoding As Encoding)
    mPath = Path
    Set mEncoding = Encoding
End Sub

Private Sub InitReader()
    Set mReader = Cor.NewStreamReader(mPath, mEncoding)
    mFinished = False
    mLine = vbNullString
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = Me
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If mReader Is Nothing Then _
        Error.InvalidOperation InvalidOperation_EnumNotStarted
    If mFinished Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
        
    IEnumerator_Current = mLine
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mFinished Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
    
    If mReader Is Nothing Then
        InitReader
    End If
    
    mLine = mReader.ReadLine
    mFinished = CorString.IsNull(mLine)
    IEnumerator_MoveNext = Not mFinished
End Function

Private Sub IEnumerator_Reset()
    InitReader
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function
