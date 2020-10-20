VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullTextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: NullTextReader
'

''
' Provides a TextReader with no data backing.
'
' @see TextReader
'
'@Folder("CorLib.System.IO")
Option Explicit
Implements IObject
Implements TextReader


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   TextReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextReader_CloseReader()
' do nothing
End Sub

Private Function TextReader_Equals(Value As Variant) As Boolean
    TextReader_Equals = IObject_Equals(Value)
End Function

Private Function TextReader_GetHashCode() As Long
    TextReader_GetHashCode = IObject_GetHashCode
End Function

Private Function TextReader_Peek() As Long
    TextReader_Peek = -1
End Function

Private Function TextReader_Read() As Long
    TextReader_Read = -1
End Function

Private Function TextReader_ReadBlock(Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
' do nothing
End Function

Private Function TextReader_ReadLine() As String
    TextReader_ReadLine = vbNullString
End Function

Private Function TextReader_ReadToEnd() As String
    TextReader_ReadToEnd = vbNullString
End Function

Private Function TextReader_ToString() As String
    TextReader_ToString = IObject_ToString
End Function
