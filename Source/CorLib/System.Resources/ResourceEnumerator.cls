VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceEnumerator"
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
' Module: ResourceEnumerator
'

Option Explicit
Implements IObject
Implements IEnumerator
Implements IDictionaryEnumerator

Private mReader     As IResourceReader
Private mEntries    As ArrayList
Private mIndex      As Long
Private mEntry      As DictionaryEntry



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Reader As IResourceReader, ByVal Entries As ArrayList)
    Set mReader = Reader
    Set mEntries = Entries
    IDictionaryEnumerator_Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyEnumerator()
    If mIndex = -1 Then _
        Error.InvalidOperation InvalidOperation_EnumNotStarted
    If mIndex > mEntries.Count Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Set IDictionaryEnumerator_Current = mEntry
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    VerifyEnumerator
    Set IDictionaryEnumerator_Entry = mEntry
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    VerifyEnumerator
    Set IDictionaryEnumerator_Key = mEntry.Key
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    If mIndex > mEntries.Count Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
    
    mIndex = mIndex + 1
    
    If mIndex = mEntries.Count Then
        mIndex = mEntries.Count + 1
        Exit Function
    End If
    
    Set mEntry = mEntries(mIndex)
    IDictionaryEnumerator_MoveNext = True
End Function

Private Sub IDictionaryEnumerator_Reset()
    mIndex = -1
End Sub

Private Property Get IDictionaryEnumerator_Value() As Variant
    VerifyEnumerator
    Helper.MoveVariant IDictionaryEnumerator_Value, mEntry.Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = mEntry
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = IDictionaryEnumerator_MoveNext
End Function

Private Sub IEnumerator_Reset()
    IDictionaryEnumerator_Reset
End Sub
