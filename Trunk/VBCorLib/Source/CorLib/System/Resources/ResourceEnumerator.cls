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
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: ResourceEnumerator
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Reader As IResourceReader, ByVal Entries As ArrayList)
    Set mReader = Reader
    Set mEntries = Entries
    Call IDictionaryEnumerator_Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyEnumerator()
    If mIndex = -1 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex > mEntries.Count Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Object.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = ObjPtr(CUnk(Me))
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Set IDictionaryEnumerator_Current = mEntry
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    Call VerifyEnumerator
    Set IDictionaryEnumerator_Entry = mEntry
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    Call VerifyEnumerator
    Set IDictionaryEnumerator_Key = mEntry.Key
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    If mIndex > mEntries.Count Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
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
    Call VerifyEnumerator
    Call Helper.MoveVariant(IDictionaryEnumerator_Value, mEntry.Value)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Set IEnumerator_Current = mEntry
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = IDictionaryEnumerator_MoveNext
End Function

Private Sub IEnumerator_Reset()
    Call IDictionaryEnumerator_Reset
End Sub
