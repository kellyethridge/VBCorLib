VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKeyEnumerator"
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
'    Module: RegistryKeyEnumerator
'

'
' This is the base implementation of an enumerator class.
'
' To use this class in a For..Each loop, it must be passed
' into the CreateEnumerator function. The result of that
' function is then returned through the NewEnum function.
'
' Public Function NewEnum() As IUnknown
'     Dim ret As New MyEnumerator
'     ret.Init Me                           ' Me would be the outer object, such as a list to be enumerated.
'     Set NewEnum = CreateEnumerator(ret)
' End Function
'
Option Explicit
Implements IObject
Implements IEnumerator
Implements IDictionaryEnumerator


Private mKey            As RegistryKey
Private mIndex          As Long
Private mNames()        As String
Private mEnumType       As IDictionaryEnumType
Private mOptions        As RegistryValueOptions



''
' We place this function here to allow the caller to
' use the GetEnumerator function directly when starting
' a For..Each loop. The user may need to alter the enumeration
' with the parameters, so the NewEnum in the RegistryKey
' object will not get called. Instead, this one will be called.
'
' @return For..Each compatible enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Key As RegistryKey, ByVal EnumType As IDictionaryEnumType, ByVal Options As RegistryValueOptions)
    mEnumType = EnumType
    mOptions = Options
    Set mKey = Key
    Call IEnumerator_Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyEnum()
    If mIndex < LBound(mNames) Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex > UBound(mNames) Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Call Helper.MoveVariant(IDictionaryEnumerator_Current, IEnumerator_Current)
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    Call VerifyEnum
    If mEnumType = detEntries Then Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(mNames(mIndex), mKey.GetValue(mNames(mIndex), , mOptions))
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    Call VerifyEnum
    If mEnumType <> detValues Then IDictionaryEnumerator_Key = mNames(mIndex)
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    IDictionaryEnumerator_MoveNext = IEnumerator_MoveNext
End Function

Private Sub IDictionaryEnumerator_Reset()
    Call IEnumerator_Reset
End Sub

Private Property Get IDictionaryEnumerator_Value() As Variant
    Call VerifyEnum
    If mEnumType <> detKeys Then Call Helper.MoveVariant(IDictionaryEnumerator_Value, mKey.GetValue(mNames(mIndex), , mOptions))
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Call VerifyEnum
    
    Select Case mEnumType
        Case detEntries
            Set IEnumerator_Current = Cor.NewDictionaryEntry(mNames(mIndex), mKey.GetValue(mNames(mIndex), , mOptions))
            Exit Function
            
        Case detValues, detKeys
            IEnumerator_Current = mNames(mIndex)
    End Select
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mIndex > UBound(mNames) Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    mIndex = mIndex + 1
    IEnumerator_MoveNext = (mIndex <= UBound(mNames))
End Function

Private Sub IEnumerator_Reset()
    mIndex = -1
    
    Select Case mEnumType
        Case detEntries: mNames = mKey.GetValueNames
        Case detKeys:    mNames = mKey.GetSubKeyNames
        Case detValues:  mNames = mKey.GetValueNames
    End Select
End Sub


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

