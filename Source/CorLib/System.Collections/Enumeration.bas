Attribute VB_Name = "Enumeration"
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: Enumeration
'

''
' This module is to support the For Each method for classes that implement IEnumerator
'
'@Folder("CorLib.System.Collections")
Option Explicit

Private Const E_NOINTERFACE             As Long = &H80004002
Private Const ENUM_FINISHED             As Long = 1
Private Const IID_IUnknown_Data1        As Long = 0
Private Const IID_IEnumVariant_Data1    As Long = &H20404

Private Type EnumeratorWrapper
   pVTable      As Long
   cRefs        As Long
   Func(0 To 6) As Long
   Enumerator   As IEnumerator
End Type

Private IID_IUnknown        As VBGUID
Private IID_IEnumVariant    As VBGUID
Private mWrapperTemplate    As EnumeratorWrapper
Private mInitialized        As Boolean


Public Function CreateEnumerator(ByVal Enumerator As IEnumerator) As IUnknown
    Init
    
    Dim This As Long
    This = CoTaskMemAlloc(LenB(mWrapperTemplate))
    If This = vbNullPtr Then _
        Throw New OutOfMemoryException

    Set mWrapperTemplate.Enumerator = Enumerator
    CopyMemory ByVal This, mWrapperTemplate, LenB(mWrapperTemplate)
    ObjectPtr(mWrapperTemplate.Enumerator) = vbNullPtr
    
    ObjectPtr(CreateEnumerator) = This
End Function

Public Function GetCollectionVersion(ByVal Obj As Object) As Long
    On Error GoTo TypeMismatch
    
    Dim Versioned As IVersionable
    Set Versioned = Obj
    GetCollectionVersion = Versioned.Version
    
TypeMismatch:
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    If Not mInitialized Then
        InitGUIDS
        InitWrapperTemplate
        mInitialized = True
    End If
End Sub

Private Sub InitGUIDS()
    With IID_IEnumVariant
        .Data1 = &H20404
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With IID_IUnknown
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Sub

Private Sub InitWrapperTemplate()
    With mWrapperTemplate
        .cRefs = 1
        .Func(0) = FuncAddr(AddressOf QueryInterface)
        .Func(1) = FuncAddr(AddressOf AddRef)
        .Func(2) = FuncAddr(AddressOf Release)
        .Func(3) = FuncAddr(AddressOf IEnumVariant_Next)
        .Func(4) = FuncAddr(AddressOf IEnumVariant_Skip)
        .Func(5) = FuncAddr(AddressOf IEnumVariant_Reset)
        .Func(6) = FuncAddr(AddressOf IEnumVariant_Clone)
        .pVTable = VarPtr(.Func(0))
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  VTable functions in the IEnumVariant and IUnknown interfaces.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function QueryInterface(ByRef This As EnumeratorWrapper, ByRef riid As VBGUID, ByRef pvObj As Long) As Long
    Dim IsMatch As Boolean
    
    ' When VB queries the interface, we support only two: IUnknown, IEnumVariant
    Select Case riid.Data1
        Case IID_IEnumVariant_Data1
            IsMatch = CBool(IsEqualGUID(riid, IID_IEnumVariant))
        Case IID_IUnknown_Data1
            IsMatch = CBool(IsEqualGUID(riid, IID_IUnknown))
    End Select
    
    If IsMatch Then
        pvObj = VarPtr(This)
        AddRef This
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function

Private Function AddRef(ByRef This As EnumeratorWrapper) As Long
    This.cRefs = This.cRefs + 1
    AddRef = This.cRefs
End Function

Private Function Release(ByRef This As EnumeratorWrapper) As Long
    This.cRefs = This.cRefs - 1
    Release = This.cRefs
    
    If This.cRefs = 0 Then
        Delete This
    End If
End Function

Private Sub Delete(ByRef This As EnumeratorWrapper)
   Set This.Enumerator = Nothing
   CoTaskMemFree VarPtr(This)
End Sub

Private Function IEnumVariant_Next(ByRef This As EnumeratorWrapper, ByVal celt As Long, ByRef prgVar As Variant, ByVal pceltFetched As Long) As Long
    Const NumberOfItemsFetched As Long = 1
    
    If This.Enumerator.MoveNext Then
        Helper.MoveVariant prgVar, This.Enumerator.Current
        
        If pceltFetched <> vbNullPtr Then
            MemLong(pceltFetched) = NumberOfItemsFetched
        End If
    Else
        IEnumVariant_Next = ENUM_FINISHED
    End If
End Function

Private Function IEnumVariant_Skip(ByRef This As EnumeratorWrapper, ByVal celt As Long) As Long
    Do While celt > 0
        If Not This.Enumerator.MoveNext Then
            IEnumVariant_Skip = ENUM_FINISHED
            Exit Function
        End If
        celt = celt - 1
    Loop
End Function

Private Function IEnumVariant_Reset(ByRef This As EnumeratorWrapper) As Long
   This.Enumerator.Reset
End Function

Private Function IEnumVariant_Clone(ByRef This As EnumeratorWrapper, ByRef ppenum As stdole.IUnknown) As Long
    Dim o As ICloneable
    Set o = This.Enumerator
    Set ppenum = o.Clone
End Function
