Attribute VB_Name = "modIEnumerator"
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: modIEnumerator
'

Option Explicit

Private Const E_NOINTERFACE As Long = &H80004002
Private Const ENUM_FINISHED As Long = 1


' This is the type that will wrap the user enumerator.
' When a new IEnumVariant compatible object is created,
' it will have the internal structure of UserEnumWrapperType
Private Type UserEnumWrapperType
   pVTable As Long
   cRefs As Long
   UserEnum As IEnumerator
End Type

' This is an array of pointers to functions that the
' object's VTable will point to.
Private Type VTable
   Functions(0 To 6) As Long
End Type



' The created VTable of function pointers
Private mVTable As VTable

' Pointer to the mVTable memory address.
Private mpVTable As Long

' GUIDs to identify IUnknown and IEnumVariant when
' the interface is queried.
Private IID_IUnknown As VBGUID
Private Const IID_IUnknown_Data1 As Long = 0
Private IID_IEnumVariant As VBGUID
Private Const IID_IEnumVariant_Data1 As Long = &H20404



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Public Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Creates the LightWeight object that will wrap the user's enumerator.
Public Function CreateEnumerator(ByVal Obj As IEnumerator) As stdole.IUnknown
    Dim this    As Long
    Dim Struct  As UserEnumWrapperType
    
    If mpVTable = 0 Then Call Init
    
    ' allocate memory to place the new object.
    this = CoTaskMemAlloc(Len(Struct))
    If this = vbNullPtr Then Throw New OutOfMemoryException
    
    ' fill the structure of the new wrapper object
    With Struct
        Set .UserEnum = Obj
        .cRefs = 1
        .pVTable = mpVTable
    End With
    
    ' move the structure to the allocated memory to complete the object
    Call CopyMemory(ByVal this, ByVal VarPtr(Struct), LenB(Struct))
    Call ZeroMemory(ByVal VarPtr(Struct), LenB(Struct))
    
    ' assign the return value to the newly create object.
    ObjectPtr(CreateEnumerator) = this
End Function



' setup the guids and vtable function pointers.
Private Sub Init()
    Call InitGUIDS
    Call InitVTable
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

Private Sub InitVTable()
    With mVTable
        .Functions(0) = FuncAddr(AddressOf QueryInterface)
        .Functions(1) = FuncAddr(AddressOf AddRef)
        .Functions(2) = FuncAddr(AddressOf Release)
        .Functions(3) = FuncAddr(AddressOf IEnumVariant_Next)
        .Functions(4) = FuncAddr(AddressOf IEnumVariant_Skip)
        .Functions(5) = FuncAddr(AddressOf IEnumVariant_Reset)
        .Functions(6) = FuncAddr(AddressOf IEnumVariant_Clone)
        
        mpVTable = VarPtr(.Functions(0))
   End With
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  VTable functions in the IEnumVariant and IUnknown interfaces.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' When VB queries the interface, we support only two.
' IUnknown
' IEnumVariant
Private Function QueryInterface(ByRef this As UserEnumWrapperType, _
                                ByRef riid As VBGUID, _
                                ByRef pvObj As Long) As Long
    Dim ok As BOOL
    
    Select Case riid.Data1
        Case IID_IEnumVariant_Data1
            ok = IsEqualGUID(riid, IID_IEnumVariant)
        Case IID_IUnknown_Data1
            ok = IsEqualGUID(riid, IID_IUnknown)
    End Select
    
    If ok Then
        pvObj = VarPtr(this)
        Call AddRef(this)
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function


' increment the number of references to the object.
Private Function AddRef(ByRef this As UserEnumWrapperType) As Long
    With this
        .cRefs = .cRefs + 1
        AddRef = .cRefs
    End With
End Function


' decrement the number of references to the object, checking
' to see if the last reference was released.
Private Function Release(ByRef this As UserEnumWrapperType) As Long
    With this
        .cRefs = .cRefs - 1
        Release = .cRefs
        If .cRefs = 0 Then Call Delete(this)
    End With
End Function


' cleans up the lightweight objects and releases the memory
Private Sub Delete(ByRef this As UserEnumWrapperType)
   Set this.UserEnum = Nothing
   Call CoTaskMemFree(VarPtr(this))
End Sub


' move to the next element and return it, signaling if we have reached the end.
Private Function IEnumVariant_Next(ByRef this As UserEnumWrapperType, ByVal celt As Long, ByRef prgVar As Variant, ByVal pceltFetched As Long) As Long
    If this.UserEnum.MoveNext Then
        Call Helper.MoveVariant(prgVar, this.UserEnum.Current)
        
        ' check to see if the pointer is valid (not zero)
        ' before we write to that memory location.
        If pceltFetched Then
            MemLong(pceltFetched) = 1
        End If
    Else
        IEnumVariant_Next = ENUM_FINISHED
    End If
End Function


' skip the requested number of elements as long as we don't run out of them.
Private Function IEnumVariant_Skip(ByRef this As UserEnumWrapperType, ByVal celt As Long) As Long
    Do While celt > 0
        If this.UserEnum.MoveNext = False Then
            IEnumVariant_Skip = ENUM_FINISHED
            Exit Function
        End If
        celt = celt - 1
    Loop
End Function


' request the user enum to reset.
Private Function IEnumVariant_Reset(ByRef this As UserEnumWrapperType) As Long
   Call this.UserEnum.Reset
End Function


' we just return a reference to the original object.
Private Function IEnumVariant_Clone(ByRef this As UserEnumWrapperType, ByRef ppenum As stdole.IUnknown) As Long
    Dim o As ICloneable
    Set o = this.UserEnum
    Set ppenum = o.Clone
End Function
