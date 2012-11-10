Attribute VB_Name = "modWeakReferenceHelpers"
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
'    Module: modWeakReferenceHelpers
'

''
' Provides a mechanism to keep track of a weak reference to an object
' without keeping that object alive, then retrieving a strong reference
' if the object is still alive.
'
' @remarks We watch the reference count from the Release method. This method
' is called everytime a variable that has a reference to the watched object
' goes out of scope. Once the reference count reaches zero, then we detach
' from the object and set our flag to no longer being alive.
'
' You can learn about this technique from Matthew Curlands excellent book -
' "Advanced Visual Basic 6: Power Techniques for Everyday Programs"
'
Option Explicit

' used for quick interface comparisons.
Private Const IID_IUnknown_Data1            As Long = 0
Private Const IID_IProvideClassInfo_Data1   As Long = &HB196B283

' our lightweight object that replaces the existing VTable.
Public Type WeakRefHookType
    VTable(3) As Long
    pOriginalVTable As Long
    Target As IProvideClassInfo
    pOwner As Long
End Type

' Used to access a WeakRefHookType through a pointer.
Private Type WeakSafeArray
    pVTable As Long
    This As IUnknown
    pRelease As Long
    SA As SafeArray1d
    WeakRef() As WeakRefHookType
End Type



' Guids for interfaces we support locally.
Private IID_IUnknown            As VBGUID
Private IID_IProvideClassInfo   As VBGUID

Private mWeak As WeakSafeArray



''
' Initialize a new weak reference that will become the new
' hook into the VTable so we can watch the Release calls.
'
' @param Weak The temporary VTable and flags for the object being watched.
' @param owner The WeakReference object that maintains the hook and returns a strong reference.
' @param Target The object to maintain a weak reference to without keeping it alive in memory.
'
Public Function InitWeakReference(ByVal Owner As WeakReference, ByVal Target As IUnknown) As Long
    Dim Weak As WeakRefHookType
    If mWeak.pVTable = 0 Then
        IID_IProvideClassInfo = GUIDFromString("{B196B283-BAB4-101A-B69C-00AA00341D07}")
        IID_IUnknown = GUIDFromString("{00000000-0000-0000-C000-000000000046}")
        
        With mWeak
            .pRelease = FuncAddr(AddressOf WeakReferenceArray_Release)
            .pVTable = VarPtr(.pVTable)
            ObjectPtr(.This) = VarPtr(mWeak)
            SAPtr(.WeakRef) = VarPtr(.SA)
            
            With .SA
                .cbElements = Len(Weak)
                .cDims = 1
                .cElements = 1
            End With
        End With
    End If
    
    Dim This As Long
    This = CoTaskMemAlloc(LenB(Weak))
    
    ' Since all the Exception classes use a WeakReference
    ' object, we can't throw an exception object, because it
    ' will need to create the WeakReference. And if we have
    ' failed to create this WeakReference, we will most certainly
    ' fail to create the WeakReferences for any Exceptions thrown.
    If This = 0 Then Err.Raise 7    ' don't use OutOfMemoryException since it may fail to create.
    
    With Weak
        .VTable(0) = FuncAddr(AddressOf WeakReference_QueryInterface)
        .VTable(1) = FuncAddr(AddressOf WeakReference_AddRef)
        .VTable(2) = FuncAddr(AddressOf WeakReference_Release)
        .VTable(3) = FuncAddr(AddressOf WeakReference_GetClassInfo)
        
        Dim pUnk As Long
        pUnk = MemLong(VarPtr(Target))
        
        Set Target = Nothing
        
        MemLong(VarPtr(.Target)) = pUnk
        .pOriginalVTable = MemLong(pUnk)
        MemLong(pUnk) = This
        
        .pOwner = ObjPtr(Owner)
    End With
    
    Call CopyMemory(ByVal This, Weak, LenB(Weak))
    Call ZeroMemory(Weak, LenB(Weak))
    InitWeakReference = This
End Function

''
' Handles the initial interface queries and delegates them to the target object.
'
' @param this The pointer to the controlling IUnknown VTable.
' @param riid The GUID of the requested intereface.
' @param pvObj An out-pointer to the the location of the object that implements the requested interface.
' @return S_OK is returned on success, otherwise E_NOINTERFACE.
' @remarks This is the function used in the VTable QueryInterface.
'
Private Function WeakReference_QueryInterface(ByRef This As Long, ByRef riid As VBGUID, ByRef pvObj As Long) As Long
    Dim OldVTable As Long
    
    OldVTable = This
    pvObj = 0
    
    mWeak.SA.pvData = This
    With mWeak.WeakRef(0)
        This = .pOriginalVTable
        WeakReference_QueryInterface = .Target.QueryInterface(riid, pvObj)
        If pvObj <> 0 Then
            If pvObj = VarPtr(This) Then
                Dim fOK As Boolean
                Select Case riid.Data1
                    Case IID_IUnknown_Data1
                        fOK = CBool(IsEqualGUID(riid, IID_IUnknown))
                    Case IID_IProvideClassInfo_Data1
                        fOK = CBool(IsEqualGUID(riid, IID_IProvideClassInfo))
                End Select
                If Not fOK Then
                    .Target.Release
                    pvObj = 0
                    WeakReference_QueryInterface = E_NOINTERFACE
                End If
            End If
        End If
    End With
    This = OldVTable
End Function

''
' Adds a new reference to the existing object.
'
' @param this The pointer to the controlling IUnknown VTable.
' @return The number of references so far.
'
Private Function WeakReference_AddRef(ByRef This As Long) As Long
    Dim OldVTable As Long
    OldVTable = This
    
    mWeak.SA.pvData = This
    
    With mWeak.WeakRef(0)
        This = .pOriginalVTable
        WeakReference_AddRef = .Target.AddRef
    End With
    
    This = OldVTable
End Function

''
' Releases a reference from the existing object. If it reaches zero
' then the weak reference is also released.
'
' @param this The pointer to the controllin IUnknown VTable.
' @return The number of references so far.
'
Private Function WeakReference_Release(ByRef This As Long) As Long
    Dim OldVTable As Long
    OldVTable = This
    
    With mWeak
        .SA.pvData = This
    
        With .WeakRef(0)
            This = .pOriginalVTable
            
            If Not .Target Is Nothing Then
                WeakReference_Release = .Target.Release
            End If
            
            If (WeakReference_Release > 0) And (.pOwner <> 0) Then
                This = OldVTable
            Else
                ObjectPtr(.Target) = 0
                If .pOwner <> 0 Then
                    Dim Owner As WeakReference
                    ObjectPtr(Owner) = .pOwner
                    Owner.Dispose
                    ObjectPtr(Owner) = 0
                    .pOwner = 0
                End If
                Call CoTaskMemFree(This)
            End If
        End With
    End With
End Function

''
' VB Object implement the IProvideClassInfo interface and we must
' be able to delegate to the function returning class info. Such info
' derived from this is the TypeName of the class.
'
' @param this A pointer to the controlling IUnknown.
' @param ppTypeInfo A pointer to the ITypeInfo object.
' @return Error codes.
'
Private Function WeakReference_GetClassInfo(ByRef This As Long, ByRef ppTypeInfo As Long) As Long
    Dim OldVTable As Long
    OldVTable = This
    
    mWeak.SA.pvData = This
    
    With mWeak.WeakRef(0)
        This = .pOriginalVTable
        WeakReference_GetClassInfo = .Target.GetClassInfo(ppTypeInfo)
    End With
    
    This = OldVTable
End Function


''
' Used to kill the mWeak.WeakRef array connection.
'
Private Function WeakReferenceArray_Release(ByVal This As Long) As Long
    SAPtr(mWeak.WeakRef) = 0
End Function
