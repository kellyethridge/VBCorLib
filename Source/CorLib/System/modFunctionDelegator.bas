Attribute VB_Name = "modFunctionDelegator"
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
'    Module: modFunctionDelegator
'
Option Explicit

Private Const DELEGATE_ASM As Currency = -368956918007638.6215@     ' from Matt Curland

''
' The structure of the lightweight COM function delegator object.

Public Type FunctionDelegator
    pVTable As Long
    pfn As Long
    cRefs As Long
    Func(3) As Long
End Type

''
' holds the ASM code used to delegate to the function address
Private mDelegateASM As Currency

''
' holds a pointer to the delegate asm code
Private mAsm As Long

''
' holds the addresses to the VTable functions
'
Private mInitDelegatorQueryInterface    As Long
Private mInitDelegatorAddRelease        As Long
Private mNewDelegatorQueryInterface     As Long
Private mNewDelegatorAddRef             As Long
Private mNewDelegatorRelease            As Long


''
' Returns an object using a supplied function delegator structure
' as a lightweight COM object. This will not remove itself from
' memory once all references have been removed.
'
' @param Delegator A user supplied function delegator structure
' @param pfn The address to the function to be called.
' @return A lightweight object that can be used to call the function address.
'
Public Function InitDelegator(ByRef Delegator As FunctionDelegator, Optional ByVal pfn As Long = 0) As IUnknown
    If mAsm = 0 Then Init
    
    With Delegator
        .pfn = pfn
        .pVTable = VarPtr(.Func(0))
        .Func(0) = mInitDelegatorQueryInterface
        .Func(1) = mInitDelegatorAddRelease
        .Func(2) = mInitDelegatorAddRelease
        .Func(3) = mAsm
        .pfn = pfn
    End With
    
    ObjectPtr(InitDelegator) = VarPtr(Delegator)
End Function


''
' This creates a new lightweight function delegator object
' that calls a user specified function using the AddressOf operator.
'
' @param pfn The address to function to be called.
' @return A lightweight COM object used to call a function.
'
Public Function NewDelegator(ByVal pfn As Long) As IUnknown
    Dim this As Long
    Dim Struct As FunctionDelegator
    
    If mAsm = 0 Then Init

    this = CoTaskMemAlloc(LenB(Struct))
    If this = 0 Then Throw New OutOfMemoryException
    
    With Struct
        .pVTable = this + 12
        .Func(0) = mNewDelegatorQueryInterface
        .Func(1) = mNewDelegatorAddRef
        .Func(2) = mNewDelegatorRelease
        .Func(3) = mAsm
        .pfn = pfn
        .cRefs = 1
    End With
        
    CopyMemory ByVal this, Struct, LenB(Struct)
    ObjectPtr(NewDelegator) = this
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    mDelegateASM = DELEGATE_ASM
    mAsm = VarPtr(mDelegateASM)
    
    Call VirtualProtect(mDelegateASM, 8, PAGE_EXECUTE_READWRITE, 0&)
    
    mInitDelegatorQueryInterface = FuncAddr(AddressOf InitDelegator_QueryInterface)
    mInitDelegatorAddRelease = FuncAddr(AddressOf InitDelegator_AddRefRelease)
    mNewDelegatorQueryInterface = FuncAddr(AddressOf NewDelegator_QueryInterface)
    mNewDelegatorAddRef = FuncAddr(AddressOf NewDelegator_AddRef)
    mNewDelegatorRelease = FuncAddr(AddressOf NewDelegator_Release)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions used by a user supplied lightweight COM function delegator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InitDelegator_QueryInterface(ByVal this As Long, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = this
End Function

Private Function InitDelegator_AddRefRelease(ByVal this As Long) As Long
    ' do nothing
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions used by a newly created lightweight COM function delegator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewDelegator_QueryInterface(ByRef this As FunctionDelegator, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(this)
    NewDelegator_AddRef this
End Function
Private Function NewDelegator_AddRef(ByRef this As FunctionDelegator) As Long
    With this
        .cRefs = .cRefs + 1
        NewDelegator_AddRef = .cRefs
    End With
End Function
Private Function NewDelegator_Release(ByRef this As FunctionDelegator) As Long
    With this
        .cRefs = .cRefs - 1
        NewDelegator_Release = .cRefs
        If .cRefs = 0 Then CoTaskMemFree VarPtr(this)
    End With
End Function

