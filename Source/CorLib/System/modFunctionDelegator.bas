Attribute VB_Name = "modFunctionDelegator"
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
' Module: modFunctionDelegator
'
Option Explicit

Private Const DelegationCode        As Currency = -368956918007638.6215@     ' delegator code from Matt Curland
Private Const OffsetToFirstFunction As Long = 12

Public Type FunctionDelegator
    pVTable     As Long
    pfn         As Long
    cRefs       As Long
    Func(3)     As Long
End Type


Private mDelegatorLayout    As FunctionDelegator
Private mDelegateCode       As Currency

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
    Init
    
    With Delegator
        .pfn = pfn
        .pVTable = VarPtr(.Func(0))
        .Func(0) = FuncAddr(AddressOf InitDelegator_QueryInterface)
        .Func(1) = FuncAddr(AddressOf InitDelegator_AddRefRelease)
        .Func(2) = FuncAddr(AddressOf InitDelegator_AddRefRelease)
        .Func(3) = VarPtr(mDelegateCode)
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
    Init

    Dim This As Long
    This = CoTaskMemAlloc(LenB(mDelegatorLayout))
    If This = vbNullPtr Then _
        Throw New OutOfMemoryException
    
    mDelegatorLayout.pfn = pfn
    mDelegatorLayout.pVTable = This + OffsetToFirstFunction
        
    CopyMemory ByVal This, mDelegatorLayout, LenB(mDelegatorLayout)
    ObjectPtr(NewDelegator) = This
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    If mDelegateCode = 0 Then
        mDelegateCode = DelegationCode
        
        VirtualProtect mDelegateCode, 8, PAGE_EXECUTE_READWRITE, 0&
        
        With mDelegatorLayout
            .Func(0) = FuncAddr(AddressOf NewDelegator_QueryInterface)
            .Func(1) = FuncAddr(AddressOf NewDelegator_AddRef)
            .Func(2) = FuncAddr(AddressOf NewDelegator_Release)
            .Func(3) = VarPtr(mDelegateCode)
            .cRefs = 1
        End With
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions used by a user supplied lightweight COM function delegator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InitDelegator_QueryInterface(ByVal This As Long, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = This
End Function

Private Function InitDelegator_AddRefRelease(ByVal This As Long) As Long
    ' do nothing
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions used by a newly created lightweight COM function delegator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewDelegator_QueryInterface(ByRef This As FunctionDelegator, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
    This.cRefs = This.cRefs + 1
End Function

Private Function NewDelegator_AddRef(ByRef This As FunctionDelegator) As Long
    This.cRefs = This.cRefs + 1
    NewDelegator_AddRef = This.cRefs
End Function

Private Function NewDelegator_Release(ByRef This As FunctionDelegator) As Long
    This.cRefs = This.cRefs - 1
    If This.cRefs = 0 Then
        Call CoTaskMemFree(VarPtr(This))
    End If
    
    NewDelegator_Release = This.cRefs
End Function

