Attribute VB_Name = "Delegation"
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: Delegation
'
Option Explicit

Private Const DelegationCode        As Currency = -368956918007638.6215@     ' delegator code from Matt Curland
Private Const OffsetToFirstFunction As Long = 12
Private Const SizeOfDelegate        As Long = 28

Public Type Delegate
    pVTable     As Long
    pfn         As Long
    cRefs       As Long
    Func(3)     As Long
End Type


Private mDelegateStructure  As FunctionDelegator
Private mDelegateCode       As Currency


''
' This creates a new lightweight function delegator object
' that calls a user specified function using the AddressOf operator.
'
' @param pfn The address to function to be called.
' @return A lightweight COM object used to call a function.
'
Public Function NewDelegate(ByVal pfn As Long) As IUnknown
    Init

    Dim This As Long
    This = CoTaskMemAlloc(SizeOfDelegate)
    If This = vbNullPtr Then _
        Throw New OutOfMemoryException
    
    mDelegateStructure.pfn = pfn
    mDelegateStructure.pVTable = This + OffsetToFirstFunction
        
    CopyMemory ByVal This, mDelegateStructure, SizeOfDelegate
    ObjectPtr(NewDelegate) = This
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    If mDelegateCode = 0@ Then
        mDelegateCode = DelegationCode
        
        VirtualProtect mDelegateCode, 8, PAGE_EXECUTE_READWRITE, 0&
        
        With mDelegateStructure
            .Func(0) = FuncAddr(AddressOf Delegate_QueryInterface)
            .Func(1) = FuncAddr(AddressOf Delegate_AddRef)
            .Func(2) = FuncAddr(AddressOf Delegate_Release)
            .Func(3) = VarPtr(mDelegateCode)
            .cRefs = 1
        End With
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions used by a newly created lightweight COM function delegate
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Delegate_QueryInterface(ByRef This As Delegate, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
    This.cRefs = This.cRefs + 1
End Function

Private Function Delegate_AddRef(ByRef This As Delegate) As Long
    This.cRefs = This.cRefs + 1
    Delegate_AddRef = This.cRefs
End Function

Private Function Delegate_Release(ByRef This As Delegate) As Long
    This.cRefs = This.cRefs - 1
    
    If This.cRefs = 0 Then
        CoTaskMemFree VarPtr(This)
    Else
        Delegate_Release = This.cRefs
    End If
End Function


