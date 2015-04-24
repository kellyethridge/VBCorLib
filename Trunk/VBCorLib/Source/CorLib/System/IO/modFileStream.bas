Attribute VB_Name = "modFileStream"
'    CopyRight (c) 2008 Kelly Ethridge
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
'    Module: modFileStream
'

Option Explicit

'Public Declare Function ReadFileEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpOverlapped As Any, ByVal lpCompletionRoutine As Long) As Long
'Public Declare Function WriteFileEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpOverlapped As NativeOverlapped, ByVal lpCompletionRoutine As Long) As Long
'
'
'Public Sub ReadFileIOCompletion(ByVal dwErrorCode As Long, ByVal dwNumberOfBytesTransfered As Long, ByVal lpOverlapped As Long)
'    Dim over As NativeOverlapped
'    Call CopyMemory(over, ByVal lpOverlapped, Len(over))
'    Call CoTaskMemFree(lpOverlapped)
'
'    Dim Callback As AsyncCallback
'    ObjectPtr(Callback) = over.CallbackHandle
'
'    Dim async As StreamAsyncResult
'    ObjectPtr(async) = over.EventHandle
'    async.BytesRead = dwNumberOfBytesTransfered
'    async.IsCompleted = True
'
'    If Not Callback Is Nothing Then
'        Call Callback.Execute(async)
'    End If
'End Sub
