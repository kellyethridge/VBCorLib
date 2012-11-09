Attribute VB_Name = "modMain"
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
'    Module: modMain
'
Option Explicit

Private mInIDE        As Boolean
Private mInDebugger   As Boolean

Public Property Get InIDE() As Boolean
    InIDE = mInIDE
End Property

Public Property Get InDebugger() As Boolean
    InDebugger = mInDebugger
End Property

Private Sub Main()
    Call SetInIDE
    Call SetInDebugger
    Call InitWin32Api
    Call InitPublicFunctions
    Call InitcDateTimeHelpers
    Call InitEncodingHelpers
End Sub

''
' This is to determine if the compiled dll is being used in the
' VB6 IDE by another project. This is primarily so that the console
' class can disable the exit button if we are running in an IDE.
'
Private Sub SetInDebugger()
    Dim Result As String
    Result = String$(1024, 0)
    
    Call GetModuleFileName(vbNullPtr, Result, Len(Result))
    
    Dim i As Long
    i = InStr(Result, vbNullChar)
    
    Result = Left$(Result, i - 1)
    
    mInDebugger = (UCase$(Right$(Result, 8)) = "\VB6.EXE")
End Sub

Private Sub SetInIDE()
    On Error GoTo errTrap
    Debug.Assert 1 \ 0
    Exit Sub
    
errTrap:
    mInIDE = True
End Sub
