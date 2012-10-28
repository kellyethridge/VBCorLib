Attribute VB_Name = "modConsoleCallbacks"
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
'    Module: modConsoleCallbacks
'

Option Explicit

''
' We keep these flags here instead of the Console class because the
' ControlBreakHandler callback routine is not called in a threadsafe
' manor. We don't want to be crossing threads when calling into a
' COM object and potentially corrupt memory causing a crash.
Private mBreak      As Boolean
Private mBreakType  As ConsoleBreakType



Public Property Get Break() As Boolean
    Break = mBreak
End Property

Public Property Let Break(ByVal RHS As Boolean)
    mBreak = RHS
End Property

Public Property Get BreakType() As ConsoleBreakType
    BreakType = mBreakType
End Property

''
' This is the callback used by the SetConsoleCtrlHandler API.
' This function is not called in a threadsafe manor, so don't
' use any COM objects during the routine.
Public Function ControlBreakHandler(ByVal dwCtrlType As Long) As Long
    mBreak = True
    mBreakType = dwCtrlType
    ControlBreakHandler = True
End Function
