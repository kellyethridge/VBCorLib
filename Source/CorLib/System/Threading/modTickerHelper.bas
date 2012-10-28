Attribute VB_Name = "modTickerHelper"
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
'    Module: modTickerHelper
'

''
' Ticker class helper methods.
'
Option Explicit

Private Const WM_TIMER As Long = &H113

''
' Holds object pointers to all the running Ticker objects.
' This is so the RaiseElapsed function can be called on the
' Ticker object associated with the idEvent.
' The TimerID (idEvent) is used as the key.
Public Tickers As New Hashtable

''
' Callback procedure used by the SetTimer method.
'
Public Sub TickerCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    If uMsg = WM_TIMER Then
        Dim ObjectPointer As Variant
        ObjectPointer = Tickers(idEvent)
        
        If Not IsEmpty(ObjectPointer) Then
            ' We do the weak reference this way so that
            ' an error can be raised in the event and we
            ' don't need to catch it here to properly unhook
            ' the Ticker object. If we did do error trapping
            ' here, then we would interfere with the error
            ' being raised during the event and not let it
            ' pass back to the application.
            
            Dim Unk As IUnknown
            ObjectPtr(Unk) = ObjectPointer
            
            Dim Ticker As Ticker
            ' Create a strong refernce with a reference count.
            Set Ticker = Unk
            
            ' Unhook the weak reference so errors won't cause
            ' it to be set to Nothing and attempt to decrement the ref count.
            ObjectPtr(Unk) = vbNullPtr
            
            Call Ticker.RaiseElapsed
        Else
            ' If we get here, then a timer is still running
            ' but we aren't tracking it, so kill it now.
            Call KillTimer(vbNullPtr, idEvent)
        End If
    End If
End Sub

