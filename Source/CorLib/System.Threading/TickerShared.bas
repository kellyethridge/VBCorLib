Attribute VB_Name = "TickerShared"
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: TickerShared
'

''
' Ticker class helper methods.
'
'@Folder("CorLib.System.Threading")
Option Explicit

Private Const WM_TIMER As Long = &H113

Private mTickers As New Hashtable


Public Function StartTicker(ByVal Source As Ticker) As Long
    Dim NewId As Long
    
    NewId = SetTimer(vbNullPtr, vbNullPtr, Source.Interval, AddressOf TickerCallback)
    
    If NewId = 0 Then _
        Error.Win32Error Err.LastDllError
    
    mTickers(NewId) = ObjPtr(CUnk(Source))
    StartTicker = NewId
End Function

Public Sub StopTicker(ByVal TimerId As Long)
    If mTickers.ContainsKey(TimerId) Then
        mTickers.Remove TimerId
        KillTimer vbNullPtr, TimerId
    End If
End Sub

''
' Callback procedure used by the SetTimer method.
'
Private Sub TickerCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    If uMsg = WM_TIMER Then
        Dim ObjectPointer As Variant
        ObjectPointer = mTickers(idEvent)
        
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
            
            Ticker.OnElapsed
        Else
            ' If we get here, then a timer is still running
            ' but we aren't tracking it, so kill it now.
            KillTimer vbNullPtr, idEvent
        End If
    End If
End Sub

