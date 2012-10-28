VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatchStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: StopWatchStatic
'

''
' Provides methods that are related to the StopWatch class.
'
' @see StopWatch
'
Option Explicit

Private mIsHighResolution   As Boolean
Private mFrequency          As Currency



''
' Returns if the StopWatch class will be using the high resolution timer.
'
' @return Returns True if there is high resolution, False otherwise.
' @remarks High resolution timers are hardward dependant.
'
Public Property Get IsHighResolution() As Boolean
    IsHighResolution = mIsHighResolution
End Property

''
' Returns the frequency of the high resolution timer.
'
' @return The frequency used by the high resolution timer.
'
Public Property Get Frequency() As Currency
    Frequency = mFrequency
End Property

''
' Returns the current time of the timer.
'
' @return The current time from the timer.
' @remarks If the timer is high resolution, then the result from
' QueryPerformanceCounter is returned, otherwise, the result from
' timeGetTime is used.
'
Public Function GetTimeStamp() As Currency
    GetTimeStamp = InternalGetTimeStamp
End Function

''
' Returns a new StopWatch object that is then started upon return.
'
' @return A new StopWatch object.
'
Public Function StartNew() As StopWatch
    Set StartNew = New StopWatch
    Call StartNew.StartCount
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function InternalGetTimeStamp() As Currency
    If mIsHighResolution Then
        Call QueryPerformanceCounter(InternalGetTimeStamp)
    Else
        InternalGetTimeStamp = timeGetTime
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mIsHighResolution = (QueryPerformanceFrequency(mFrequency) <> BOOL_FALSE)
End Sub
