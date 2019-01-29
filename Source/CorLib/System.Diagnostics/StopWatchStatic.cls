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
' Module: StopWatchStatic
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
    StartNew.StartCount
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function InternalGetTimeStamp() As Currency
    If mIsHighResolution Then
        QueryPerformanceCounter InternalGetTimeStamp
    Else
        InternalGetTimeStamp = timeGetTime
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mIsHighResolution = (QueryPerformanceFrequency(mFrequency) <> BOOL_FALSE)
End Sub
