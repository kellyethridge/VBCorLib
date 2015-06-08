VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: Ticker
'

''
' Provides a recurring timer event.
'
' @remarks To create new <b>Ticker</b> using the constructor, call the <b>NewTicker</b> method.
' <pre>
' Private WithEvents mTicker As Ticker
'
' Private Sub Class_Initialize()
'     Set mTicker = NewTicker(2000)
'     mTicker.StartTicker
' End Sub
'
' Private Sub mTicker_Elapsed()
'     '' occurs every 2000 milliseconds (2 seconds.)
' End Sub
' </pre>
' The Ticker can also use a callback function when the time has elapsed instead
' of only raising an event. The event is still raised, however, a callback function
' is called as well.
' <pre>
' Private mTicker As Ticker  '' declared without using WithEvents
'
' Private Sub Main()
'     Set mTicker = NewTicker(2000,,, AddressOf TimerCallback)
'     mTicker.StartTicker
'
'     '' Application Code
' End Sub
'
' Private Sub TickerCallback(ByRef Ticker As Ticker, ByRef Data As Variant)
'     '' occurs every 2000 milliseconds (2 seconds.)
' End Sub
' </pre>
'
' @see Constructors
'
Option Explicit
Implements IObject

''
' This event is raised when the Ticker interval has been reached.
'
' @param Data A user-defined value associated with the <b>Ticker</b>.
'
Public Event Elapsed(ByRef Data As Variant)

Private mInterval   As Long
Private mAutoReset  As Boolean
Private mTimerId    As Long
Private mData       As Variant
Private mCallback   As Action_T_T


''
' Returns custom data associated with this Ticker instance.
'
' @return Custom data.
' @remarks The data is not used by the Ticker. It is to allow
' a developer to associate values to the Ticker if needed.
'
Public Property Get Data() As Variant
    VariantCopy Data, mData
End Property

''
' Sets custom data associated with this Ticker instance.
'
' @param RHS The data to associate with this Ticker instance.
' @remarks The data is not used by the Ticker. It is to allow
' a developer to associate values to the Ticker if needed.
'
Public Property Let Data(ByRef Value As Variant)
    mData = Value
End Property

''
' Sets custom data associated with this Ticker instance.
'
' @param RHS The data to associate with this Ticker instance.
' @remarks The data is not used by the Ticker. It is to allow
' a developer to associate values to the Ticker if needed.
'
Public Property Set Data(ByRef Value As Variant)
    Set mData = Value
End Property

''
' Returns if the Ticker event will be recurring.
'
' @return AutoReset flag.
' @remarks If AutoReset is set to false, then the Ticker event
' will happen only once. The Ticker will need to be restarted
' for another event to occur.
' <p>The default is True.</p>
'
Public Property Get AutoReset() As Boolean
    AutoReset = mAutoReset
End Property

''
' Sets if the Ticker event will be recurring or not.
'
' @param RHS Flag indicating if the Ticker will be recurring.
' @remarks If AutoReset is set to false, then the Ticker event
' will happen only once. The Ticker will need to be restarted
' for another event to occur.
' <p>The default is True.</p>
'
Public Property Let AutoReset(ByVal Value As Boolean)
    mAutoReset = Value
End Property

''
' Returns the duration between Ticker events in milliseconds.
'
' @return Duration between events in milliseconds
' @remarks If the Ticker is already running, it will be reset
' and begin a new duration with the new interval.
'
Public Property Get Interval() As Long
    Interval = mInterval
End Property

''
' Sets the duration between Ticker events in milliseconds.
'
' @param RHS Duration between events in milliseconds.
' @remarks If the Ticker is already running, it will be reset
' and begin a new duration with the new interval.
'
Public Property Let Interval(ByVal Value As Long)
    If Value < 0 Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_NeedNonNegNum
    
    mInterval = Value
    If Enabled Then
        StartTicker
    End If
End Property

''
' Returns if the Ticker is currently running.
'
' @return If the Ticker is running.
'
Public Property Get Enabled() As Boolean
    Enabled = (mTimerId <> 0)
End Property

''
' Sets the Ticker to be stopped or running.
'
' @param RHS Value indicating if the Ticker is stopped or running.
' @remarks If setting Enabled to True and the Ticker is already
' running, then the Ticker will be stopped and restarted.
'
Public Property Let Enabled(ByVal Value As Boolean)
    If Value Then
        StartTicker
    Else
        StopTicker
    End If
End Property

''
' Starts the Ticker.
'
' @remarks If the Ticker is already running, it will be stopped and restarted.
' A new full interval will be elapsed before the next event occurs.
'
Public Sub StartTicker()
    StopTicker
    
    If mInterval > 0 Then
        mTimerId = modTicker.StartTicker(Me)
    End If
End Sub

''
' Stops the Ticker if it is running.
'
Public Sub StopTicker()
    If Enabled Then
        modTicker.StopTicker mTimerId
    End If
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

Friend Sub OnElapsed()
    If Not mAutoReset Then
        StopTicker
    End If
    
    RaiseEvent Elapsed(mData)
    
    If Not mCallback Is Nothing Then
        mCallback.Invoke VarPtr(Me), VarPtr(mData)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Interval As Long, ByRef Data As Variant, ByVal AutoReset As Long, ByVal AddressOfCallback As Long)
    Me.Interval = Interval
    Me.AutoReset = AutoReset
    VariantCopyInd mData, Data
    
    If AddressOfCallback <> vbNullPtr Then
        Set mCallback = NewDelegate(AddressOfCallback)
    End If
End Sub

Private Sub Class_Initialize()
    mAutoReset = True
End Sub

Private Sub Class_Terminate()
    StopTicker
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
