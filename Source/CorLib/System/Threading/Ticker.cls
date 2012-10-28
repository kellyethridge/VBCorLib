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
'    Module: Ticker
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
Private mTimerID    As Long
Private mData       As Variant
Private mDelegator  As FunctionDelegator
Private mCallback   As ITwoArgReturnVoid



''
' Returns custom data associated with this Ticker instance.
'
' @return Custom data.
' @remarks The data is not used by the Ticker. It is to allow
' a developer to associate values to the Ticker if needed.
'
Public Property Get Data() As Variant
    Call VariantCopy(Data, mData)
End Property

''
' Sets custom data associated with this Ticker instance.
'
' @param RHS The data to associate with this Ticker instance.
' @remarks The data is not used by the Ticker. It is to allow
' a developer to associate values to the Ticker if needed.
'
Public Property Let Data(ByVal RHS As Variant)
    mData = RHS
End Property

''
' Sets custom data associated with this Ticker instance.
'
' @param RHS The data to associate with this Ticker instance.
' @remarks The data is not used by the Ticker. It is to allow
' a developer to associate values to the Ticker if needed.
'
Public Property Set Data(ByVal RHS As Variant)
    Set mData = RHS
End Property

''
' Returns the callback address of the function used by the Ticker object.
'
' @return Callback address. Setting this to zero will stop the callback.
' @remarks A callback method is used to allow a Ticker object to
' notify the function without having to use an Event. A function callback
' must have the following signature, or the application may crash.
' <pre>
' Public Sub TickerCallback(ByRef Ticker As Ticker, ByRef Data As Variant)
'     '' do stuff
' End Sub
' </pre>
' The two parameters must be declared as ByRef.
'
Public Property Get Callback() As Long
    Callback = mDelegator.pfn
End Property

''
' Sets the callback address of the function to be used by the Ticker object.
'
' @param RHS The address of the callback function. Setting this to zero will stop the callback.
' @remarks A callback method is used to allow a Ticker object to
' notify the function without having to use an Event. A function callback
' must have the following signature, or the application may crash.
' <pre>
' Public Sub TickerCallback(ByRef Ticker As Ticker, ByRef Data As Variant)
'     '' do stuff
' End Sub
' </pre>
' The two parameters must be declared as ByRef.
'
Public Property Let Callback(ByVal RHS As Long)
    mDelegator.pfn = RHS
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
Public Property Let AutoReset(ByVal RHS As Boolean)
    mAutoReset = RHS
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
Public Property Let Interval(ByVal RHS As Long)
    If RHS < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Interval", RHS)
        
    mInterval = RHS
    If Enabled Then Call StartTicker
End Property

''
' Returns if the Ticker is currently running.
'
' @return If the Ticker is running.
'
Public Property Get Enabled() As Boolean
    Enabled = (mTimerID <> 0)
End Property

''
' Sets the Ticker to be stopped or running.
'
' @param RHS Value indicating if the Ticker is stopped or running.
' @remarks If setting Enabled to True and the Ticker is already
' running, then the Ticker will be stopped and restarted.
'
Public Property Let Enabled(ByVal RHS As Boolean)
    If RHS = True Then
        Call StartTicker
    Else
        Call StopTicker
    End If
End Property

''
' Starts the Ticker.
'
' @remarks If the Ticker is already running, it will be stopped and restarted.
' A new full interval will be elapsed before the next event occurs.
'
Public Sub StartTicker()
    Call StopTicker
    
    If mInterval > 0 Then
        mTimerID = SetTimer(vbNullPtr, vbNullPtr, mInterval, AddressOf TickerCallback)
        If mTimerID = 0 Then IOError Err.LastDllError
        
        ' Set a weak callback reference that the timer event can use
        ' to call the RaiseElapsed function from the callback method.
        Tickers(mTimerID) = ObjPtr(CUnk(Me))
    End If
End Sub

''
' Stops the Ticker if it is running.
'
Public Sub StopTicker()
    If Not Enabled Then Exit Sub
    
    Dim t As Long
    t = mTimerID
    mTimerID = 0
    If KillTimer(vbNullPtr, t) = BOOL_FALSE Then IOError Err.LastDllError
    Call Tickers.Remove(t)
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interfacek
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub RaiseElapsed()
    If Not mAutoReset Then StopTicker
    RaiseEvent Elapsed(mData)
    If mDelegator.pfn <> vbNullPtr Then Call mCallback.Call(VarPtr(Me), VarPtr(mData))
End Sub

Friend Sub Init(ByVal Interval As Long, ByVal Data As Variant, ByVal AutoReset As Long, ByVal AddressOfCallback As Long)
    Me.Interval = Interval
    Me.AutoReset = AutoReset
    Call Helper.MoveVariant(mData, Data)
    Me.Callback = AddressOfCallback
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mAutoReset = True
    Set mCallback = InitDelegator(mDelegator)
End Sub

Private Sub Class_Terminate()
    Call StopTicker
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
