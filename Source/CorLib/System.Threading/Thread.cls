VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: Thread
'

''
' Creates and controls a thread, sets its priority, and gets its status.
'
Option Explicit

Private mCurrentCulture As CultureInfo


''
' Gets the culture for the current thread.
'
' @return An object that represents the culture for the current thread.
' @remarks The CultureInfo object that is returned by this property, together with its associated objects,
' determine the default format for dates, times, numbers, and currency values.
Public Property Get CurrentCulture() As CultureInfo
    If mCurrentCulture Is Nothing Then
        Set mCurrentCulture = Cor.NewCultureInfo(GetSystemDefaultLCID)
        mCurrentCulture.MarkReadOnly
    End If

    Set CurrentCulture = mCurrentCulture
End Property

''
' Sets the culture for the current thread.
'
' @param Value An object that represents the culture for the current thread.
' @remarks The CultureInfo object that is returned by this property, together with its associated objects,
' determine the default format for dates, times, numbers, and currency values.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>The property is set to <b>Nothing</b>.</td>
'   </tr>
' </table>
Public Property Set CurrentCulture(ByVal Value As CultureInfo)
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
        
    Set mCurrentCulture = Value
End Property
