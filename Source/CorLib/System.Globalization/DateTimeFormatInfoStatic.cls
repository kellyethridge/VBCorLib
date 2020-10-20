VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeFormatInfoStatic"
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
' Module: DateTimeFormatInfoStatic
'

''
' Provides static methods to retrieve various DateTimeFormatInfo objects.
'
' @remarks This class cannot be directly instantiated. To access the methods
' use the variable directly.
' <pre>
' Dim Culture As CultureInfo
' Set Culture = CultureInfo.CurrentInfo
' </pre>
'
' @see DateTimeFormatInfo
'
'@Folder("CorLib.System.Globalization")
Option Explicit


''
' Returns the DateTimeFormatInfo object from the Invariant culture type.
'
' @return A DateTimeFormatInfo object that is a general case. This returns
' the same formatting set of information as using New DateTimeFormatInfo.
'
Public Property Get InvariantInfo() As DateTimeFormatInfo
    Set InvariantInfo = CultureInfo.InvariantCulture.DateTimeFormat
End Property

''
' Returns the DateTimeFormatInfo object associated with the current CultureInfo.
'
' @return The DateTimeFormatInfo for the current CultureInfo of this thread.
'
Public Property Get CurrentInfo() As DateTimeFormatInfo
    Set CurrentInfo = Thread.CurrentThread.CurrentCulture.DateTimeFormat
End Property

''
' Returns a ReadOnly version of a DateTimeFormatInfo object.
'
' @param info The DateTimeFormatInfo object to create a ReadOnly version of.
' @return A ReadOnly version of the supplied DateTimeFormatInfo object.
' @remarks If the supplied DateTimeFormatInfo object is already ReadOnly, then
' that object is returned.
'
Public Function ReadOnly(ByVal Info As DateTimeFormatInfo) As DateTimeFormatInfo
    If Info Is Nothing Then _
        Throw New ArgumentNullException
    
    If Info.IsReadOnly Then
        Set ReadOnly = Info
    Else
        Set ReadOnly = Info.Clone
        ReadOnly.MarkReadOnly
    End If
End Function

''
' Returns the DateTimeFormatInfo object associated with the supplied IFormatProvider.
'
' @param FormatProvider The provider of the DateTimeFormatInfo object.
' @return The DateTimeFormatInfo object from the provider.
' @remarks If the provider does not have a DateTimeFormatInfo object, then Nothing is returned.
'
Public Function GetInstance(ByVal FormatProvider As IFormatProvider) As DateTimeFormatInfo
    If FormatProvider Is Nothing Then
        Set GetInstance = CurrentInfo
    Else
        Set GetInstance = FormatProvider.GetFormat(vbDateTimeFormatInfo)
        
        If GetInstance Is Nothing Then
            Set GetInstance = CurrentInfo
        End If
    End If
End Function

