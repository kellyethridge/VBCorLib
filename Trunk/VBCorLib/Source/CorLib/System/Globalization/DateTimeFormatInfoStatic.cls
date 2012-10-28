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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: DateTimeFormatInfoStatic
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
Option Explicit

Private mStringBuilder As StringBuilder


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
    Set CurrentInfo = CultureInfo.CurrentCulture.DateTimeFormat
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
        ReadOnly.IsReadOnly = True
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
    If Not FormatProvider Is Nothing Then
        Set GetInstance = FormatProvider.GetFormat("DateTimeFormatInfo")
        If GetInstance Is Nothing Then Set GetInstance = CurrentInfo
    Else
        Set GetInstance = CurrentInfo
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get StringBuilder() As StringBuilder
    Set StringBuilder = mStringBuilder
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mStringBuilder = New StringBuilder
End Sub
