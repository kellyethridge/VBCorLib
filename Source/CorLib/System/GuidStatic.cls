VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuidStatic"
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
' Module: GuidStatic
'

''
' Provides static methods used to create and manipulate <b>Guid</b> objects.
'
' @remarks The name Guid conflicts with a hidden type in VB. In order to
' specific the VBCorLib version of Guid, a variable must be defined as
' VBCorLib.Guid instead of just Guid.
' </p>This class cannot be directly created. To access the methods, use
' the variable name directly.
' <pre>
' Dim g As VBCorLib.Guid
' Set g = Guid.Parse("{533217B3-CDEC-40A2-B01C-1EA8593B850F}")
' </pre>
' </p>
'
' @see Guid
'
Option Explicit

Private mParser         As New GuidParser
Private mEmptyGuid      As Guid
Private mIUnknownGuid   As Guid
Private mIDispatchGuid  As Guid


''
' Returns a guid that is guaranteed to be all zeros.
'
' @return A zero value guid.
Public Property Get EmptyGuid() As Guid
    Set EmptyGuid = mEmptyGuid
End Property

''
' A constant <a href="Guid.html">Guid</a> representing the <b>IUnknown</b> interface UUID.
'
' @return The <b>IUnknown</b> interface guid.
' @remarks The <b>IUnknown</b> UUID is {00000000-0000-0000-C000-000000000046}.
Public Property Get IUnknownGuid() As Guid
    Set IUnknownGuid = mIUnknownGuid
End Property

''
' A constant <a href="Guid.html">Guid</a> representing the <b>IDispatch</b> interface UUID.
'
' @return The <b>IDispatch</b> interface guid.
' @remarks the <b>IDispatch</b> UUID is {00020400-0000-0000-C000-000000000046}.
Public Property Get IDispatchGuid() As Guid
    Set IDispatchGuid = mIDispatchGuid
End Property

''
' Returns a new <b>Guid</b> object with a random guid generated.
'
' @return A new <b>Guid</b> object.
' @see Guid
'
Public Function NewGuid() As Guid
    Set NewGuid = New Guid
    NewGuid.InitReadOnly CoCreateGuid
End Function

''
' Converts the string representation of a GUID to the equivalent <a href="Guid.html">Guid</a> object, provided that the string is in the specified format.
'
' @param InputString The GUID to convert.
' @param Format One of the following specifiers that indicates the exact format to use when interpreting <i>StringInput</i>: "N", "D", "B", "P", or "X".
' @param Result The structure that will contain the parsed value. If the method returns <b>True</b>, <i>Result</i> contains a valid <a href="Guid.html">Guid</a>. If the method returns <b>False</b>, result equals <a href="GuidStatic.get.Empty.html">Guid.Empty</a>.
' @return <b>True</b> if the parse operation was successful; otherwise, <b>False</b>.
Public Function TryParseExact(ByRef InputString As String, ByRef Format As String, ByRef Result As Guid) As Boolean
    Dim Output      As VBGUID
    Dim FormatCh    As Integer
    
    If Len(Format) = 1 Then
        FormatCh = AscW(Format)
    End If
        
    If mParser.TryParseExact(InputString, FormatCh, Output) Then
        Set Result = New Guid
        Result.InitReadOnly Output
        TryParseExact = True
    Else
        Set Result = mEmptyGuid
    End If
End Function

''
' Converts the string representation of a GUID to the equivalent <a href="Guid.html">Guid</a> object.
'
' @param InputString The GUID to convert.
' @param Result The structure that will contain the parsed value. If the method returns <b>True</b>, <i>Result</i> contains a valid <a href="Guid.html">Guid</a>. If the method returns <b>False</b>, result equals <a href="GuidStatic.get.EmptyGuid.html">Guid.EmptyGuid</a>.
' @return <b>True</b> if the parse operation was successful; otherwise, <b>False</b>.
Public Function TryParse(ByRef InputString As String, ByRef Result As Guid) As Boolean
    Dim Output As VBGUID
    
    If mParser.TryParse(InputString, Output) Then
        Set Result = New Guid
        Result.InitReadOnly Output
        TryParse = True
    Else
        Set Result = mEmptyGuid
    End If
End Function

''
' Converts the string representation of a GUID to the equivalent <a href="Guid.html">Guid</a> object.
'
' @param InputString The string to convert.
' @return A <a href="Guid.html">Guid</a> object of the parsed value.
Public Function Parse(ByRef InputString As String) As Guid
    If Not TryParse(InputString, Parse) Then
        mParser.ThrowError
    End If
End Function

''
' Converts the string representation of a GUID to the equivalent <a href="Guid.html">Guid</a> object.
'
' @param InputString The string to convert.
' @param Format One of the following specifiers that indicates the exact format to use when interpreting <i>StringInput</i>: "N", "D", "B", "P", or "X".
' @return A <a href="Guid.html">Guid</a> object of the parsed value.
Public Function ParseExact(ByRef InputString As String, ByRef Format As String) As Guid
    If Not TryParseExact(InputString, Format, ParseExact) Then
        mParser.ThrowError
    End If
End Function

''
' Creates a guid using specified values and byte array.
'
' @param a A value representing 12345678-xxxx-xxxx-xxxx-xxxxxxxxxxxx
' @param b A value representing xxxxxxxx-1234-xxxx-xxxx-xxxxxxxxxxxx
' @param c A value representing xxxxxxxx-xxxx-1234-xxxx-xxxxxxxxxxxx
' @param d An array representing xxxxxxxx-xxxx-xxxx-1234-123456789012
' @return A new guid initialized to the values and byte array specified.
Public Function FromParts(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByRef d() As Byte) As Guid
    Dim g As VBGUID
    g.Data1 = a
    g.Data2 = b
    g.Data3 = c
    
    If CorArray.IsNull(d) Then _
        Throw Cor.NewArgumentNullException("d", Environment.GetResourceString(ArgumentNull_Array))
    If CorArray.Length(d) <> 8 Then _
        Throw Cor.NewArgumentException("Array must be 8 bytes in length.", "Bytes")
    
    CopyMemory g.Data4(0), d(LBound(d)), 8
    
    Set FromParts = New Guid
    FromParts.InitReadOnly g
End Function

''
' Creates a new guid from the individual values and bytes.
'
' @param a A Long
' @param b An Integer
' @param c An Integer
' @param d A Byte
' @param e A Byte
' @param f A Byte
' @param g A Byte
' @param h A Byte
' @param i A Byte
' @param j A Byte
' @param k A Byte
' @return A new guid initialized to the values and bytes specified.
Public Function FromValues(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByVal d As Byte, ByVal e As Byte, ByVal f As Byte, ByVal g As Byte, ByVal h As Byte, ByVal i As Byte, ByVal j As Byte, ByVal k As Byte) As Guid
    Dim x As VBGUID
    x.Data1 = a
    x.Data2 = b
    x.Data3 = c
    x.Data4(0) = d
    x.Data4(1) = e
    x.Data4(2) = f
    x.Data4(3) = g
    x.Data4(4) = h
    x.Data4(5) = i
    x.Data4(6) = j
    x.Data4(7) = k
    
    Set FromValues = New Guid
    FromValues.InitReadOnly x
End Function

''
' Returns a Read-Only wrapper for the supplied Guid object.
'
' @param g The Guid object to wrap as Read-Only.
' @return A Read-Only Guid object.
' @remarks This is to allow the passing of a Guid object around an
' application without allowing the guid to be modified through the
' Handle property, which allows direct access to the guid structure.
Public Function ReadOnly(ByVal g As Guid) As Guid
    If g Is Nothing Then _
        Error.ArgumentNull "g", ArgumentNull_Generic
    
    If g.IsReadOnly Then
        Set ReadOnly = g
    Else
        Set ReadOnly = g.CreateReadOnly
    End If
End Function

Friend Property Get Parser() As GuidParser
    Set Parser = mParser
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mEmptyGuid = ReadOnly(New Guid)
    Set mIUnknownGuid = ReadOnly(Parse("{00000000-0000-0000-C000-000000000046}"))
    Set mIDispatchGuid = ReadOnly(Parse("{00020400-0000-0000-C000-000000000046}"))
End Sub


