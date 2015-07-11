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
'
Public Property Get EmptyGuid() As Guid
    Set EmptyGuid = mEmptyGuid
End Property

Public Property Get IUnknownGuid() As Guid
    Set IUnknownGuid = mIUnknownGuid
End Property

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

Public Function TryParseExact(ByRef s As String, ByRef Format As String, ByRef Result As Guid) As Boolean
    Dim FormatCh As Integer
    
    If Len(Format) = 1 Then
        FormatCh = AscW(Format)
    End If
    
    TryParseExact = TryParseCore(s, FormatCh, Result)
End Function

Public Function TryParse(ByRef s As String, ByRef Result As Guid) As Boolean
    Dim Format As Integer
    
    Select Case Len(s)
        Case 38
            If AscW(s) = corOpenBraceChar Then
                Format = corUpperBChar
            Else
                Format = corUpperPChar
            End If
        Case 36
            Format = corUpperDChar
        Case 32
            Format = corUpperNChar
    End Select
    
    TryParse = TryParseCore(s, Format, Result)
End Function

''
' Parses a string representation of a guid, returning a <b>Guid</b>
' containing the parsed value.
'
' @param s The guid string to be parsed.
' @param LockGuid Locks the guids handle.
' @return A <b>Guid</b> set to the parsed value.
'
Public Function Parse(ByRef s As String) As Guid
    If Not TryParse(s, Parse) Then
        mParser.ThrowError
    End If
End Function

Public Function ParseExact(ByRef s As String, ByRef Format As String) As Guid
    If Not TryParseExact(s, Format, ParseExact) Then
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
'
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
'
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
'
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TryParseCore(ByRef s As String, ByVal Format As Integer, ByRef Result As Guid) As Boolean
    Dim g As VBGUID
    
    If mParser.TryParseExact(s, Format, g) Then
        Set Result = New Guid
        Result.InitReadOnly g
        TryParseCore = True
    Else
        Set Result = mEmptyGuid
    End If
End Function


