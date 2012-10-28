VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Guid"
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
'    Module: Guid
'

''
' Represents a globally unique identifier (GUID).
'
' @remarks A GUID is a 128-bit integer (16 bytes) that can be used across all
' computers and networks wherever a unique identifier is required. Such an identifier
' has a very low probability of being duplicated.
'
' @see GuidStatic
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable


Private Const REGDB_E_CLASSNOTREG   As Long = &H80040154
Private Const REGDB_E_READREGDB     As Long = &H80040150

Private mGuid As VBGUID



''
' Returns a pointer to the internal GUID structure.
'
' @return A long containing the pointer to the structure.
' @remarks This is provided to allow for APIs to access the
' actual guid structure directly, since they expect to be
' accessing a structure, not an object.
' <p>Guid style APIs declare the guid parameter something like
' <pre>ByRef g As VBGUID</pre>
' In order to use the VBCorLib <b>Guid</b> then the API would
' need to be declared something like
' <pre>ByVal ptrG As Long</pre>
' and pass in the Handle to the guid object.</p>
' <p>Exposing the handle is very dangerous in that the object
' is not immutable.</p>
'
Public Property Get Handle() As Long
    Handle = VarPtr(mGuid)
End Property

''
' Returns if the guid is read-only.
'
' @return Returns <b>True</b> if the guid is locked, otherwise <b>False</b>.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns the guid value has an array of bytes.
'
' @return An array of bytes representing the guid.
'
Public Function ToByteArray() As Byte()
    Dim Ret(15) As Byte
    
    Call CopyMemory(Ret(0), mGuid, 16)
    ToByteArray = Ret
End Function

''
' Compares this guid to another.
'
' @param Value The guid to compare to this guid.
' @return A Long
'
Public Function CompareTo(ByRef Value As Variant) As Long
    Dim g As Guid
    
    On Error GoTo errTrap
    Set g = Value
    If g Is Nothing Then
        CompareTo = 1
    Else
        CompareTo = -g.InternalCompare(mGuid)
    End If
    
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("A Guid object is required.", "Value")
End Function

''
' Returns a string representation of guid value.
'
' @param Format The format of the guid value to be returned.
' @param Provider A format provider used to format the guid.
' @return String representing of the guid.
' @remarks There are 4 format types: "D","B","N","P".<br>
' <pre>
' "D" = XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
' "B" = {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
' "P" = (XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX)
' "N" = XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' </pre>
' <b>B</b> is the default format.
'
Public Function ToString(Optional ByVal Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim Buf As String
    Dim Size As Long
    
    Buf = String$(45, 0)
    Size = StringFromGUID2(mGuid, Buf, Len(Buf))
    ToString = Left$(Buf, Size - 1)
    
    Select Case UCase$(Format)
        Case "D"
            ToString = Mid$(ToString, 2, 36)
        Case "N"
            ToString = Replace$(Mid$(ToString, 2, 36), "-", "")
        Case "P"
            Mid$(ToString, 1, 1) = "("
            Mid$(ToString, 38, 1) = ")"
        Case "B", ""
            ' already formatted to this and is the default
        Case Else
            Throw Cor.NewArgumentException("Invalid Format specified.", "Format")
    End Select
End Function

''
' Returns the Program ID associated with the guid in the <Server.Class> format.
'
' @return A string containing the Program ID associated with the guid.
'
Public Function ToProgramID() As String
    Dim lpszProgID As Long
    Select Case ProgIDFromCLSID(mGuid, lpszProgID)
        Case REGDB_E_CLASSNOTREG
            Throw Cor.NewInvalidOperationException("The Program ID is not a registered class.")
            
        Case REGDB_E_READREGDB
            Throw Cor.NewUnauthorizedAccessException("The Registry could not be read.")
    End Select
    
    ToProgramID = SysAllocString(lpszProgID)
    Call CoTaskMemFree(lpszProgID)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        If TypeOf Value Is Guid Then
            Dim g As Guid
            Set g = Value
            Equals = g.InternalEquals(mGuid)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    With mGuid
        GetHashCode = .Data1 Xor AsLong(.Data2) Xor AsLong(.Data4(0)) Xor AsLong(.Data4(4))
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function InternalEquals(ByRef g As VBGUID) As Boolean
    InternalEquals = IsEqualGUID(g, mGuid)
End Function

Friend Function InternalCompare(ByRef g As VBGUID) As Long
    With mGuid
        If Not EqualFields(.Data1, g.Data1, InternalCompare) Then Exit Function
        If Not EqualFields(.Data2, g.Data2, InternalCompare) Then Exit Function
        If Not EqualFields(.Data3, g.Data3, InternalCompare) Then Exit Function
        
        Dim i As Long
        For i = 0 To 7
            If Not EqualFields(.Data4(i), g.Data4(i), InternalCompare) Then Exit Function
        Next i
    End With
End Function

Friend Sub NewGuid()
    mGuid = CoCreateGuid
End Sub

Friend Sub Parse(ByVal s As String)
    
    If Len(s) = 38 Then
        '{C200E360-38C5-11CE-AE62-08002B2B79EF} or
        '(C200E360-38C5-11CE-AE62-08002B2B79EF)
        Select Case Asc(s)
            Case 123:   If Asc(Right$(s, 1)) <> 125 Then Call FormatError
            Case 40:    If Asc(Right$(s, 1)) <> 41 Then Call FormatError
            Case Else:  FormatError
        End Select
        s = Mid$(s, 2, 36)
    End If
    
    Select Case Len(s)
        Case 36
            'C200E360-38C5-11CE-AE62-08002B2B79EF
            If s Like "*[!0-9a-fA-F,-]*" Then Call FormatError
            If Asc(Mid$(s, 9, 1)) <> 45 Then Call FormatError
            If Asc(Mid$(s, 14, 1)) <> 45 Then Call FormatError
            If Asc(Mid$(s, 19, 1)) <> 45 Then Call FormatError
            If Asc(Mid$(s, 24, 1)) <> 45 Then Call FormatError
            s = "{" & s & "}"
        Case 32
            'C200E36038C511CEAE6208002B2B79EF
            If s Like "*[!0-9a-fA-F]*" Then Call FormatError
            s = "{" & Left$(s, 8) & "-" & Mid$(s, 9, 4) & "-" & Mid$(s, 13, 4) & "-" & Mid$(s, 17, 4) & "-" & Mid$(s, 21) & "}"
        
        Case Else
            Call FormatError
    End Select
    
    On Error GoTo errTrap
    mGuid = GUIDFromString(s)
    Exit Sub
    
errTrap:
    FormatError
End Sub

Private Sub FormatError()
    Throw Cor.NewFormatException("Invalid Guid string format.")
End Sub

Friend Sub FromByteArray(ByRef b() As Byte)
    If cArray.IsNull(b) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    If cArray.GetLength(b) <> 16 Then _
        Throw Cor.NewArgumentException("Array must be 16 bytes in length.", "Bytes")
    
    Call CopyMemory(mGuid, b(LBound(b)), 16)
End Sub

Friend Sub FromParts(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByRef d() As Byte)
    With mGuid
        .Data1 = a
        .Data2 = b
        .Data3 = c
        
        If cArray.IsNull(d) Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
        If cArray.GetLength(d) <> 8 Then _
            Throw Cor.NewArgumentException("Array must be 8 bytes in length.", "Bytes")
        
        Call CopyMemory(.Data4(0), d(LBound(d)), 8)
    End With
End Sub

Friend Sub FromProgramID(ByVal ProgID As String)
    On Error GoTo errTrap
    mGuid = CLSIDFromProgID(ProgID)
    Exit Sub
    
errTrap:
    Throw Cor.NewArgumentException(cString.Format("Program ID '{0}' was not found.", ProgID), "ProgID")
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EqualFields(ByVal MyField As Long, ByVal TheirField As Long, ByRef RetVal As Long) As Boolean
    If MyField < TheirField Then
        RetVal = -1
    ElseIf MyField > TheirField Then
        RetVal = 1
    Else
        RetVal = 0
        EqualFields = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mGuid.Data1 = .ReadProperty("Data1")
        mGuid.Data2 = .ReadProperty("Data2")
        mGuid.Data3 = .ReadProperty("Data3")
        
        Dim b() As Byte
        b = .ReadProperty("Data4")
        Call CopyMemory(mGuid.Data4(0), b(0), 8)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Data1", mGuid.Data1)
        Call .WriteProperty("Data2", mGuid.Data2)
        Call .WriteProperty("Data3", mGuid.Data3)
        Call .WriteProperty("Data4", mGuid.Data4)
    End With
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(ByVal Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = ToString(Format, Provider)
End Function
