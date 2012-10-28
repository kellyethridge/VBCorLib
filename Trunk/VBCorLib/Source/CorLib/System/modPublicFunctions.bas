Attribute VB_Name = "modPublicFunctions"
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
'    Module: modPublicFunctions
'

''
'   Provides access to the static classes and some shared functions.
'
Option Explicit

Public Powers(31)       As Long
Public MissingVariant   As Variant
Public PowersOf2()      As Integer

''
' Initializes any values for this module.
'
Public Sub InitPublicFunctions()
    InitPowers
    InitPowersOf2
    SetMissingVariant
End Sub

Private Sub SetMissingVariant(Optional ByVal Missing As Variant)
    MissingVariant = Missing
End Sub

''
' Helper function to retrieve the return value of the AddressOf method.
'
' @param pfn Value supplied using AddressOf.
' @return The value returned from AddressOf.
' @remarks The only way to retrieve the value returned from a call to
' AddressOf is to use the AddressOf function when supplying parameter
' values to a function call. By calling this function and using the
' AddressOf method to supply the parameter, the address of the function
' can be obtained.
'
' <h4>Example</h4>
' <pre>
' pfn = FuncAddr(AddressOf MyFunction)
' </pre>
'
Public Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

''
' Returns the variant value as an object.
'
' @param Value The variant containing an object reference.
' @return The object in the variant.
' @remarks This function is a helper function for dealing with
' variant array elements that were originally ParamArray array
' elements. Some methods switch the array pointer with a ParamArray
' and a local variant array. This is the only way to pass a variant
' array into other functions as a Variant() datatype. However, the
' values in the array are all ByRef, where a normal variant array
' can only contain values as ByVal within the array. Since VB doesn't
' expect normal arrays to contain ByRef values, the code generated
' cannot handle ByRef variants that contain objects correctly. So,
' by passing the specific array element into this function, the code
' generated here knows how to deal with a variant that has a ByRef
' value in it, because no matter what is passed in, it will be a
' ByRef value because of the declare type.
'
Public Function CObj(ByRef value As Variant) As Object
    Set CObj = value
End Function

Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function

''
' Modulus method used for large values held within currency datatypes.
'
' @param x The value to be divided.
' @param y The value used to divide.
' @return The remainder of the division.
'
Public Function Modulus(ByVal x As Currency, ByVal y As Currency) As Currency
  Modulus = x - (y * Fix(x / y))
End Function

''
' returns an integer value from the system locale settings.
'
' @param LCID The locale identifier.
' @param LCTYPE The type of value to retrieve from the system.
' @return The value retrieved from the system for the specified locale.
'
Public Function GetLocaleLong(ByVal LCID As Long, ByVal LCType As Long) As Long
    GetLocaleLong = GetLocaleString(LCID, LCType)
End Function

''
' returns a string value from the system locale settings.
'
' @param LCID The locale identifier.
' @param LCTYPE The type of value to retrieve from the system.
' @return The value retrieved from the system for the specified locale.
'
Public Function GetLocaleString(ByVal LCID As Long, ByVal LCType As Long) As String
    Dim Buf As String
    Dim size As Long
    Dim er As Long
    
    size = 128
    Do
        Buf = String$(size, vbNullChar)
        size = API.GetLocaleInfo(LCID, LCType, Buf, size)
        If size > 0 Then Exit Do
        er = Err.LastDllError
        If er <> ERROR_INSUFFICIENT_BUFFER Then IOError er
        size = API.GetLocaleInfo(LCID, LCType, vbNullString, 0)
    Loop
    
    GetLocaleString = Left$(Buf, size - 1)
End Function

''
' Verifies that the FileAccess flags are within a valid range of values.
'
' @param Access The flags to verify.
'
Public Sub VerifyFileAccess(ByVal Access As FileAccess)
    Select Case Access
        Case FileAccess.ReadAccess, FileAccess.ReadWriteAccess, FileAccess.WriteAccess
        Case Else
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Enum), "Access", Access)
    End Select
End Sub

''
' Verifies that the FileShare flags are within a valid range of values.
'
' @param Share The flags to verify.
'
Public Sub VerifyFileShare(ByVal Share As FileShare)
    Select Case Share
        Case FileShare.None, FileShare.ReadShare, FileShare.ReadWriteShare, FileShare.WriteShare
        Case Else
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Enum), "Share", Share)
    End Select
End Sub

''
' Attempts to create a Stream object based on the source.
'
' vbString:     Attempts to open a FileStream.
' vbByte Array: Attempts to create a MemoryStream.
' vbLong:       Attempts to open a FileStream from a file handle.
' vbObject:     Attempts to convert the object to a Stream object.
'
Public Function GetStream(ByRef Source As Variant, ByVal Mode As FileMode, Optional ByVal Access As FileAccess = -1, Optional ByVal Share As FileShare = ReadShare) As Stream
    Select Case VarType(Source)
        Case vbString
            ' We have a filename.
            Set GetStream = Cor.NewFileStream(Source, Mode, Access, Share)
            
        Case vbLong, vbInteger, vbByte
            ' We have a handle.
            Set GetStream = Cor.NewFileStreamFromHandle(Source, Access)
            
        Case vbByteArray
            Dim bytes() As Byte
            SAPtr(bytes) = GetArrayPointer(Source, True)
            Set GetStream = Cor.NewMemoryStream(bytes, Writable:=False)
            SAPtr(bytes) = 0
            
        Case vbObject, vbDataObject
            If Source Is Nothing Then _
                Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Stream))
            If Not TypeOf Source Is Stream Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_StreamRequired), "Source")
            
            Set GetStream = Source
        
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_StreamRequired), "Source")
    End Select
End Function

''
' Attempts to return an LCID from the specified source.
'
' CultureInfo:      Returns the LCID.
' vbLong:           Returns the value.
' vbString:         Assumes culture name, loads culture, returning LCID.
'
Public Function GetLanguageID(ByRef CultureID As Variant) As Long
    Dim Info As CultureInfo
    
    If IsMissing(CultureID) Then
        GetLanguageID = CultureInfo.CurrentCulture.LCID
    Else
        Select Case VarType(CultureID)
            Case vbObject
                If TypeOf CultureID Is CultureInfo Then
                    Set Info = CultureID
                    GetLanguageID = Info.LCID
                Else
                    Throw Cor.NewArgumentException("CultureInfo object required.", "CultureID")
                End If
            
            Case vbLong, vbInteger, vbByte
                GetLanguageID = CultureID
            
            Case vbString
                Set Info = Cor.NewCultureInfo(CultureID)
                GetLanguageID = Info.LCID
                
            Case Else
                Throw Cor.NewArgumentException("CultureInfo object, Name or Language ID required.")
        End Select
    End If
End Function

''
' Returns if the value is an integer value datatype.
'
' @param Value The value to determine if is an integer datatype.
' @return Returns True if the value is an integer datatype, False otherwise.
'
Public Function IsInteger(ByRef value As Variant) As Boolean
    Select Case VarType(value)
        Case vbLong, vbInteger, vbByte: IsInteger = True
    End Select
End Function

Public Function SwapEndian(ByVal value As Long) As Long
    SwapEndian = (((value And &HFF000000) \ &H1000000) And &HFF&) Or _
                 ((value And &HFF0000) \ &H100&) Or _
                 ((value And &HFF00&) * &H100&) Or _
                 ((value And &H7F&) * &H1000000)
    If (value And &H80&) Then SwapEndian = SwapEndian Or &H80000000
End Function

Public Function RRotate(ByVal value As Long, ByVal count As Long) As Long
    RRotate = Helper.ShiftRight(value, count) Or Helper.ShiftLeft(value, 32 - count)
End Function

Public Function LRotate(ByVal value As Long, ByVal count As Long) As Long
    LRotate = Helper.ShiftLeft(value, count) Or Helper.ShiftRight(value, 32 - count)
End Function

Public Function ReverseByteCopy(ByRef bytes() As Byte) As Byte()
    Dim ub As Long
    ub = UBound(bytes)
    
    Dim ret() As Byte
    ReDim ret(0 To ub)
    
    Dim i As Long
    For i = 0 To ub
        ret(i) = bytes(ub - i)
    Next i
    
    ReverseByteCopy = ret
End Function


''
' Initializes an array for quick powers of 2 lookup.
'
Private Sub InitPowers()
    Dim i As Long
    For i = 0 To 30
        Powers(i) = 2 ^ i
    Next i
    Powers(31) = &H80000000
End Sub

Private Sub InitPowersOf2()
    ReDim PowersOf2(0 To 15)
    Dim i As Long
    For i = 0 To 14
        PowersOf2(i) = 2 ^ i
    Next i
    
    PowersOf2(15) = &H8000
End Sub

