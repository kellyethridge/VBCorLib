Attribute VB_Name = "HelperFunctions"
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
' Module: modPublicFunctions
'

''
'   Provides access to the static classes and some shared functions.
'
Option Explicit

Public Powers(31)       As Long
Public PowersOf2()      As Integer

''
' Initializes any values for this module.
'
Public Sub InitPublicFunctions()
    InitPowers
    InitPowersOf2
End Sub

''
' Retrieves the return value of the AddressOf method.
'
' @param pfn Value supplied using AddressOf.
' @return The value returned from AddressOf.
' @remarks This allows us to obtain the result of the VB AddressOf method.
'
' Example:
'   pfn = FuncAddr(AddressOf MyFunction)
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
Public Function CObj(ByRef Value As Variant) As Object
    Set CObj = Value
End Function

Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function

Public Function WeakPtr(ByVal Obj As IUnknown) As Long
    WeakPtr = ObjPtr(Obj)
End Function

Public Function StrongPtr(ByVal Ptr As Long) As IUnknown
    Dim Obj As IUnknown
    ObjectPtr(Obj) = Ptr
    Set StrongPtr = Obj
    ObjectPtr(Obj) = vbNullPtr
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
    Dim Buf         As String
    Dim Size        As Long
    Dim ErrorCode   As Long
    
    Size = 128
    Do
        Buf = String$(Size, vbNullChar)
        Size = GetLocaleInfoW(LCID, LCType, Buf, Size)
        
        If Size > 0 Then
            Exit Do
        End If
        
        ErrorCode = Err.LastDllError
        
        If ErrorCode <> ERROR_INSUFFICIENT_BUFFER Then _
            Error.Win32Error ErrorCode
            
        Size = Api.GetLocaleInfo(LCID, LCType, vbNullString, 0)
    Loop
    
    GetLocaleString = Left$(Buf, Size - 1)
End Function

''
' Attempts to create a Stream object based on the source.
'
' vbString:     Attempts to open a FileStream.
' vbByte Array: Attempts to create a MemoryStream.
' vbObject:     Attempts to convert the object to a Stream object.
'
Public Function GetStream(ByRef Source As Variant, ByVal Mode As FileMode, Optional ByVal Access As FileAccess, Optional ByVal Share As FileShare = FileShare.ReadShare) As Stream
    Select Case VarType(Source)
        Case vbString
            Set GetStream = Cor.NewFileStream(CStr(Source), Mode, Access, Share)
            
        Case vbByteArray
            Dim Bytes() As Byte
            SAPtr(Bytes) = SAPtrV(Source)
            If CorArray.IsNull(Bytes) Then _
                Error.ArgumentNull "Source", ArgumentNull_Array
            Set GetStream = Cor.NewMemoryStream(Bytes, Writable:=False)
            SAPtr(Bytes) = 0
            
        Case vbObject
            If Source Is Nothing Then _
                Error.ArgumentNull "Source", ArgumentNull_Stream
            If TypeOf Source Is Stream Then
                Set GetStream = Source
            ElseIf TypeOf Source Is SafeFileHandle Then
                Set GetStream = Cor.NewFileStreamWithHandle(Source, Access)
            Else
                Error.Argument Argument_StreamRequired
            End If
                
        Case Else
            Error.Argument Argument_StreamRequired
    End Select
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
    Dim i As Long
    ReDim PowersOf2(0 To 15)
    
    For i = 0 To 14
        PowersOf2(i) = 2 ^ i
    Next i
    
    PowersOf2(15) = &H8000
End Sub

