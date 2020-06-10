VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinResourceReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: WinResourceReader
'

''
' A class that reads resources out of EXE and DLL files.
'
' @remarks Use this class to iterate the resources within an EXE or DLL file.
' <p>The only resource types recognized are Strings, Bitmaps, Icons and Cursors.
' All other resource types are maintained as byte arrays.</p>
' <p>Pass this object into the <b>ResourceManager</b> constructor to retrieve
' specific resources based on ID, Type, and Culture. Or, pass this object into
' a <b>ResourceSet</b> to retrieve resources for a specific culture.</p>
'
' @see Constructors
' @see ResourceManager
' @see ResourceSet
' @see ResourceReader
' @see IResourceReader
' @see Win32Resource
'
Option Explicit
Implements IObject
Implements IEnumerable
Implements IResourceReader


Private mEntries    As ArrayList
Private mDecoders   As Hashtable
Private mFileHandle As Long


''
' Adds additional resource decoders to interpret binary
' data that is unknown to the ResourceReader.
'
' @param DecoderToAdd A resource decoder associated with a specific resource type.
' @remarks This allows the ResourceReader to be extended to handle
' additional resource types. If no decoder exists for the specific
' resource type, then the raw data is returned in a Byte array.
' <p>Resource types can be identified by either a numeric value, or
' a string name.</p>
'
Public Sub AddDecoder(ByVal DecoderToAdd As IResourceDecoder)
    If DecoderToAdd Is Nothing Then _
        Throw Cor.NewArgumentNullException("DecoderToAdd")
    
    ' This will replace any existing decoder of the same decode type.
    Dim DecodeType As Variant
    DecodeType = DecoderToAdd.GetDecodeType
    Select Case VarType(DecodeType)
        Case vbString
            Set mDecoders(DecodeType) = DecoderToAdd
        Case vbLong, vbInteger, vbByte
            Set mDecoders(CLng(DecodeType)) = DecoderToAdd
        Case Else
            Throw Cor.NewArgumentException("Invalid Decode Type.", "DecoderToAdd")
    End Select
End Sub

''
' Closes the reader
'
Public Sub CloseReader()
    Set mEntries = Nothing
    FreeLibrary mFileHandle
    mFileHandle = NULL_HANDLE
End Sub

''
' Returns an <a href="IDictionaryEnumerator.html">IDictionaryEnumerator</a> object.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The <a href="DictionaryEntry.Get.Value.html">Value</a> property in the <a href="DictionaryEntry.html">DictionaryEntry</a> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the EXE or DLL file.
' The <a href="DictionaryEntry.Get.Key.html">Key</a> property returns the ID for the specific resource.
' The keys may not be unique across all resources, os using the Key as a unique identifier is not recommended.
Public Function GetEnumerator() As IDictionaryEnumerator
    VerifyReader
    LoadResources
    
    Dim Ret As New ResourceEnumerator
    Ret.Init Me, mEntries
    Set GetEnumerator = Ret
End Function

''
' Returns an <a href="IDictionaryEnumerator.html">IDictionaryEnumerator</a> object.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The <a href="DictionaryEntry.Get.Value.html">Value</a> property in the <a href="DictionaryEntry.html">DictionaryEntry</a> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the EXE or DLL file.
' The <a href="DictionaryEntry.Get.Key.html">Key</a> property returns the ID for the specific resource.
' The keys may not be unique across all resources, os using the Key
' as a unique identifier is not recommended.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mDecoders = Cor.NewHashtable(EqualityComparer:=StringComparer.OrdinalIgnoreCase)
    
    ' Add our known set of decoders.
    AddDecoder New StringResourceDecoder
    AddDecoder New BitMapResourceDecoder
    AddDecoder New IconResourceDecoder
    AddDecoder New IconResourceGroupDecoder
    AddDecoder New CursorResourceDecoder
    AddDecoder New CursorResourceGroupDecoder
End Sub

Friend Sub Init(ByRef FileName As String)
    If Not File.Exists(FileName) Then _
        Throw Cor.NewFileNotFoundException(FileName)
        
    mFileHandle = LoadLibraryA(FileName)
    If mFileHandle = NULL_HANDLE Then _
        Throw Cor.NewIOException("Invalid Application file (" & FileName & ").")
End Sub

Friend Sub InitFromHandle(ByVal FileHandle As Long)
    mFileHandle = FileHandle
End Sub

Friend Sub AddResource(ByVal Key As ResourceKey, ByRef Bytes() As Byte)
    If mDecoders.Contains(Key.ResourceType) Then
        ' If someone can interpret the byte data, let them.
        Dim Decoder As IResourceDecoder
        Set Decoder = mDecoders(Key.ResourceType)
        
        Decoder.Decode Key, Bytes
        
        ' And get all the resources from this set of bytes.
        Dim ReturnValue As Variant
        Do While Decoder.GetResource(Key, ReturnValue)
            mEntries.Add Cor.NewDictionaryEntry(Key, ReturnValue)
        Loop
    Else
        ' Couldn't decode the bytes, so just store them.
        mEntries.Add Cor.NewDictionaryEntry(Key, Bytes)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyReader()
    If mFileHandle = NULL_HANDLE Then
        Throw Cor.NewInvalidOperationException("The Reader is closed and cannot be accessed.")
    End If
End Sub

Private Sub LoadResources()
    If mEntries Is Nothing Then
        Set mEntries = New ArrayList
        
        If EnumResourceTypes(mFileHandle, AddressOf EnumResTypeProc, VarPtr(Me)) = BOOL_FALSE Then
            CloseReader
            
            Dim Ex As Exception
            If Catch(Ex) Then
                Throw Ex
            End If
        End If
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceReader
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceReader_CloseReader()
    CloseReader
End Sub

Private Function IResourceReader_GetEnumerator() As IDictionaryEnumerator
    Set IResourceReader_GetEnumerator = GetEnumerator
End Function

Private Function IResourceReader_NewEnum() As stdole.IUnknown
    Set IResourceReader_NewEnum = NewEnum
End Function
