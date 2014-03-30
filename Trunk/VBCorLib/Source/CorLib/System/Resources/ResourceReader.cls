VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceReader"
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
' Module: ResourceReader
'

''
' Provides a means to iterate through the resources in a .RES file.
'
' @remarks The only resource types recognized are Strings, Icons, Cursors, and Bitmaps.
' All other resource types are contained as a byte array.
' <p>Pass this object into the <b>ResourceManager</b> constructor to access specific
' resources based on ID, Type, and Culture. Or pass this object into a <b>ResourceSet</b>
' to retrieve the resources for a specific culture.</p>
'
' @see Constructors
' @see ResourceManager
' @see ResourceSet
' @see IResourceReader
' @see Win32Resource
'
Option Explicit
Implements IObject
Implements IEnumerable
Implements IResourceReader

Private Type ResourceHeader
    DataSize As Long
    HeaderSize As Long
    ResourceType As Variant
    ResourceName As Variant
    DataVersion As Long
    MemoryFlags As Integer
    LanguageID As Integer
    Version As Long
    Characteristics As Long
End Type


Private mReader             As BinaryReader
Private mEntries            As ArrayList
Private mDecoders           As Hashtable
Private mResourceDecoder    As IResourceDecoder
Private mResourceCount      As Long



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
' Returns an IDictionaryEnumerator object.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
' The <i>Key</b> property returns the ID for the specific resource.
' The Keys may not be unique across all resources, os using the Key
' as a unique identifier is not recommended.
'
Public Function GetEnumerator() As IDictionaryEnumerator
    Call VerifyReader
    
    Dim Ret As New ResourceEnumerator
    Call Ret.Init(Me, mEntries)
    Set GetEnumerator = Ret
End Function

''
' Returns a For..Each compatible enumerator.
'
' @return Enumerator object.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
' The <i>Key</b> property returns the ID for the specific resource.
' The Keys may not be unique across all resources, os using the Key
' as a unique identifier is not recommended.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Closes the reader and releases any resources.
'
' @remarks <b>Close</b> can be called safely multiple times.
'
Public Sub CloseReader()
    If Not mReader Is Nothing Then
        mReader.CloseReader
        Set mReader = Nothing
        Set mEntries = Nothing
    End If
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Source As Variant)
    Set mReader = Cor.NewBinaryReader(GetStream(Source, OpenExisting, ReadAccess))
    Call VerifyResourceFile
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyReader()
    If mReader Is Nothing Then Throw Cor.NewInvalidOperationException("The Reader is closed and cannot be accessed.")
    Call LoadResources
End Sub

Private Function GetNextResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
    If mResourceCount > 0 Then
        GetNextResource = mResourceDecoder.GetResource(ReturnKey, ReturnValue)
        mResourceCount = mResourceCount - 1
        Exit Function
    End If
    
    Dim Header As ResourceHeader
    Call GetNextResourceHeader(Header)
    
    Set ReturnKey = Cor.NewResourceKey(Header.ResourceName, Header.ResourceType, Header.LanguageID)
    
    Dim Data() As Byte
    Data = cArray.CreateInstance(vbByte, Header.DataSize)
    Call mReader.Read(Data, 0, Header.DataSize)
    
    If mDecoders.Contains(Header.ResourceType) Then
        ' If someone can interpret the byte data, let them.
        Set mResourceDecoder = mDecoders(Header.ResourceType)
        mResourceCount = mResourceDecoder.Decode(ReturnKey, Data)
        GetNextResource = mResourceDecoder.GetResource(ReturnKey, ReturnValue)
        mResourceCount = mResourceCount - 1
    Else
        ' Otherwise just return the raw byte data.
        ReturnValue = Data
        mResourceCount = 0
    End If
    
    GetNextResource = True
End Function

Private Function HasMoreResources() As Boolean
    If mResourceCount > 0 Then
        HasMoreResources = True
    Else
        With mReader.BaseStream
            HasMoreResources = (.Position < .Length - 3)
        End With
    End If
End Function

Private Sub LoadResources()
    Set mEntries = New ArrayList
    
    Do While HasMoreResources
        Dim Key As ResourceKey
        Dim Value As Variant
        If Not GetNextResource(Key, Value) Then Exit Do
        Call mEntries.Add(Cor.NewDictionaryEntry(Key, Value))
    Loop
End Sub

Private Sub VerifyResourceFile()
    ' A .RES file has atleast one empty 32byte header.
    If mReader.BaseStream.Length < 32 Then Call ResourceError
    
    ' DataSize should be 0.
    If mReader.ReadLong <> 0 Then Call ResourceError
    
    ' HeaderSize should be 32.
    If mReader.ReadLong <> 32 Then Call ResourceError
    
    ' ResourceType should be &H0000FFFF
    If mReader.ReadLong <> &HFFFF& Then Call ResourceError
    
    ' ResourceName should be &H0000FFFF
    If mReader.ReadLong <> &HFFFF& Then Call ResourceError
    
    ' We are pretty sure this is a valid .RES file, so
    ' move to the beginning of the next entry, which
    ' should be the first valid resource entry in the file.
    mReader.BaseStream.Position = 32
End Sub

Private Sub ResourceError()
    Throw Cor.NewIOException("Invalid resource file format.")
End Sub

''
' Resource names and types can be either a name or a number.
' This function determines which type is being used to
' identify the name or type and returns the appropriate value.
'
Private Function GetOrdinalOrName() As Variant
    
    ' Get the first 2 bytes.
    Dim ch As Integer
    ch = mReader.ReadInteger
    
    ' If the first 2 bytes equal &HFFFF then we have a number.
    If ch = -1 Then
        ' So the next 2 bytes represent a 16bit number.
        GetOrdinalOrName = CLng(mReader.ReadInteger)
    Else
        ' The first 2 bytes represent a character,
        ' we need to append characters until we hit a null.
        
        ' We don't want to keep creating a new builder
        ' for every resource in the file, so use Static.
        Static NameBuilder As New StringBuilder
        
        ' And just set the length to reset the builder.
        NameBuilder.Length = 0
        
        ' Keep doing this until we reach a terminating null.
        Do While ch <> 0
            ' Append the first character we retrieved.
            Call NameBuilder.AppendChar(ch)
            
            ' And get the next character.
            ch = mReader.ReadInteger
        Loop
        
        ' Return the name.
        GetOrdinalOrName = NameBuilder.ToString
    End If
End Function

Private Sub DWordAlignStreamPosition()
    Dim CurrentPosition As Long
    
    With mReader.BaseStream
        CurrentPosition = .Position
        If (CurrentPosition Mod 4) <> 0 Then Call .SeekPosition(4 - (CurrentPosition Mod 4), SeekOrigin.FromCurrent)
    End With
End Sub

Private Sub GetNextResourceHeader(ByRef Header As ResourceHeader)
    With Header
        DWordAlignStreamPosition
        .DataSize = mReader.ReadLong
        .HeaderSize = mReader.ReadLong
        
        ' We don't DWord align between these two because the
        ' resource compiler only aligns once both of them are written.
        .ResourceType = GetOrdinalOrName
        .ResourceName = GetOrdinalOrName
        
        ' We need to DWord align the stream position because the
        ' name of the resource may have left us out of alignment.
        Call DWordAlignStreamPosition
        
        .DataVersion = mReader.ReadLong
        .MemoryFlags = mReader.ReadInteger
        .LanguageID = mReader.ReadInteger
        .Version = mReader.ReadLong
        .Characteristics = mReader.ReadLong
    End With
    
    ' The Stream should be positioned on the first byte
    ' of the data for the current resource.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mDecoders = Cor.NewHashtable(hcp:=New CaseInsensitiveHashCodePrvdr, Comparer:=New CaseInsensitiveComparer)
    
    ' Add our known set of decoders.
    Call AddDecoder(New ResStringTableDecoder)
    Call AddDecoder(New ResBitmapDecoder)
    Call AddDecoder(New ResIconDecoder)
    Call AddDecoder(New ResIconGroupDecoder)
    Call AddDecoder(New ResCursorDecoder)
    Call AddDecoder(New ResCursorGroupDecoder)
End Sub

Private Sub Class_Terminate()
    Call CloseReader
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceReader Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceReader_CloseReader()
    Call CloseReader
End Sub

Private Function IResourceReader_GetEnumerator() As IDictionaryEnumerator
    Set IResourceReader_GetEnumerator = GetEnumerator
End Function

Private Function IResourceReader_NewEnum() As stdole.IUnknown
    Set IResourceReader_NewEnum = NewEnum
End Function
