VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceWriter"
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
' Module: ResourceWriter
'

''
' Writes encoded resources out to a .RES file.
'
Option Explicit
Implements IObject
Implements IResourceWriter

Private mWriter     As BinaryWriter
Private mEncoders   As New ArrayList
Private mGenerated  As Boolean


''
' Allows additional encoders to be added to the writer to encode
' additional datatypes to be written to a .RES file.
'
' @param EncoderToAdd The encodoer to be added to the writer.
' @remarks The encoder is inserted at the beginning of an internal
' list. This gives the most recent encoders first chance at encoding values.
'
Public Sub AddEncoder(ByVal EncoderToAdd As IResourceEncoder)
    VerifyNotGenerated
    
    If EncoderToAdd Is Nothing Then _
        Error.ArgumentNull "EncoderToAdd"
    
    EncoderToAdd.Reset
    
    ' We insert the encoder at the beginning of the list
    ' to allow it to have the first shot at encoding data.
    mEncoders.Insert 0, EncoderToAdd
End Sub

''
' Adds a new value to the writer.
'
' @param Value The value to be added to a .RES file.
' @param ResourceName The identifier of the value. Can be a string or number.
' @param ResourceType The type of resource being writting.
' @param LanguageID The Locale ID associated with the resource.
'
Public Sub AddResource(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant)
    VerifyNotGenerated
    EncodeResource Value, ResourceName, ResourceType, LanguageID
End Sub

''
' Creates a .RES file from the added resources.
'
Public Sub Generate()
    VerifyNotGenerated
    WriteResources
    mGenerated = True
End Sub

''
' Closes the writer.
'
Public Sub CloseWriter()
    If Not mWriter Is Nothing Then
        If Not mGenerated Then
            Generate
        End If
        
        mWriter.CloseWriter
        Set mWriter = Nothing
    End If
End Sub

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compare with this instance.
' @return Returns True if the value equals this instance, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return Returns a hashcode value.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return Returns a string representation of this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mEncoders.Add New BinaryResourceEncoder
    mEncoders.Add New StringResourceEncoder
    mEncoders.Add New BitMapResourceEncoder
    mEncoders.Add New CursorResourceEncoder
    mEncoders.Add New CursorResourceGroupEncoder
    mEncoders.Add New IconResourceEncoder
    mEncoders.Add New IconResourceGroupEncoder
End Sub

Friend Sub Init(ByRef Source As Variant)
    Set mWriter = Cor.NewBinaryWriter(GetStream(Source))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetStream(ByRef Source As Variant) As Stream
    Select Case VarType(Source)
        Case vbString
            ' We will overwrite any existing file.
            Set GetStream = Cor.NewFileStream(CStr(Source), FileMode.Create)
        Case vbObject
            If Source Is Nothing Then _
                Error.ArgumentNull "Source"
            If Not TypeOf Source Is Stream Then _
                Error.Argument Argument_InvalidStreamSource, "Source"
                
            Set GetStream = Source
            
            If Not GetStream.CanWrite Then _
                Error.Argument Argument_StreamNotWritable, "Source"
        Case Else
            Error.Argument Argument_InvalidStreamSource, "Source"
    End Select
End Function

Private Sub VerifyNotGenerated()
    If mGenerated Then
        Error.InvalidOperation InvalidOperation_ResourceWriterSaved
    End If
End Sub

Private Sub WriteResources()
    Dim Encoder As IResourceEncoder
    
    WriteBlankHeader
    
    For Each Encoder In mEncoders
        WriteEncodedResources Encoder
    Next Encoder
End Sub

Private Sub WriteBlankHeader()
    mWriter.WriteValue Cor.NewBytes(0, 0, 0, 0, &H20, 0, 0, 0, &HFF, &HFF, 0, 0, &HFF, &HFF, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
End Sub

Private Sub WriteEncodedResources(ByVal Encoder As IResourceEncoder)
    Dim Key     As ResourceKey
    Dim Data()  As Byte
    
    ' Each encoder may have multiple resources that
    ' need to be written, so we will keep writing
    ' them from the encoder until it says done.
    Do While Encoder.GetEncodedResource(Key, Data)
        WriteResource Key, Data
    Loop
End Sub

Private Sub WriteResource(ByRef Key As ResourceKey, ByRef Data() As Byte)
    Const RESOURCEFLAG_MOVEABLE As Integer = &H10
    Const RESOURCEFLAG_PURE As Integer = &H20
    Const RESOURCEFLAG_DISCARDABLE As Integer = &H1000
    Const MEMORY_FLAGS As Integer = RESOURCEFLAG_MOVEABLE Or RESOURCEFLAG_PURE Or RESOURCEFLAG_DISCARDABLE
    Const MinHeaderSize As Long = 24
    
    Dim i As Long
    Dim EncodedName()           As Byte
    Dim EncodedType()           As Byte
    Dim EncodedTypeAndNameSize  As Long
    
    ' Write the number of bytes in the data portion of the resource.
    mWriter.WriteValue CorArray.Length(Data)
    
    ' Calculate the size of the header. Requires the
    ' length of the resource type and name.
    EncodedType = EncodeStringOrInt(Key.ResourceType)
    EncodedName = EncodeStringOrInt(Key.ResourceName)
    EncodedTypeAndNameSize = CorArray.Length(EncodedName) + CorArray.Length(EncodedType)
    
    ' 24 is the size of the header without the name and type,
    ' so that is our base size to start with.
    mWriter.WriteValue MinHeaderSize + EncodedTypeAndNameSize + ComputeDWordAlignedOffset(EncodedTypeAndNameSize)
    mWriter.WriteValue EncodedType
    mWriter.WriteValue EncodedName
    
    ' We must DWord align after both the resource name
    ' and resource type have been written.
    For i = 1 To ComputeDWordAlignedOffset(EncodedTypeAndNameSize)
        mWriter.WriteValue CByte(0)
    Next i
    
    ' DataVersion is zero. Must be a vbLong though
    ' so we will write 4 bytes.
    mWriter.WriteValue CLng(0)
    
    mWriter.WriteValue MEMORY_FLAGS
    
    ' The LanguageID must be 2 bytes, also.
    mWriter.WriteValue CInt(Key.LanguageID)
    
    ' Version is zero and 4 bytes.
    mWriter.WriteValue CLng(0)
    
    ' Characteristics is zero and 4 bytes.
    mWriter.WriteValue CLng(0)
    
    ' Write the resource data.
    mWriter.WriteValue Data

    ' DWord align the end of the entry for the next entry.
    For i = 1 To ComputeDWordAlignedOffset(CorArray.Length(Data))
        mWriter.WriteValue CByte(0)
    Next i
End Sub

Private Function ComputeDWordAlignedOffset(ByVal Size As Long) As Long
    If (Size Mod 4) <> 0 Then
        ComputeDWordAlignedOffset = 4 - (Size Mod 4)
    End If
End Function

Private Function EncodeStringOrInt(ByRef Value As Variant) As Byte()
    If VarType(Value) = vbString Then
        EncodeStringOrInt = CStr(Value) & Chr$(0)
    Else
        ' A numeric name is identified by being a DWord with the
        ' lower 16-bits being &HFFFF.
        Dim ID As Long
        ID = CLng(Value) And &HFFFF&
        EncodeStringOrInt = BitConverter.GetBytes(ShiftLeftInt32(ID, 16) Or &HFFFF&)
    End If
End Function

Private Sub EncodeResource(ByRef Value As Variant, ByRef ResourceName As Variant, ByRef ResourceType As Variant, ByRef LanguageID As Variant)
    Dim Encoder As IResourceEncoder
    
    For Each Encoder In mEncoders
        ' Give the encoder a shot.
        If Encoder.Encode(Value, ResourceName, ResourceType, LanguageID) Then
            Exit Sub
        End If
    Next

    Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_UnsupportedResourceType, TypeName(Value)))
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
'   IResourceWriter
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceWriter_AddResource(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant)
    AddResource Value, ResourceName, ResourceType, LanguageID
End Sub

Private Sub IResourceWriter_CloseWriter()
    CloseWriter
End Sub

Private Sub IResourceWriter_Generate()
    Generate
End Sub
