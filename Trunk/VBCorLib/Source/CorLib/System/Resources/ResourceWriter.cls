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

Private mStream             As Stream
Private mWriter             As BinaryWriter
Private mEncoders           As New ArrayList
Private mDefaultEncoder     As IResourceEncoder
Private mIsAlreadyGenerated As Boolean


''
' Allows additional encoders to be added to the writer to encode
' additional datatypes to be written to a .RES file.
'
' @param EncoderToAdd The encodoer to be added to the writer.
' @remarks The encoder is inserted at the beginning of an internal
' list. This gives the most recent encoders first chance at encoding values.
'
Public Sub AddEncoder(ByVal EncoderToAdd As IResourceEncoder)
    Call VerifyNotGenerated
    Call VerifyOpen
    
    If EncoderToAdd Is Nothing Then _
        Throw Cor.NewArgumentNullException("EncoderToAdd")
    
    Call EncoderToAdd.Reset
    
    ' We insert the encoder at the beginning of the list
    ' to allow it to have the first shot at encoding data.
    Call mEncoders.Insert(0, EncoderToAdd)
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
    Call VerifyNotGenerated
    Call VerifyOpen
    
    If IsArray(Value) Then
        If VarType(Value) = corByteArray Then
            If IsMissing(ResourceType) Then
                Throw Cor.NewArgumentException("A ResourceType must be supplied for Byte Arrays.", "ResourceType")
            End If
        Else
            Throw Cor.NewArgumentException("Only Byte Arrays are supported.", "Value")
        End If
    End If
    
    Call EncodeResource(Value, ResourceName, ResourceType, LanguageID)
End Sub

''
' Creates a .RES file from the added resources.
'
Public Sub Generate()
    Call VerifyNotGenerated
    Call VerifyOpen
       
    mIsAlreadyGenerated = True
    Call WriteResources
End Sub

''
' Closes the writer.
'
Public Sub CloseWriter()
    If Not mStream Is Nothing Then
        Call mWriter.CloseWriter
        Set mWriter = Nothing
        Set mStream = Nothing
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Source As Variant)
    Select Case VarType(Source)
        Case vbString
            ' We will overwrite any existing file.
            Set mStream = Cor.NewFileStream(CStr(Source), FileMode.Create)
        Case vbObject
            If TypeOf Source Is Stream Then
                Set mStream = Source
                If Not mStream.CanRead Then Throw Cor.NewArgumentException("The Stream must support reading.", "Source")
            Else
                Throw Cor.NewArgumentException("Source must implement Stream interface", "Source")
            End If
        Case Else
            Throw Cor.NewArgumentException("Source must be a file path or file stream.", "Source")
    End Select

    Set mWriter = Cor.NewBinaryWriter(mStream)
    
    Call AddEncoder(New ResStringTableEncoder)
    Call AddEncoder(New ResBitMapEncoder)
    Call AddEncoder(New ResIconEncoder)
    Call AddEncoder(New ResIconGroupEncoder)
    Call AddEncoder(New ResCursorEncoder)
    Call AddEncoder(New ResCursorGroupEncoder)
    
    Set mDefaultEncoder = New ResByteEncoder
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyOpen()
    If mStream Is Nothing Then Throw Cor.NewInvalidOperationException("The writer has been closed.")
End Sub

Private Sub VerifyNotGenerated()
    If mIsAlreadyGenerated Then Throw Cor.NewInvalidOperationException("Resources have already been generated.")
End Sub

Private Sub WriteResources()
    ' Write a blank entry. This identifies the .RES file.
    Call WriteResource(Cor.NewResourceKey(0, 0, 0), Cor.NewBytes())
    
    Dim Encoder As IResourceEncoder
    For Each Encoder In mEncoders
        Call WriteEncodedResources(Encoder)
    Next Encoder
    
    ' And add any default encoded resources.
    Call WriteEncodedResources(mDefaultEncoder)
End Sub

Private Sub WriteEncodedResources(ByVal Encoder As IResourceEncoder)
    Dim Key As ResourceKey
    Dim Data As Variant
    
    ' Each encoder may have multiple resources that
    ' need to be written, so we will keep writing
    ' them from the encoder until it says done.
    Do While Encoder.GetEncodedResource(Key, Data)
        Call WriteResource(Key, Data)
    Loop
End Sub

Private Sub WriteResource(ByRef Key As ResourceKey, ByRef Data As Variant)
    Dim EncodedName() As Byte
    Dim EncodedType() As Byte
    Dim i As Long
    
    If VarType(Data) <> corByteArray Then _
        Throw Cor.NewInvalidCastException("Resource Encoder must return a Byte Array.")
    
    ' Write the number of bytes in the data portion of the resource.
    Call mWriter.WriteValue(CorArray.GetLength(Data))
    
    ' Calculate the size of the header. Requires the
    ' length of the resource type and name.
    EncodedType = EncodeStringOrInt(Key.ResourceType)
    EncodedName = EncodeStringOrInt(Key.ResourceName)
    
    Dim EncodedTypeAndNameSize As Long
    EncodedTypeAndNameSize = CorArray.GetLength(EncodedName) + CorArray.GetLength(EncodedType)
    
    
    ' 24 is the size of the header without the name and type,
    ' so that is our base size to start with.
    Call mWriter.WriteValue(24 + EncodedTypeAndNameSize + GetDWordAlignedOffset(EncodedTypeAndNameSize))
    
    ' Write the resource type.
    Call mWriter.WriteValue(EncodedType)
    
    ' Write the resource name or ordinal.
    Call mWriter.WriteValue(EncodedName)
    
    ' We must DWord align after both the resource name
    ' and resource type have been written.
    For i = 1 To GetDWordAlignedOffset(EncodedTypeAndNameSize)
        Call mWriter.WriteValue(CByte(0))
    Next i
    
    ' DataVersion is zero. Must be a vbLong though
    ' so we will write 4 bytes.
    Call mWriter.WriteValue(0&)
    
    ' MemoryFlags is zero. Must be a vbInteger
    ' so we will write 2 bytes.
    Call mWriter.WriteValue(0)
    'Call mWriter.WriteValue(&H3010)
    
    ' The LanguageID must be 2 bytes, also.
    Call mWriter.WriteValue(CInt(Key.LanguageID))
    
    ' Version is zero and 4 bytes.
    Call mWriter.WriteValue(0&)
    
    ' Characteristics is zero and 4 bytes.
    Call mWriter.WriteValue(0&)
    
    ' Write the resource data.
    Call mWriter.WriteValue(Data)

    ' DWord align the end of the entry for the next entry.
    For i = 1 To GetDWordAlignedOffset(CorArray.GetLength(Data))
        Call mWriter.WriteValue(CByte(0))
    Next i
End Sub

Private Function GetDWordAlignedOffset(ByVal Size As Long) As Long
    If (Size Mod 4) <> 0 Then
        GetDWordAlignedOffset = 4 - (Size Mod 4)
    End If
End Function

Private Function EncodeStringOrInt(ByRef Value As Variant) As Byte()
    If VarType(Value) = vbString Then
        ' VB will assign a string to a byte array, copying each
        ' 2-byte character to 2 bytes in the array.
        Dim Ret() As Byte
        
        ' We include 2 bytes for terminating null.
        ReDim Ret(0 To LenB(Value) + 1)
        Call CopyMemory(Ret(0), ByVal StrPtr(Value), LenB(Value))
        EncodeStringOrInt = Ret
        Exit Function
    Else
        ' we assume a numeric value. It must be within
        ' the range of vbInteger values.
        '
        ' A numeric name is identified by being a DWord with the
        ' lower 16-bits being &HFFFF.
        Dim ID As Long
        AsWord(ID) = Value  ' removes the negative bit.
        EncodeStringOrInt = BitConverter.GetBytes(Helper.ShiftLeft(ID, 16) Or &HFFFF&)
    End If
End Function

Private Sub EncodeResource(ByRef Value As Variant, ByRef ResourceName As Variant, ByRef ResourceType As Variant, ByRef LanguageID As Variant)
    Dim Encoded As Boolean
    Encoded = False
    Dim Encoder As IResourceEncoder
    For Each Encoder In mEncoders
        ' Give the encoder a shot, if it succeeds, then
        ' move on to the next resource to be encoded.
        If Encoder.Encode(Value, ResourceName, ResourceType, LanguageID) Then
            Encoded = True
            Exit For
        End If
    Next Encoder

    ' The default encoder will only work for byte arrays.
    If Not Encoded Then
        If Not mDefaultEncoder.Encode(Value, ResourceName, ResourceType, LanguageID) Then
            Throw Cor.NewArgumentException("Resource type is not supported.")
        End If
    End If
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
'   IResourceWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceWriter_AddResource(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant)
    Call AddResource(Value, ResourceName, ResourceType, LanguageID)
End Sub

Private Sub IResourceWriter_CloseWriter()
    Call CloseWriter
End Sub

Private Sub IResourceWriter_Generate()
    Call Generate
End Sub
