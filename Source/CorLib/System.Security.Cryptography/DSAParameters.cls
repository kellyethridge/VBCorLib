VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSAParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: DSAParameters
'

''
' Represents the DSA key parameters.
'
Option Explicit
Implements IObject

Private Const PROP_COUNTER  As String = "Counter"
Private Const PROP_G        As String = "G"
Private Const PROP_J        As String = "J"
Private Const PROP_P        As String = "P"
Private Const PROP_Q        As String = "Q"
Private Const PROP_SEED     As String = "Seed"
Private Const PROP_Y        As String = "Y"

Private Const SizeOfHeader      As Long = 8
Private Const SizeOfDssPubKey   As Long = 8
Private Const SizeOfQ           As Long = 20
Private Const SizeOfX           As Long = 20
Private Const SizeOfSeed        As Long = 20

Private Type DssSeed
    Counter As Long
    seed()  As Byte
End Type

Private mCounter    As Long
Private mG()        As Byte
Private mJ()        As Byte
Private mP()        As Byte
Private mQ()        As Byte
Private mSeed()     As Byte
Private mX()        As Byte
Private mY()        As Byte


''
' Gets the Counter value for the DSA key.
'
' @return Returns the Counter value.
'
Public Property Get Counter() As Long
    Counter = mCounter
End Property

''
' Sets the Counter value for the DSA key.
'
' @param Value The new Counter value.
'
Public Property Let Counter(ByVal Value As Long)
    mCounter = Value
End Property

''
' Gets the G parameter of the key.
'
' @return The G parameter.
'
Public Property Get g() As Byte()
    g = mG
End Property

''
' Sets the G parameter of the key.
'
' @param Value The new G value.
'
Public Property Let g(ByRef Value() As Byte)
    mG = Value
End Property

''
' Gets the J parameter of the key.
'
' @return The J parameter.
'
Public Property Get j() As Byte()
    j = mJ
End Property

''
' Sets the J parameter of the key.
'
' @param Value The new J parameter.
'
Public Property Let j(ByRef Value() As Byte)
    mJ = Value
End Property

''
' Gets the P parameter of the key.
'
' @return The P parameter.
'
Public Property Get p() As Byte()
    p = mP
End Property

''
' Sets the P parameter of the key.
'
' @param Value The new P parameter.
'
Public Property Let p(ByRef Value() As Byte)
    mP = Value
End Property

''
' Gets the Q parameter of the key.
'
' @return The Q parameter.
'
Public Property Get q() As Byte()
    q = mQ
End Property

''
' Sets the Q parameter of the key.
'
' @param Value The new Q parameter.
'
Public Property Let q(ByRef Value() As Byte)
    mQ = Value
End Property

''
' Gets the Seed of the key.
'
' @return The Seed value.
'
Public Property Get seed() As Byte()
    seed = mSeed
End Property

''
' Sets the Seed for the key.
'
' @param Value The new Seed value.
'
Public Property Let seed(ByRef Value() As Byte)
    mSeed = Value
End Property

''
' Gets the X parameter of the key.
'
' @return The X parameter.
'
Public Property Get x() As Byte()
    x = mX
End Property

''
' Sets the X parameter of the key.
'
' @param Value The new X parameter.
'
Public Property Let x(ByRef Value() As Byte)
    mX = Value
End Property

''
' Gets the Y parameter of the key.
'
' @return The Y parameter.
'
Public Property Get y() As Byte()
    y = mY
End Property

''
' Sets the Y parameter of the key.
'
' @param Value The new Y parameter.
'
Public Property Let y(ByRef Value() As Byte)
    mY = Value
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
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

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

Friend Sub FromPrivateKeyBlob(ByRef PrivateKeyBlob() As Byte, ByRef PublicKeyBlob() As Byte, ByRef ParameterJ() As Byte)
    FromPublicKeyBlob PublicKeyBlob, ParameterJ
    LoadPrivateParameter PrivateKeyBlob
End Sub

Friend Sub FromPublicKeyBlob(ByRef Blob() As Byte, ByRef ParameterJ() As Byte)
    Dim Reader      As BinaryReader
    Dim ByteLength  As Long
    
    Set Reader = MakeReader(Blob, SizeOfHeader)
    ByteLength = ReadByteLength(Reader)
    LoadPublicParameters Reader, ByteLength
    mJ = ParameterJ
End Sub

Friend Function ToCspBlob() As Byte()
    If IsPublicParameters Then
        ToCspBlob = BuildPublicCspBlob
    Else
        ToCspBlob = BuildPrivateCspBlob
    End If
End Function

Friend Function ToXmlString()
    Dim Root As SecurityElement
    Set Root = MakeXmlRoot
    AddPublicXmlParameters Root
    AddPrivateXmlParameters Root
    
    ToXmlString = Root.ToString
End Function

Friend Sub FromXmlString(ByRef Xml As String)
    Dim Root As SecurityElement
    Set Root = GetXmlRoot(Xml)
    ReadXmlParameters Root
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadPrivateParameter(ByRef Blob() As Byte)
    Dim Reader As BinaryReader
    
    Set Reader = MakeReader(Blob, SizeOfHeader + SizeOfDssPubKey + CorArray.Length(mP) + SizeOfQ + CorArray.Length(mG))
    mX = ReadBytesReverse(Reader, SizeOfX)
End Sub

Private Sub LoadPublicParameters(ByVal Reader As BinaryReader, ByVal ByteLength As Long)
    Dim DssSeed     As DssSeed
    
    mP = ReadBytesReverse(Reader, ByteLength)
    mQ = ReadBytesReverse(Reader, SizeOfQ)
    mG = ReadBytesReverse(Reader, ByteLength)
    mY = ReadBytesReverse(Reader, ByteLength)
    DssSeed = ReadDssSeed(Reader)
    mCounter = DssSeed.Counter
    mSeed = DssSeed.seed
End Sub

Private Function MakeReader(ByRef Blob() As Byte, ByVal InitialPosition As Long) As BinaryReader
    Dim Stream As MemoryStream
    
    Set Stream = Cor.NewMemoryStream(Blob, Writable:=False)
    Stream.Position = InitialPosition
    Set MakeReader = Cor.NewBinaryReader(Stream)
End Function

Private Function ReadBytesReverse(ByVal Reader As BinaryReader, ByVal Count As Long) As Byte()
    Dim Bytes() As Byte
    
    Bytes = Reader.ReadBytes(Count)
    CorArray.Reverse Bytes
    ReadBytesReverse = Bytes
End Function

Private Function ReadByteLength(ByVal Reader As BinaryReader) As Long
    Const BitsPerByte As Long = 8
    Reader.ReadInt32
    ReadByteLength = Reader.ReadInt32 \ BitsPerByte
End Function

Private Function ReadDssSeed(ByVal Reader As BinaryReader) As DssSeed
    ReadDssSeed.Counter = Reader.ReadInt32
    ReadDssSeed.seed = ReadBytesReverse(Reader, SizeOfSeed)
End Function

Private Property Get IsPublicParameters() As Boolean
    IsPublicParameters = CorArray.IsNull(mX)
End Property

Private Function BuildPublicCspBlob() As Byte()
    Dim ms      As New MemoryStream
    Dim Writer  As BinaryWriter

    Set Writer = Cor.NewBinaryWriter(ms)
    WritePublicCspBlob Writer
    BuildPublicCspBlob = ms.ToArray
End Function

Private Sub WritePublicCspBlob(ByVal Writer As BinaryWriter)
    Const DSS1Chars As Long = &H31535344
    WriteCspBlobHeader Writer, PublicKeyBlob, DSS1Chars
    WriteCspBlobParameters Writer, mY
End Sub

Private Function BuildPrivateCspBlob() As Byte()
    Dim ms      As New MemoryStream
    Dim Writer  As BinaryWriter

    Set Writer = Cor.NewBinaryWriter(ms)
    WritePrivateCspBlob Writer
    BuildPrivateCspBlob = ms.ToArray
End Function

Private Sub WritePrivateCspBlob(ByVal Writer As BinaryWriter)
    Const DSS2Chars As Long = &H32535344
    WriteCspBlobHeader Writer, PrivateKeyBlob, DSS2Chars
    WriteCspBlobParameters Writer, mX
End Sub

Private Sub WriteCspBlobHeader(ByVal Writer As BinaryWriter, ByVal BlobType As Byte, ByVal DSSChars As Long)
    Writer.WriteValue BlobType
    Writer.WriteValue CByte(2)
    Writer.WriteValue CInt(0)
    Writer.WriteValue CALG_DSS_SIGN
    Writer.WriteValue DSSChars
    Writer.WriteValue CorArray.Length(mP) * 8
End Sub

Private Sub WriteCspBlobParameters(ByVal Writer As BinaryWriter, ByRef KeyParameter() As Byte)
    Writer.WriteValue ReverseByteCopy(mP)
    Writer.WriteValue ReverseByteCopy(mQ)
    Writer.WriteValue ReverseByteCopy(mG)
    Writer.WriteValue ReverseByteCopy(KeyParameter)
    Writer.WriteValue mCounter
    Writer.WriteValue ReverseByteCopy(mSeed)
End Sub

Private Function MakeXmlRoot() As SecurityElement
    Set MakeXmlRoot = Cor.NewSecurityElement("DSAKeyValue")
End Function

Private Sub AddPublicXmlParameters(ByVal Root As SecurityElement)
    Root.AddChild MakeXmlElement("P", mP)
    Root.AddChild MakeXmlElement("Q", mQ)
    Root.AddChild MakeXmlElement("G", mG)
    Root.AddChild MakeXmlElement("Y", mY)
    
    AddXmlJParameter Root
    
    Root.AddChild MakeXmlElement("Seed", mSeed)
    Root.AddChild MakeXmlElement("PgenCounter", GetTrimmedCounter)
End Sub

Private Sub AddXmlJParameter(ByVal Root As SecurityElement)
    AddOptionalXmlParameter Root, "J", mJ
End Sub

Private Sub AddPrivateXmlParameters(ByVal Root As SecurityElement)
    AddOptionalXmlParameter Root, "X", mX
End Sub

Private Sub AddOptionalXmlParameter(ByVal Root As SecurityElement, ByRef Tag As String, ByRef Value() As Byte)
    If Not CorArray.IsNull(Value) Then
        Root.AddChild MakeXmlElement(Tag, Value)
    End If
End Sub

Private Function MakeXmlElement(ByRef Tag As String, ByRef Value() As Byte) As SecurityElement
    Dim Xml As SecurityElement
    
    Set Xml = Cor.NewSecurityElement(Tag)
    Xml.Text = Convert.ToBase64String(Value)
    
    Set MakeXmlElement = Xml
End Function

Private Function GetTrimmedCounter() As Byte()
    Dim Bytes() As Byte
    Dim i       As Long
    
    Bytes = BitConverter.GetBytes(mCounter)
    i = UBound(Bytes)
    
    Do While i >= 0
        If Bytes(i) <> 0 Then
            Exit Do
        End If
        
        i = i - 1
    Loop
    
    ReDim Preserve Bytes(0 To i)

    GetTrimmedCounter = Bytes
End Function

Private Function GetXmlRoot(ByRef Xml As String) As SecurityElement
    Dim Root As SecurityElement
    Set Root = SecurityElement.FromString(Xml)
    
    If Root.Tag <> "DSAKeyValue" Then _
        Throw Cor.NewCryptographicException("Invalid Xml string.")

    Set GetXmlRoot = Root
End Function

Private Sub ReadXmlParameters(ByVal Root As SecurityElement)
    mP = GetBytes(Root, "P")
    mQ = GetBytes(Root, "Q")
    mG = GetBytes(Root, "G")
    mY = GetBytes(Root, "Y")
    mJ = GetOptionalBytes(Root, "J")
    mSeed = GetBytes(Root, "Seed")
    mCounter = GetNumber(Root, "PgenCounter")
    mX = GetOptionalBytes(Root, "X")
End Sub

Private Function GetNumber(ByVal Element As SecurityElement, ByRef Name As String) As Long
    Dim Bytes() As Byte
    
    Bytes = GetBytes(Element, Name)
    
    Select Case CorArray.Length(Bytes)
        Case 1
            GetNumber = Bytes(0)
        Case 2
            GetNumber = BitConverter.ToInt16(Bytes, 0)
        Case Else
            GetNumber = BitConverter.ToInt32(Bytes, 0)
    End Select
End Function

Private Function GetBytes(ByVal Root As SecurityElement, ByRef Tag As String) As Byte()
    Dim Text As String
    
    Text = Root.SearchForTextOfTag(Tag)
    If Len(Text) = 0 Then _
        Throw Cor.NewCryptographicException("Invalid Xml format.")
    
    GetBytes = Convert.FromBase64String(Text)
End Function

Private Function GetOptionalBytes(ByVal Root As SecurityElement, ByRef Tag As String) As Byte()
    Dim Text As String
    
    Text = Root.SearchForTextOfTag(Tag)
    If Len(Text) > 0 Then
        GetOptionalBytes = Convert.FromBase64String(Text)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mCounter = .ReadProperty(PROP_COUNTER)
        mG = .ReadProperty(PROP_G)
        mJ = .ReadProperty(PROP_J)
        mP = .ReadProperty(PROP_P)
        mQ = .ReadProperty(PROP_Q)
        mSeed = .ReadProperty(PROP_SEED)
        mY = .ReadProperty(PROP_Y)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_COUNTER, mCounter
        .WriteProperty PROP_G, mG
        .WriteProperty PROP_J, mJ
        .WriteProperty PROP_P, mP
        .WriteProperty PROP_Q, mQ
        .WriteProperty PROP_SEED, mSeed
        .WriteProperty PROP_Y, mY
    End With
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
