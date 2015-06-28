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
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: DSAParameters
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
' @param RHS The new Counter value.
'
Public Property Let Counter(ByVal RHS As Long)
    mCounter = RHS
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
' @param RHS The new G value.
'
Public Property Let g(ByRef RHS() As Byte)
    mG = g
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
' @param RHS The new J parameter.
'
Public Property Let j(ByRef RHS() As Byte)
    mJ = RHS
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
' @param RHS The new P parameter.
'
Public Property Let p(ByRef RHS() As Byte)
    mP = RHS
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
' @param RHS The new Q parameter.
'
Public Property Let q(ByRef RHS() As Byte)
    mQ = RHS
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
' @param RHS The new Seed value.
'
Public Property Let seed(ByRef RHS() As Byte)
    mSeed = RHS
End Property

''
' Gets the X parameter of the key.
'
' @return The X parameter.
'
Public Property Get X() As Byte()
    X = mX
End Property

''
' Sets the X parameter of the key.
'
' @param RHS The new X parameter.
'
Public Property Let X(ByRef RHS() As Byte)
    mX = RHS
End Property

''
' Gets the Y parameter of the key.
'
' @return The Y parameter.
'
Public Property Get Y() As Byte()
    Y = mY
End Property

''
' Sets the Y parameter of the key.
'
' @param RHS The new Y parameter.
'
Public Property Let Y(ByRef RHS() As Byte)
    mY = RHS
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub FromCspBlob(ByRef PublicBlob() As Byte, ByRef PrivateBlob() As Byte)
    Dim ms As MemoryStream
    Dim MagicNumber As Long
    
    If CorArray.IsNull(PrivateBlob) Then
        Set ms = Cor.NewMemoryStream(PublicBlob, Writable:=False)
        MagicNumber = &H31535344
    Else
        Set ms = Cor.NewMemoryStream(PrivateBlob, Writable:=False)
        MagicNumber = &H32535344
        
        ' Manually check for the public blob magic number.
        If AsLong(PublicBlob(8)) <> &H31535344 Then
            Throw Cor.NewCryptographicException("Invalid blob.")
        End If
    End If
    
    Dim Reader As BinaryReader
    Set Reader = Cor.NewBinaryReader(ms)

    ms.Position = 8
    
    ' Verify type of algorithm
    If Reader.ReadInt32 <> MagicNumber Then _
        Throw Cor.NewCryptographicException("Invalid blob.")
    
    Dim ByteLen As Long
    ByteLen = Reader.ReadInt32 \ 8
    
    Dim pBytes() As Byte
    Dim qBytes() As Byte
    
    pBytes = Reader.ReadBytes(ByteLen)
    qBytes = Reader.ReadBytes(20)
    
    mP = ReverseByteCopy(pBytes)
    mQ = ReverseByteCopy(qBytes)
    mG = ReverseByteCopy(Reader.ReadBytes(ByteLen))
    
    If CorArray.IsNull(PrivateBlob) Then
        mY = ReverseByteCopy(Reader.ReadBytes(ByteLen))
    Else
        mX = ReverseByteCopy(Reader.ReadBytes(20))
    
        ' We still need the Y parameter from the blob, but
        ' we have to retrieve it manually.
        ReDim mY(0 To ByteLen - 1)
        Dim Index As Long
        Index = CorArray.Length(PublicBlob) - (24 + ByteLen)
        Buffer.BlockCopy PublicBlob, Index, mY, 0, ByteLen
        CorArray.Reverse mY
    End If
    
    ' Read DSSSEED structure
    mCounter = Reader.ReadInt32
    mSeed = ReverseByteCopy(Reader.ReadBytes(20))
    
    Call Reader.CloseReader
    
    ReDim Preserve pBytes(UBound(pBytes) + 2)   ' ensure we have a positive number
    ReDim Preserve qBytes(UBound(qBytes) + 2)   ' ensure we have a positive number
    
    Dim pInt As BigInteger
    Dim qInt As BigInteger
    
    Set pInt = Cor.BInt(pBytes)
    Set qInt = Cor.BInt(qBytes)
    Set pInt = pInt.Subtract(BigInteger.One).Divide(qInt)
    mJ = ReverseByteCopy(pInt.ToArray)
End Sub

Friend Function ToCspBlob() As Byte()
    Dim ms As New MemoryStream
    Dim writer As BinaryWriter
    Set writer = Cor.NewBinaryWriter(ms)
    
    Dim BlobType As Byte
    If CorArray.IsNull(mX) Then
        BlobType = PUBLICKEYBLOB
    ElseIf CorArray.Length(mX) = 0 Then
        BlobType = PUBLICKEYBLOB
    Else
        BlobType = PRIVATEKEYBLOB
    End If
    
    ' Build BLOBHEADER structure
    Call writer.WriteValue(BlobType)
    Call writer.WriteValue(CByte(2))
    Call writer.WriteValue(CInt(0))
    Call writer.WriteValue(CALG_DSS_SIGN)
    
    ' Build DSSPUBKEY structure
    Call writer.WriteValue(CLng(IIf(BlobType = PUBLICKEYBLOB, &H31535344, &H32535344)))
    Call writer.WriteValue(CLng(CorArray.Length(mP) * 8))

    Call writer.WriteValue(ReverseByteCopy(mP))
    Call writer.WriteValue(ReverseByteCopy(mQ))
    Call writer.WriteValue(ReverseByteCopy(mG))
    
    If BlobType = PUBLICKEYBLOB Then
        Call writer.WriteValue(ReverseByteCopy(mY))
    Else
        Call writer.WriteValue(ReverseByteCopy(mX))
    End If
    
    ' Build DSSSEED structure
    Call writer.WriteValue(mCounter)
    Call writer.WriteValue(ReverseByteCopy(mSeed))
    
    ToCspBlob = ms.ToArray
End Function

Friend Function ToXmlString()
    Dim sb As New StringBuilder
    Call sb.AppendString("<DSAKeyValue>")
    Call AddElement(sb, "P", mP)
    Call AddElement(sb, "Q", mQ)
    Call AddElement(sb, "G", mG)
    Call AddElement(sb, "Y", mY)
    Call AddElement(sb, "J", mJ)
    Call AddElement(sb, "Seed", mSeed)
    
'    Dim Bytes(0 To 3) As Byte
'    AsLong(Bytes(0)) = mCounter
    Dim Bytes() As Byte
    Bytes = BitConverter.GetBytes(mCounter)
    
    Dim i As Long
    i = 3
    Do While i >= 0
        If Bytes(i) <> 0 Then Exit Do
        i = i - 1
    Loop
    ReDim Preserve Bytes(0 To i)
    
    Call AddElement(sb, "PgenCounter", Bytes)
    
    If Not CorArray.IsNull(mX) Then
        Call AddElement(sb, "X", mX)
    End If
    
    Call sb.AppendString("</DSAKeyValue>")
    
    ToXmlString = sb.ToString
End Function

Friend Sub FromXmlString(ByVal Xml As String)
    Dim Elements As SecurityElement
    Set Elements = SecurityElement.FromString(Xml)
    
    If Elements.Tag <> "DSAKeyValue" Then _
        Throw Cor.NewCryptographicException("Invalid Xml string.")
    
    mP = GetBytes(Elements, "P")
    mQ = GetBytes(Elements, "Q")
    mG = GetBytes(Elements, "G")
    mY = GetBytes(Elements, "Y")
    mSeed = GetBytes(Elements, "Seed")
    mCounter = GetBytes(Elements, "PgenCounter")
    mX = GetBytes(Elements, "X", False)
    mJ = GetBytes(Elements, "J", False)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddElement(ByVal sb As StringBuilder, ByVal Tag As String, ByRef Bytes() As Byte)
    Call sb.AppendFormat("<{0}>{1}</{0}>", Tag, Convert.ToBase64String(Bytes))
End Sub

Private Function GetBytes(ByVal Elements As SecurityElement, ByVal Name As String, Optional ByVal ThrowOnMissing As Boolean = True) As Byte()
    Dim Text As String
    
    Text = Elements.SearchForTextOfTag(Name)
    If (Len(Text) = 0) And ThrowOnMissing Then _
        Throw Cor.NewCryptographicException("Invalid Xml format.")
    
    GetBytes = Convert.FromBase64String(Text)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
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
        Call .WriteProperty(PROP_COUNTER, mCounter)
        Call .WriteProperty(PROP_G, mG)
        Call .WriteProperty(PROP_J, mJ)
        Call .WriteProperty(PROP_P, mP)
        Call .WriteProperty(PROP_Q, mQ)
        Call .WriteProperty(PROP_SEED, mSeed)
        Call .WriteProperty(PROP_Y, mY)
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
