VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSAParameters"
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
' Module: RSAParameters
'

''
' Represents the Key information for an RSA implementation.
'
' @remarks To learn more about RSAParameters, please visit this Shawnfa blog entry:
' http://blogs.msdn.com/shawnfa/archive/2005/11/17/493972.aspx
'
Option Explicit
Implements IObject

Private Const PropModulus  As String = "Modulus"
Private Const PropExponent As String = "Exponent"

Private mD()        As Byte ' PrivateExponent
Private mDP()       As Byte ' Exponent1
Private mDQ()       As Byte ' Exponent2
Private mExponent() As Byte ' PublicExponent
Private mInverseQ() As Byte ' Coefficient
Private mModulus()  As Byte ' P * Q (Prime1 * Prime2 = n)
Private mP()        As Byte ' Prime1
Private mQ()        As Byte ' Prime2



''
' Returns the D parameter of the RSA key.
'
' @return A byte array of the D parameter.
' @remarks This is the Private Exponent of the RSA key.
'
Public Property Get d() As Byte()
    d = mD
End Property

''
' Sets the D parameter of the RSA key.
'
' @param Value The new D parameter value.
' @remarks This is the Private Exponent of the RSA key.
'
Public Property Let d(ByRef Value() As Byte)
    mD = Value
End Property

''
' Returns the DP parameter of the key.
'
' @return A byte array of the DP parameter.
' @remarks This is the Exponent1 value of the key. It has the numeric value "d mod (p - 1)".
'
Public Property Get DP() As Byte()
    DP = mDP
End Property

''
' Sets the DP parameter of the key.
'
' @param Value The new DP value.
' @remarks This is the Exponent1 value of the key. It has the numeric value "d mod (p - 1)".
'
Public Property Let DP(ByRef Value() As Byte)
    mDP = Value
End Property

''
' Returns the DQ parameter of the key.
'
' @return A byte array of the DQ paramter
' @remarks This is the Exponent2 value of the key. It has the numeric value "d mod (q - 1)".
'
Public Property Get DQ() As Byte()
    DQ = mDQ
End Property

''
' Sets the DQ parameter of the key.
'
' @param Value The new DQ value.
' @remarks This is the Exponent2 value of the key. It has the numeric value "d mod (q - 1)".
'
Public Property Let DQ(ByRef Value() As Byte)
    mDQ = Value
End Property

''
' Returns the public exponent of the key.
'
' @return A byte array of the public exponent.
'
Public Property Get Exponent() As Byte()
    Exponent = mExponent()
End Property

''
' Sets the public exponent of the key.
'
' @param Value The new exponent value.
'
Public Property Let Exponent(ByRef Value() As Byte)
    mExponent = Value
End Property

''
' Returns InverseQ parameter of the key.
'
' @return A byte array of the InverseQ parameter.
' @remarks This is the Coefficient. It has a numeric value "(inverse of q) mod p".
'
Public Property Get InverseQ() As Byte()
    InverseQ = mInverseQ
End Property

''
' Sets the InverseQ parameter of the key
'
' @param Value The new InverseQ value.
' @remarks This is the Coefficient. It has a numeric value "(inverse of q) mod p".
'
Public Property Let InverseQ(ByRef Value() As Byte)
    mInverseQ = Value
End Property

''
' Returns the modulus of the key.
'
' @return A byte array of the modulus of the key.
' @remarks This has a value of "prime1 * prime2" and is often known as "n".
'
Public Property Get Modulus() As Byte()
    Modulus = mModulus
End Property

''
' Sets the modulus of the key.
'
' @param Value The new modulus value.
' @remarks This has a value of "prime1 * prime2" and is often known as "n".
'
Public Property Let Modulus(ByRef Value() As Byte)
    mModulus = Value
End Property

''
' Returns the P parameter of the key.
'
' @return A byte array of the P parameter.
' @remarks This is the Prime1 value of the key.
'
Public Property Get p() As Byte()
    p = mP
End Property

''
' Sets the P parameter of the key.
'
' @param Value The new P value.
' @remarks This is the Prime1 value of the key.
'
Public Property Let p(ByRef Value() As Byte)
    mP = Value
End Property

''
' Returns the Q parameter of the key.
'
' @return A byte array of the Q value.
' @remarks This is the Prime2 value of the key.
'
Public Property Get q() As Byte()
    q = mQ
End Property

''
' Sets the Q parameter of the key.
'
' @param Value The new Q value.
' @remarks This is the Prime2 value of the key.
'
Public Property Let q(ByRef Value() As Byte)
    mQ = Value
End Property

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to test for equality.
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
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

Friend Function ToXmlString() As String
    Dim sb As StringBuilder
    Set sb = StringBuilderCache.Acquire
    
    sb.AppendString "<RSAKeyValue>"
    AddElement sb, "Modulus", mModulus
    AddElement sb, "Exponent", mExponent
    
    If Not CorArray.IsNull(mD) Then
        AddElement sb, "P", mP
        AddElement sb, "Q", mQ
        AddElement sb, "DP", mDP
        AddElement sb, "DQ", mDQ
        AddElement sb, "InverseQ", mInverseQ
        AddElement sb, "D", mD
    End If
    
    sb.AppendString "</RSAKeyValue>"
    
    ToXmlString = StringBuilderCache.GetStringAndRelease(sb)
End Function

Friend Sub FromXmlString(ByRef Xml As String)
    Dim Elements As SecurityElement
    Set Elements = SecurityElement.FromString(Xml)
    
    If Elements.Tag <> "RSAKeyValue" Then _
        Throw Cor.NewCryptographicException("Invalid Xml string.")
    
    mModulus = GetBytes(Elements, "Modulus")
    mExponent = GetBytes(Elements, "Exponent")
    
    If Elements.Children.Count = 8 Then
        ' we assume there is private data
        mP = GetBytes(Elements, "P")
        mQ = GetBytes(Elements, "Q")
        mDP = GetBytes(Elements, "DP")
        mDQ = GetBytes(Elements, "DQ")
        mInverseQ = GetBytes(Elements, "InverseQ")
        mD = GetBytes(Elements, "D")
    End If
End Sub

Friend Sub FromCspBlob(ByRef Blob() As Byte)
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(Blob, Writable:=False)
    
    Dim Reader As BinaryReader
    Set Reader = Cor.NewBinaryReader(ms)
    
    Dim BlobType As Long
    BlobType = Reader.ReadByte
    
    ms.Position = 8
    
    ' read the magic number.
    If Reader.ReadInt32 <> CLng(IIf(BlobType = PUBLICKEYBLOB, &H31415352, &H32415352)) Then _
        Throw Cor.NewCryptographicException("Invalid blob")
        
    Dim BitLen As Long
    BitLen = Reader.ReadInt32
    
    mExponent = Reader.ReadBytes(4)
    
    ' Remove leading zeros. The array is little-endian,
    ' so the leading zeros will be at the end
    Dim i As Long
    i = 3
    Do While mExponent(i) = 0
        i = i - 1
    Loop
    If i < 3 Then
        ReDim Preserve mExponent(0 To i)
    End If
    
    ' make the array big-endian
    CorArray.Reverse mExponent
    
    ' start reading in all the key stuff.
    mModulus = ReverseByteCopy(Reader.ReadBytes(BitLen \ 8))
    
    If BlobType = PRIVATEKEYBLOB Then
        mP = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mQ = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mDP = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mDQ = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mInverseQ = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mD = ReverseByteCopy(Reader.ReadBytes(BitLen \ 8))
    End If
    
    Reader.CloseReader
End Sub

Friend Function ToCspBlob(ByVal KeyNumber As Long) As Byte()
    Dim ms As New MemoryStream
    Dim Writer As BinaryWriter
    Set Writer = Cor.NewBinaryWriter(ms)
    
    Dim BlobType As Byte
    BlobType = IIf(CorArray.IsNull(mD), PUBLICKEYBLOB, PRIVATEKEYBLOB)
    
    ' Build BLOBHEADER structure
    Writer.WriteValue BlobType
    Writer.WriteValue CByte(2)
    Writer.WriteValue CInt(0)
    Writer.WriteValue CLng(IIf(KeyNumber = AT_KEYEXCHANGE, CALG_RSA_KEYX, CALG_RSA_SIGN))
    
    ' Build RSAPUBKEY structure
    Writer.WriteValue CLng(IIf(BlobType = PUBLICKEYBLOB, &H31415352, &H32415352))
    Writer.WriteValue CLng(CorArray.Length(mModulus) * 8)
    Writer.WriteValue ReverseByteCopy(mExponent)
    
    Dim i As Long
    i = 4 - Len1D(mExponent)
    Do While i > 0
        Writer.WriteValue CByte(0)
        i = i - 1
    Loop
    
    ' Start adding all the key stuff.
    Writer.WriteValue ReverseByteCopy(mModulus)
    
    If BlobType = PRIVATEKEYBLOB Then
        Writer.WriteValue ReverseByteCopy(mP)
        Writer.WriteValue ReverseByteCopy(mQ)
        Writer.WriteValue ReverseByteCopy(mDP)
        Writer.WriteValue ReverseByteCopy(mDQ)
        Writer.WriteValue ReverseByteCopy(mInverseQ)
        Writer.WriteValue ReverseByteCopy(mD)
    End If
    
    ToCspBlob = ms.ToArray
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBytes(ByVal Elements As SecurityElement, ByRef Name As String) As Byte()
    Dim Text As String
    
    Text = Elements.SearchForTextOfTag(Name)
    If Len(Text) = 0 Then _
        Throw Cor.NewCryptographicException("Invalid Xml format.")
    
    GetBytes = Convert.FromBase64String(Text)
End Function

Private Sub AddElement(ByVal sb As StringBuilder, ByRef Tag As String, ByRef Bytes() As Byte)
    sb.AppendFormat "<{0}>{1}</{0}>", Tag, Convert.ToBase64String(Bytes)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mModulus = .ReadProperty(PropModulus)
        mExponent = .ReadProperty(PropExponent)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropModulus, mModulus
        .WriteProperty PropExponent, mExponent
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
