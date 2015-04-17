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
'    Module: RSAParameters
'

''
' Represents the Key information for an RSA implementation.
'
' @remarks To learn more about RSAParameters, please visit this Shawnfa blog entry:
' http://blogs.msdn.com/shawnfa/archive/2005/11/17/493972.aspx
'
Option Explicit
Implements IObject

Private Const PROP_MODULUS  As String = "Modulus"
Private Const PROP_EXPONENT As String = "Exponent"
Private Const PROP_INVERSEQ As String = "InverseQ"

Private md()        As Byte ' PrivateExponent
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
    d = md
End Property

''
' Sets the D parameter of the RSA key.
'
' @param RHS The new D parameter value.
' @remarks This is the Private Exponent of the RSA key.
'
Public Property Let d(ByRef RHS() As Byte)
    md = RHS
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
' @param RHS The new DP value.
' @remarks This is the Exponent1 value of the key. It has the numeric value "d mod (p - 1)".
'
Public Property Let DP(ByRef RHS() As Byte)
    mDP = RHS
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
' @param RHS The new DQ value.
' @remarks This is the Exponent2 value of the key. It has the numeric value "d mod (q - 1)".
'
Public Property Let DQ(ByRef RHS() As Byte)
    mDQ = RHS
End Property

''
' Returns the public exponent of the key.
'
' @return A byte array of the public exponent.
'
Public Property Get exponent() As Byte()
    exponent = mExponent()
End Property

''
' Sets the public exponent of the key.
'
' @param RHS The new exponent value.
'
Public Property Let exponent(ByRef RHS() As Byte)
    mExponent = RHS
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
' @param RHS The new InverseQ value.
' @remarks This is the Coefficient. It has a numeric value "(inverse of q) mod p".
'
Public Property Let InverseQ(ByRef RHS() As Byte)
    mInverseQ = RHS
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
' @param RHS The new modulus value.
' @remarks This has a value of "prime1 * prime2" and is often known as "n".
'
Public Property Let Modulus(ByRef RHS() As Byte)
    mModulus = RHS
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
' @param RHS The new P value.
' @remarks This is the Prime1 value of the key.
'
Public Property Let p(ByRef RHS() As Byte)
    mP = RHS
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
' @param RHS The new Q value.
' @remarks This is the Prime2 value of the key.
'
Public Property Let q(ByRef RHS() As Byte)
    mQ = RHS
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



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function ToXmlString() As String
    Dim sb As New StringBuilder
    Call sb.AppendString("<RSAKeyValue>")
    Call AddElement(sb, "Modulus", mModulus)
    Call AddElement(sb, "Exponent", mExponent)
    
    If Not cArray.IsNull(md) Then
        Call AddElement(sb, "P", mP)
        Call AddElement(sb, "Q", mQ)
        Call AddElement(sb, "DP", mDP)
        Call AddElement(sb, "DQ", mDQ)
        Call AddElement(sb, "InverseQ", mInverseQ)
        Call AddElement(sb, "D", md)
    End If
    
    Call sb.AppendString("</RSAKeyValue>")
    
    ToXmlString = sb.ToString
End Function

Friend Sub FromXmlString(ByVal Xml As String)
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
        md = GetBytes(Elements, "D")
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
    Call cArray.Reverse(mExponent)
    
    'Call Reader.ReadByte
    
    ' start reading in all the key stuff.
    mModulus = ReverseByteCopy(Reader.ReadBytes(BitLen \ 8))
    
    If BlobType = PRIVATEKEYBLOB Then
        mP = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mQ = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mDP = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mDQ = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        mInverseQ = ReverseByteCopy(Reader.ReadBytes(BitLen \ 16))
        md = ReverseByteCopy(Reader.ReadBytes(BitLen \ 8))
    End If
    
    Call Reader.CloseReader
End Sub

Friend Function ToCspBlob(ByVal KeyNumber As Long) As Byte()
    Dim ms As New MemoryStream
    Dim writer As BinaryWriter
    Set writer = Cor.NewBinaryWriter(ms)
    
    Dim BlobType As Byte
    BlobType = IIf(cArray.IsNull(md), PUBLICKEYBLOB, PRIVATEKEYBLOB)
    
    ' Build BLOBHEADER structure
    Call writer.WriteValue(BlobType)
    Call writer.WriteValue(CByte(2))
    Call writer.WriteValue(CInt(0))
    Call writer.WriteValue(CLng(IIf(KeyNumber = AT_KEYEXCHANGE, CALG_RSA_KEYX, CALG_RSA_SIGN)))
    
    ' Build RSAPUBKEY structure
    Call writer.WriteValue(CLng(IIf(BlobType = PUBLICKEYBLOB, &H31415352, &H32415352)))
    Call writer.WriteValue(CLng(cArray.GetLength(mModulus) * 8))
    Call writer.WriteValue(ReverseByteCopy(mExponent))
    
    Dim i As Long
    i = 4 - cArray.GetLength(mExponent)
    Do While i > 0
        Call writer.WriteValue(CByte(0))
        i = i - 1
    Loop
    
    ' Start adding all the key stuff.
    Call writer.WriteValue(ReverseByteCopy(mModulus))
    
    If BlobType = PRIVATEKEYBLOB Then
        Call writer.WriteValue(ReverseByteCopy(mP))
        Call writer.WriteValue(ReverseByteCopy(mQ))
        Call writer.WriteValue(ReverseByteCopy(mDP))
        Call writer.WriteValue(ReverseByteCopy(mDQ))
        Call writer.WriteValue(ReverseByteCopy(mInverseQ))
        Call writer.WriteValue(ReverseByteCopy(md))
    End If
    
    ToCspBlob = ms.ToArray
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBytes(ByVal Elements As SecurityElement, ByVal Name As String) As Byte()
    Dim Text As String
    
    Text = Elements.SearchForTextOfTag(Name)
    If Len(Text) = 0 Then _
        Throw Cor.NewCryptographicException("Invalid Xml format.")
    
    GetBytes = Convert.FromBase64String(Text)
End Function

Private Sub AddElement(ByVal sb As StringBuilder, ByVal Tag As String, ByRef Bytes() As Byte)
    Call sb.AppendFormat("<{0}>{1}</{0}>", Tag, Convert.ToBase64String(Bytes))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mModulus = .ReadProperty(PROP_MODULUS)
        mExponent = .ReadProperty(PROP_EXPONENT)
        mInverseQ = .ReadProperty(PROP_INVERSEQ)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_MODULUS, mModulus)
        Call .WriteProperty(PROP_EXPONENT, mExponent)
        Call .WriteProperty(PROP_INVERSEQ, mInverseQ)
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
