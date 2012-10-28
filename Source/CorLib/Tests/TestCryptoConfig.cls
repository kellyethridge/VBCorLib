VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCryptoConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCryptoConfig
'
Option Explicit

Public Sub TestMapNameToOID()
    Assert.AreEqual "1.3.14.3.2.26", CryptoConfig.MapNameToOID("sha1")
    Assert.AreEqual "1.3.14.3.2.26", CryptoConfig.MapNameToOID("sha")
    Assert.AreEqual "2.16.840.1.101.3.4.2.1", CryptoConfig.MapNameToOID("sha256")
    Assert.AreEqual "2.16.840.1.101.3.4.2.2", CryptoConfig.MapNameToOID("sha384")
    Assert.AreEqual "2.16.840.1.101.3.4.2.3", CryptoConfig.MapNameToOID("sha512")
    Assert.AreEqual "1.2.840.113549.2.5", CryptoConfig.MapNameToOID("md5")
    Assert.AreEqual "1.3.36.3.2.1", CryptoConfig.MapNameToOID("ripemd160")
    Assert.AreEqual "", CryptoConfig.MapNameToOID("unknown")
End Sub

Public Sub TestEncodeOID()
    Dim SHAOID() As Byte
    SHAOID = Cor.NewBytes(&H6, &H5, &H2B, &HE, &H3, &H2, &H1A)
    Assert.AreEqual SHAOID, CryptoConfig.EncodeOID(CryptoConfig.MapNameToOID("SHA"))
    
    Dim SHA256OID() As Byte
    SHA256OID = Cor.NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H1)
    Assert.AreEqual SHA256OID, CryptoConfig.EncodeOID(CryptoConfig.MapNameToOID("SHA256"))
    
    Dim SHA384OID() As Byte
    SHA384OID = Cor.NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H2)
    Assert.AreEqual SHA384OID, CryptoConfig.EncodeOID(CryptoConfig.MapNameToOID("SHA384"))

    Dim SHA512OID() As Byte
    SHA512OID = Cor.NewBytes(&H6, &H9, &H60, &H86, &H48, &H1, &H65, &H3, &H4, &H2, &H3)
    Assert.AreEqual SHA512OID, CryptoConfig.EncodeOID(CryptoConfig.MapNameToOID("SHA512"))

    Dim RIPEMD160OID() As Byte
    RIPEMD160OID = Cor.NewBytes(&H6, &H5, &H2B, &H24, &H3, &H2, &H1)
    Assert.AreEqual RIPEMD160OID, CryptoConfig.EncodeOID(CryptoConfig.MapNameToOID("RIPEMD160"))

    Dim MD5OID() As Byte
    MD5OID = Cor.NewBytes(&H6, &H8, &H2A, &H86, &H48, &H86, &HF7, &HD, &H2, &H5)
    Assert.AreEqual MD5OID, CryptoConfig.EncodeOID(CryptoConfig.MapNameToOID("MD5"))

End Sub
