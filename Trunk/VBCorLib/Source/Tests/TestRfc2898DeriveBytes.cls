VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRfc2898DeriveBytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Test2898DeriveBytes
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Public Sub SetSaltValue()
    Dim db As Rfc2898DeriveBytes
    Set db = Cor.NewRfc2898DeriveBytes("hello", Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8))
    Assert.AreEqual Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8), db.Salt
    
    db.Salt = Cor.NewBytes(8, 7, 6, 5, 4, 3, 2, 1)
    Assert.AreEqual Cor.NewBytes(8, 7, 6, 5, 4, 3, 2, 1), db.Salt
End Sub

Public Sub Get8BytesUsingBytePassword()
    Dim pwd() As Byte
    pwd = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Salt() As Byte
    Salt = pwd
    
    Dim db As Rfc2898DeriveBytes
    Set db = Cor.NewRfc2898DeriveBytes(pwd, Salt)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H6D, &H29, &HE5, &HD8, &H7B, &H96, &H5E, &H99)
    
    Assert.AreEqual Expected, db.GetBytes(8)
End Sub

Public Sub Get8BytesTwiceAreDifferent()
    Dim pwd() As Byte
    pwd = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Salt() As Byte
    Salt = pwd
    
    Dim db As Rfc2898DeriveBytes
    Set db = Cor.NewRfc2898DeriveBytes(pwd, Salt)
    
    Dim b1() As Byte
    Dim b2() As Byte
    b1 = db.GetBytes(8)
    b2 = db.GetBytes(8)
    
    Assert.That b1, Iz.Not.EqualTo(b2)
End Sub

Public Sub Get8BytesTwiceWithResetAreSame()
    Dim pwd() As Byte
    pwd = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Salt() As Byte
    Salt = pwd
    
    Dim db As Rfc2898DeriveBytes
    Set db = Cor.NewRfc2898DeriveBytes(pwd, Salt)
    
    Dim b1() As Byte
    Dim b2() As Byte
    b1 = db.GetBytes(8)
    db.Reset
    b2 = db.GetBytes(8)
    
    Assert.AreEqual b1, b2
End Sub

Public Sub Get10BytesTwiceSameAs20BytesOnce()
    Dim pwd() As Byte
    pwd = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim Salt() As Byte
    Salt = pwd
    
    Dim db As Rfc2898DeriveBytes
    Set db = Cor.NewRfc2898DeriveBytes(pwd, Salt)
    
    Dim b1() As Byte
    Dim b2() As Byte
    Dim b3() As Byte
    b1 = db.GetBytes(10)
    b2 = db.GetBytes(10)
    
    db.Reset
    b3 = db.GetBytes(20)
    
    Dim i As Long
    For i = 0 To 9
        Assert.AreEqual b1(i), b3(i), "i = " & i
        Assert.AreEqual b2(i), b3(i + 10), "i = " & i
    Next i
End Sub

Public Sub Get20BytesUsingStringPassword()
    Dim pwd As String
    pwd = String$(10, ChrW$(&H2030))
    
    Dim Salt() As Byte
    Salt = Cor.NewBytes(1, 2, 3, 4, 5, 6, 7, 8)
    
    Dim db As Rfc2898DeriveBytes
    Set db = Cor.NewRfc2898DeriveBytes(pwd, Salt)
    
    Dim Expected() As Byte
    Expected = Cor.NewBytes(&H46, &HE8, &HDD, &H5F, &H9F, &H18, &H66, &HB3, &H82, &HC3, &HA8, &HE0, &H39, &HB8, &H7B, &HB, &HF7, &H2, &H21, &H2E)
    
    Assert.AreEqual Expected, db.GetBytes(20)
End Sub



Private Sub PrintBytes(ByRef Bytes() As Byte)
    Dim i As Long
    For i = LBound(Bytes) To UBound(Bytes)
        Debug.Print cString.Format("{0:X2} ", Bytes(i));
    Next i
    Debug.Print
End Sub
