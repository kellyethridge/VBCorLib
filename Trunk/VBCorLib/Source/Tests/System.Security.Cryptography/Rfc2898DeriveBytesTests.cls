VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rfc2898DeriveBytesTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Rfc2898DeriveBytesTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithValue_ReturnsExpected"
            Test.Use 8, NewBytes(&H13, &HD1, &HC9, &HA3, &HFA, &H2E, &H93, &H5)
            Test.Use 13, NewBytes(&H13, &HD1, &HC9, &HA3, &HFA, &H2E, &H93, &H5, &H22, &HDD, &HA4, &HD, &HEF)
        
        Case "GetBytes_WhenResetCalled_RestartsGeneration"
            Test.Use 8, NewBytes(&H13, &HD1, &HC9, &HA3, &HFA, &H2E, &H93, &H5)
            Test.Use 13, NewBytes(&H13, &HD1, &HC9, &HA3, &HFA, &H2E, &H93, &H5, &H22, &HDD, &HA4, &HD, &HEF)
        
        Case "GetBytes_WithMultipleCalls_ReturnsContinuedGeneration"
            Test.Use 8, 13, NewBytes(&H13, &HD1, &HC9, &HA3, &HFA, &H2E, &H93, &H5), NewBytes(&H22, &HDD, &HA4, &HD, &HEF, &HD6, &HC0, &HE7, &HFB, &H5B, &H1E, &H70, &HA1)
            Test.Use 13, 8, NewBytes(&H13, &HD1, &HC9, &HA3, &HFA, &H2E, &H93, &H5, &H22, &HDD, &HA4, &HD, &HEF), NewBytes(&HD6, &HC0, &HE7, &HFB, &H5B, &H1E, &H70, &HA1)
            
    End Select
End Sub

Public Sub IterationCount_WhenSetLessThanOne_ThrowsArgumentOutOfRangeException()
    Dim Rfc As Rfc2898DeriveBytes
    Set Rfc = MakeRfc
    
    On Error Resume Next
    Rfc.IterationCount = 0
    
    AssertArgumentOutOfRangeException Err, "Value"
End Sub

Public Sub Salt_WhenSetWithNull_ThrowsArgumentNullException()
    Dim Rfc As Rfc2898DeriveBytes
    Set Rfc = MakeRfc
    
    On Error Resume Next
    Rfc.Salt = NullBytes
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub Salt_WhenSetLessThan8Bytes_ThrowsArgumentException()
    Dim Rfc As Rfc2898DeriveBytes
    Set Rfc = MakeRfc
    
    On Error Resume Next
    Rfc.Salt = NewBytes(1, 2, 3, 4, 5, 6, 7)
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub Constructor_WithNonStringOrByteArrayPassword_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewRfc2898DeriveBytes 1234, 8
    
    AssertArgumentException Err, "Password"
End Sub

Public Sub Constructor_WithNotNumberOrByteArraySalt_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewRfc2898DeriveBytes "abc", "def"
    
    AssertArgumentException Err, "Salt"
End Sub

Public Sub Constructor_WithSaltSizeLessThanZero_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewRfc2898DeriveBytes "abc", -1
    
    AssertArgumentOutOfRangeException Err, "Salt"
End Sub

Public Sub GetBytes_WithValue_ReturnsExpected(ByVal Size As Long, ByRef Expected As Variant)
    Dim Rfc As Rfc2898DeriveBytes
    Set Rfc = MakeRfc
    
    Dim Actual() As Byte
    Actual = Rfc.GetBytes(Size)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WhenResetCalled_RestartsGeneration(ByVal Size As Long, ByRef Expected As Variant)
    Dim Rfc As Rfc2898DeriveBytes
    Set Rfc = MakeRfc
    Rfc.GetBytes Size
    Rfc.Reset
    
    Dim Actual() As Byte
    Actual = Rfc.GetBytes(Size)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetBytes_WithMultipleCalls_ReturnsContinuedGeneration(ByVal Size1 As Long, ByVal Size2 As Long, ByRef Expected1 As Variant, ByRef Expected2 As Variant)
    Dim Rfc As Rfc2898DeriveBytes
    Set Rfc = MakeRfc
    
    Dim Actual1() As Byte
    Dim Actual2() As Byte
    Actual1 = Rfc.GetBytes(Size1)
    Actual2 = Rfc.GetBytes(Size2)
    
    Assert.That Actual1, Iz.EqualTo(Expected1)
    Assert.That Actual2, Iz.EqualTo(Expected2)
End Sub

Private Function MakeRfc() As Rfc2898DeriveBytes
    Set MakeRfc = Cor.NewRfc2898DeriveBytes("abc", NewBytes(1, 2, 3, 4, 5, 6, 7, 8), 10)
End Function


