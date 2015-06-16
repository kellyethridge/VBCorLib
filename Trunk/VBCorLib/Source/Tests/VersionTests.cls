VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithValue_ReturnsExpected"
            Test.Use Cor.NewVersion(1, 2, 3, 4), True
            Test.Use Cor.NewVersion(1, 2, 3), False
            Test.Use Cor.NewVersion(1, 2), False
            Test.Use Cor.NewVersion(1, 2, 3, 5), False
            Test.Use Cor.NewVersion(1, 2, 4, 4), False
            Test.Use Cor.NewVersion(1, 3, 3, 4), False
            Test.Use Cor.NewVersion(2, 2, 3, 4), False
            Test.Use Nothing, False
            Test.Use Empty, False
            Test.Use Null, False
            Test.Use "abc", False
            
        Case "CompareTo_WithValue_ReturnsExpected"
            Test.Use Cor.NewVersion(1, 1, 1, 1), 0
            Test.Use Cor.NewVersion(1, 1, 1, 2), -1
            Test.Use Cor.NewVersion(1, 1, 1, 0), 1
            Test.Use Cor.NewVersion(1, 1, 2, 1), -1
            Test.Use Cor.NewVersion(1, 1, 0, 1), 1
            Test.Use Cor.NewVersion(1, 2, 1, 1), -1
            Test.Use Cor.NewVersion(1, 0, 1, 1), 1
            Test.Use Cor.NewVersion(2, 1, 1, 1), -1
            Test.Use Cor.NewVersion(0, 1, 1, 1), 1
            Test.Use Nothing, 1
        
        Case "CompareTo_WithNoBuildOrRevision_ReturnsExpected"
            Test.Use Cor.NewVersion(1, 1), 0
            Test.Use Cor.NewVersion(1, 1, 0), -1
        
        Case "CompareTo_WithNoRevision_ReturnsExpected"
            Test.Use Cor.NewVersion(1, 1, 1), 0
            Test.Use Cor.NewVersion(1, 1, 1, 0), -1
            
        Case "ToString_WhenCalled_ReturnsExpected"
            Test.Use New Version, "0.0"
            Test.Use Cor.NewVersion(1, 2), "1.2"
            Test.Use Cor.NewVersion(1, 2, 3), "1.2.3"
            Test.Use Cor.NewVersion(1, 2, 3, 4), "1.2.3.4"
            Test.Use Cor.NewVersion(0, 0, 0, 0), "0.0.0.0"
            
        Case "ToString_WithFieldCount_ReturnsExpected"
            Test.Use 0, ""
            Test.Use 1, "1"
            Test.Use 2, "1.2"
            Test.Use 3, "1.2.3"
            Test.Use 4, "1.2.3.4"
            
    End Select
End Sub

Public Sub Constructor_WhenCalled_SetsVersionToZero()
    Dim SUT As New Version
    
    Assert.That SUT.Major, Iz.EqualTo(0)
    Assert.That SUT.Minor, Iz.EqualTo(0)
    Assert.That SUT.Revision, Iz.EqualTo(-1)
    Assert.That SUT.Build, Iz.EqualTo(-1)
End Sub

Public Sub NewVersion_WithComponentVersions_SetsMajorComponent()
    Dim SUT As Version
    
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Assert.That SUT.Major, Iz.EqualTo(1)
    Assert.That SUT.Minor, Iz.EqualTo(2)
    Assert.That SUT.Build, Iz.EqualTo(3)
    Assert.That SUT.Revision, Iz.EqualTo(4)
End Sub

Public Sub NewVersion_WithNegativeMajor_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewVersion -1, 0
    
    AssertArgumentOutOfRangeException Err, "Major"
End Sub

Public Sub NewVersion_WithNegativeMinor_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewVersion 0, -1
    
    AssertArgumentOutOfRangeException Err, "Minor"
End Sub

Public Sub NewVersion_WithNegativeBuild_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewVersion 0, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Build"
End Sub

Public Sub NewVersion_WithNegativeRevision_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewVersion 0, 0, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Revision"
End Sub

Public Sub NewVersion_WithRevisionAndNoBuild_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewVersion 0, 0, Revision:=1
    
    AssertArgumentException Err, "Build"
End Sub

Public Sub Equals_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As Boolean)
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Dim Actual As Boolean
    Actual = SUT.Equals(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CompareTo_WithValue_ReturnsExpected(ByVal Value As Version, ByVal Expected As Long)
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 1, 1, 1)
    
    Dim Actual As Long
    Actual = SUT.CompareTo(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CompareTo_WithNoBuildOrRevision_ReturnsExpected(ByVal Version As Version, ByVal Expected As Long)
     Dim SUT As Version
     Set SUT = Cor.NewVersion(1, 1)
     
     Dim Actual As Long
     Actual = SUT.CompareTo(Version)
     
     Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CompareTo_WithNoRevision_ReturnsExpected(ByVal Version As Version, ByVal Expected As Long)
     Dim SUT As Version
     Set SUT = Cor.NewVersion(1, 1, 1)
     
     Dim Actual As Long
     Actual = SUT.CompareTo(Version)
     
     Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetHashCode_WhenCalled_ReturnsExpected()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Dim Actual As Long
    Actual = SUT.GetHashCode
    
    Const Expected As Long = 16909060
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WhenCalled_ReturnsExpected(ByVal Version As Version, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Version.ToString
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFieldCount_ReturnsExpected(ByVal FieldCount As Long, ByVal Expected As String)
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Dim Actual As String
    Actual = SUT.ToString(FieldCount)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFieldCountNegative_ThrowsArgumentException()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2)
    
    On Error Resume Next
    SUT.ToString -1
    
    AssertArgumentException Err, "FieldCount"
End Sub

Public Sub ToString_WithFieldCount3AndNotBuildDefined_ThrowsArgumentException()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2)
    
    On Error Resume Next
    SUT.ToString 3
    
    AssertArgumentException Err, "FieldCount"
End Sub

Public Sub ToString_WithFieldCount4AndNotRevisionDefined_ThrowsArgumentException()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3)
    
    On Error Resume Next
    SUT.ToString 4
    
    AssertArgumentException Err, "FieldCount"
End Sub

Public Sub ToString_WithFieldCount5_ThrowsArgumentException()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    On Error Resume Next
    SUT.ToString 5
    
    AssertArgumentException Err, "FieldCount"
End Sub

Public Sub Clone_WhenCalled_ReturnsCopy()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Dim Actual As Version
    Set Actual = SUT.Clone
    
    Assert.That Actual, Iz.Not.SameAs(SUT)
    Assert.That Actual, Equals(SUT)
End Sub

Public Sub Serialization_WhenCalled_SerializesInstance()
    Dim SUT As Version
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Dim Actual As Version
    Set Actual = Persist(SUT)
    
    Assert.That Actual, Iz.Not.SameAs(SUT)
    Assert.That Actual, Equals(SUT)
End Sub

'Public Sub TestParse()
'    Dim v As Version
'
'    Set v = Version.Parse("1.2.3.4")
'    Assert.IsTrue v.Equals(newversion(1, 2, 3, 4)), "Versions should be equal."
'End Sub

