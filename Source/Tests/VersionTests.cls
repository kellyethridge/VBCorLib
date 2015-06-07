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

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub DefaultConstructor_WhenCalled_SetsVersionToZero()
    Dim SUT As New Version
    
    Assert.That SUT.Major, Iz.EqualTo(0)
    Assert.That SUT.Minor, Iz.EqualTo(0)
    Assert.That SUT.Revision, Iz.EqualTo(-1)
    Assert.That SUT.Build, Iz.EqualTo(-1)
End Sub

Public Sub Constructor_WithComponentVersions_SetsMajorComponent()
    Dim SUT As Version
    
    Set SUT = Cor.NewVersion(1, 2, 3, 4)
    
    Assert.That SUT.Major, Iz.EqualTo(1)
    Assert.That SUT.Minor, Iz.EqualTo(2)
    Assert.That SUT.Build, Iz.EqualTo(3)
    Assert.That SUT.Revision, Iz.EqualTo(4)
End Sub

'Public Sub TestZeroVersion()
'    Dim v As New Version
'    Assert.AreEqual "0.0", v.ToString, "ToString should return 0.0"
'End Sub
'
'Public Sub TestOnePtTwo()
'    Dim v As Version
'    Set v = newversion(1, 2)
'    Assert.AreEqual "1.2", v.ToString, "ToString should return 1.2"
'End Sub
'
'Public Sub TestOnePtTwoPtThree()
'    Dim v As Version
'    Set v = newversion(1, 2, 3)
'    Assert.AreEqual "1.2.3", v.ToString
'End Sub
'
'Public Sub TestOnePtTwoPtThreeptFour()
'    Dim v As Version
'    Set v = newversion(1, 2, 3, 4)
'    Assert.AreEqual "1.2.3.4", v.ToString
'End Sub
'
'Public Sub TestToStringFieldCount()
'    Dim v As Version
'    Set v = newversion(1, 2, 3, 4)
'    Assert.AreEqual "", v.ToString(0)
'    Assert.AreEqual "1", v.ToString(1)
'    Assert.AreEqual "1.2", v.ToString(2)
'    Assert.AreEqual "1.2.3", v.ToString(3)
'    Assert.AreEqual "1.2.3.4", v.ToString(4)
'End Sub
'
'Public Sub TestClone()
'    Dim v As Version
'    Set v = newversion(1, 2, 3, 4)
'
'    Dim v2 As Version
'    Set v2 = v.Clone
'    Assert.AreEqual "1.2.3.4", v2.ToString
'End Sub
'
'Public Sub TestCompareTo()
'    Dim v As Version
'    Set v = newversion(1, 2, 3, 4)
'    Assert.AreEqual 0, v.CompareTo(newversion(1, 2, 3, 4)), "CompareTo should return zero."
'    Assert.AreEqual -1, v.CompareTo(newversion(1, 2, 4)), "CompareTo should return negative value."
'    Assert.AreEqual 1, v.CompareTo(newversion(1, 2, 2)), "CompareTo should return positive value."
'    Assert.AreEqual 1, v.CompareTo(Nothing), "CompareTo should return positive value."
'End Sub
'
'Public Sub TestEquals()
'    Dim v As Version
'    Set v = newversion(1, 2)
'    Assert.IsTrue v.Equals(newversion(1, 2)), "Equals should return true."
'End Sub
'
'Public Sub TestGetHashCode()
'    Dim v As Version
'    Set v = newversion(1, 2, 3, 4)
'    Assert.AreEqual 16909060, v.GetHashCode, "Wrong hash code."
'End Sub
'
'Public Sub TestPersistence()
'    Dim v1 As Version
'    Set v1 = newversion(1, 2, 3, 4)
'
'    Dim pp1 As New PropertyBag
'    Dim pp2 As New PropertyBag
'    Dim v2 As Version
'
'    pp1.WriteProperty "Version", v1
'    pp2.Contents = pp1.Contents
'    Set v2 = pp2.ReadProperty("Version")
'
'    Assert.IsTrue v2.Equals(v1), "V2 should equal V1."
'End Sub
'
'Public Sub TestParse()
'    Dim v As Version
'
'    Set v = Version.Parse("1.2.3.4")
'    Assert.IsTrue v.Equals(newversion(1, 2, 3, 4)), "Versions should be equal."
'End Sub
