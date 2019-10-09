VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystemTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' OperatingSystemTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "VersionString_WithPlatformID_ReturnsInfo"
            Test.Use PlatformID.Win32NT, "Microsoft Windows NT 1.2.3.4"
            Test.Use PlatformID.Win32S, "Microsoft Win32S 1.2.3.4"
            Test.Use PlatformID.WinCE, "Microsoft Windows CE 1.2.3.4"
            
        Case "VersionString_WithWin32WindowsPlatformAndVersion_ReturnsExpected"
            Test.Use Cor.NewVersion(4, 0), "Microsoft Windows 95 4.0"
            Test.Use Cor.NewVersion(4, 1), "Microsoft Windows 98 4.1"
            Test.Use Cor.NewVersion(5, 0), "Microsoft Windows 98 5.0"
            
    End Select
End Sub

Public Sub NewOperatingSystem_WhenCalled_CreatesNewInstance()
    Dim Actual As OperatingSystem
    
    Set Actual = Cor.NewOperatingSystem(Win32NT, Cor.NewVersion(1, 2, 3, 4), "SP1")
    
    Assert.That Actual.Platform, Iz.EqualTo(PlatformID.Win32NT)
    Assert.That Actual.Version, Equals(Cor.NewVersion(1, 2, 3, 4))
    Assert.That Actual.ServicePack, Iz.EqualTo("SP1")
End Sub

Public Sub VersionString_WithPlatformID_ReturnsInfo(ByVal PlatformID As PlatformID, ByVal Expected As String)
    Dim OS As OperatingSystem
    Set OS = Cor.NewOperatingSystem(PlatformID, Cor.NewVersion(1, 2, 3, 4))
    
    Dim Actual As String
    Actual = OS.VersionString
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub VersionString_WithWin32WindowsPlatformAndVersion_ReturnsExpected(ByVal Version As Version, ByVal Expected As String)
    Dim OS As OperatingSystem
    Set OS = Cor.NewOperatingSystem(Win32Windows, Version)
    
    Dim Actual As String
    Actual = OS.VersionString
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub VersionString_WithServicePack_IncludesServicePack()
    Dim OS As OperatingSystem
    Set OS = Cor.NewOperatingSystem(Win32NT, Cor.NewVersion(6, 2, 9200, 0), "SP1")
    
    Dim Actual As String
    Actual = OS.VersionString
    
    Assert.That Actual, Iz.EqualTo("Microsoft Windows NT 6.2.9200 SP1")
End Sub
