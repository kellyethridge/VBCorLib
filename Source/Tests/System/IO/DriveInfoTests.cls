VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DriveInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' DriveInfoTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "NewDriveInfo_WithDrive_ConvertsToRoot"
            Test.Use "c", "c:\"
            Test.Use "c:\", "c:\"
    End Select
End Sub


Public Sub GetDrives_WhenCalled_ReturnsNonEmptyList()
    Dim Drives() As DriveInfo
    
    Drives = DriveInfo.GetDrives
    
    Assert.That Drives, Iz.Not.Empty
End Sub

Public Sub NewDriveInfo_WithDrive_ConvertsToRoot(ByVal Drive As String, ByVal Expected As String)
    Dim Info As DriveInfo
    
    Set Info = NewDriveInfo(Drive)
    
    Assert.That Info.Name, Iz.EqualTo(Expected).IgnoreCase
End Sub

Public Sub DriveFormat_WhenCalled_ReturnsNTFS()
    Dim Info As DriveInfo
    Set Info = NewDriveInfo("c")
    Dim Actual As String
    
    Actual = Info.DriveFormat
    
    Assert.That Actual, Iz.EqualTo("NTFS")
End Sub

Public Sub DriveType_WhenCalled_ReturnsFixed()
    Dim Info As DriveInfo
    Set Info = NewDriveInfo("c")
    Dim Actual As DriveType
    
    Actual = Info.DriveType
    
    Assert.That Actual, Iz.EqualTo(DriveType.Fixed)
End Sub

Public Sub IsReady_WhenCalled_ReturnsTrue()
    Dim Info As DriveInfo
    Set Info = NewDriveInfo("c")
    Dim Actual As Boolean
    
    Actual = Info.IsReady
    
    Assert.That Actual, Iz.True
End Sub

Public Sub AvailableFreeSpace_WhenCalled_ReturnsGreaterThanZero()
    Dim Info As DriveInfo
    Set Info = NewDriveInfo("c")
    Dim Actual As Currency
    
    Actual = Info.AvailableFreeSpace
    
    Assert.That Actual, Iz.GreaterThan(0)
End Sub

Public Sub TotalFreeSpace_WhenCalled_ReturnsGreaterThanZero()
    Dim Info As DriveInfo
    Set Info = NewDriveInfo("c")
    Dim Actual As Currency
    
    Actual = Info.TotalFreeSpace
    
    Assert.That Actual, Iz.GreaterThan(0)
End Sub

Public Sub TotalSize_WhenCalled_ReturnsGreaterThanZero()
    Dim Info As DriveInfo
    Set Info = NewDriveInfo("c")
    Dim Actual As Currency
    
    Actual = Info.TotalSize
    
    Assert.That Actual, Iz.GreaterThan(0)
End Sub


