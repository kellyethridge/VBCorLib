VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDriveInfo
'
Option Explicit

Private di As DriveInfo


Public Sub TestFixtureSetup()
    WriteLine "***** DriveInfo Class *****"
End Sub

Public Sub Setup()
    Set di = NewDriveInfo("c")
End Sub

Public Sub TestGetDrives()
    Dim di() As DriveInfo
    Dim i As Long
    
    di = DriveInfo.GetDrives
    Assert.IsTrue cArray.GetLength(di) > 0, "Should have atleast one drive."
End Sub

Public Sub TestDriveInfoOnC()
    Assert.AreEqual "c:\", di.Name
End Sub

Public Sub TestDriveFormat()
    WriteLine "Drive Format: " & di.DriveFormat
End Sub

Public Sub TestVolumeLabel()
    WriteLine "Volume Label: " & di.VolumeLabel
End Sub

Public Sub TestDriveType()
    Assert.AreEqual DriveType.Fixed, di.DriveType
End Sub

Public Sub TestIsReady()
    Assert.IsTrue di.IsReady
End Sub

Public Sub TestSerialNumber()
    WriteLine "Serial Number: " & di.SerialNumber
End Sub

Public Sub TestAvailableFreeSpace()
    WriteLine "Available Free Space: " & di.AvailableFreeSpace
End Sub

Public Sub TestTotalFreeSpace()
    WriteLine "Total Free Space: " & di.TotalFreeSpace
End Sub

Public Sub TestTotalSize()
    WriteLine "Total Size: " & di.TotalSize
End Sub

