VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResourceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestResourceManager
'
Option Explicit

Private mRm As ResourceManager



Public Sub Setup()
    Set mRm = Cor.NewResourceManager(Path.Combine(App.Path, "VBCorLibTests.res"))
End Sub

Public Sub Teardown()
    mRm.ReleaseAllResources
End Sub


Public Sub TestGetString1033()
    Assert.AreEqual "hello", mRm.GetString(101, Cor.NewCultureInfo("en-US"))
End Sub

Public Sub TestGetString1096()
    Assert.AreEqual "aloha", mRm.GetString(101, 1096)
End Sub
