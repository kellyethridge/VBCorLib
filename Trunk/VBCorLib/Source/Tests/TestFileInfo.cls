VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestFileInfo
'
Option Explicit
Implements ICategorizable

Private info As FileInfo

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set info = Cor.NewFileInfo(App.Path & "\TestFileInfo.cls")
End Sub

Public Sub TestExists()
    Assert.IsTrue info.Exists
End Sub

Public Sub TestLength()
    Assert.AreEqual CCur(FileLen(App.Path & "\TestFileInfo.cls")), info.Length
End Sub

Public Sub TestName()
    Assert.AreEqual "TestFileInfo.cls", info.Name
End Sub

Public Sub TestParent()
    Assert.AreEqual LCase$(App.Path), LCase$(info.DirectoryName)
End Sub

Public Sub TestExtension()
    Assert.AreEqual ".cls", info.Extension
End Sub

Public Sub TestFullName()
    Assert.AreEqual LCase$(App.Path & "\testfileinfo.cls"), LCase$(info.FullName)
End Sub
