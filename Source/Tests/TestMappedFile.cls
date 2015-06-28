VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMappedFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMappedFile
'
Option Explicit
Implements ICategorizable

Private Const CHAR_A As Byte = 65
Private Const CHAR_B As Byte = 66

Private File As MemoryMappedFile

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set File = NewMemoryMappedFile(App.Path & "\mappedfile.txt", ReadWriteAccess)
End Sub

Public Sub Teardown()
    File.CloseFile
End Sub

Public Sub TestMapping()
    Assert.IsTrue File.IsOpen, "Mapped file should be open."
    Assert.IsTrue File.BaseAddress <> 0, "Mapped file should have memory set."
End Sub

Public Sub TestGetCultureCount()
    Dim i As Long
    
    i = MemByte(File.BaseAddress)
    Assert.AreEqual CHAR_A, i, "Memory should contain 'A'."
    i = MemByte(File.BaseAddress + 1)
    Assert.AreEqual CHAR_B, i, "Memory should contain 'B'."
End Sub

Public Sub TestByteView()
    Dim b() As Byte
    
    b = File.CreateView
    
    Assert.IsFalse CorArray.IsNull(b), "Array should not be null."
    Assert.AreEqual 2, CorArray.Length(b), "Array should be 2 bytes long."
    Assert.AreEqual CHAR_A, b(0)
    Assert.AreEqual CHAR_B, b(1)
    
    File.DeleteView b
    Assert.IsTrue CorArray.IsNull(b), "Array should be null."
End Sub

