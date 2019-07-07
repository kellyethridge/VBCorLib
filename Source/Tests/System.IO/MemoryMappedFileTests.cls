VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedFileTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MemoryMappedFileTests
'
Option Explicit
Implements ICategorizable

Private mIO As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    Set mIO = New IOHelper
End Sub

Public Sub Teardown()
    Set mIO = Nothing
End Sub

Public Sub BaseAddress_WithOpenFile_IsNotZero()
    Dim FileName As String
    FileName = mIO.CreateFile(Data:="AB")
    Dim Map As MemoryMappedFile
    
    Set Map = Cor.NewMemoryMappedFile(FileName)
    
    Assert.That Map.IsOpen, Iz.True
    Assert.That Map.BaseAddress, Iz.Not.EqualTo(0)
End Sub

Public Sub BaseAddress_WithOpenFile_CanAccessContent()
    Dim FileName As String
    FileName = mIO.CreateFile(Data:="AB")
    Dim Map As MemoryMappedFile
    
    Set Map = Cor.NewMemoryMappedFile(FileName)
    
    Assert.That MemByte(Map.BaseAddress), Iz.EqualTo(AscW("A"))
    Assert.That MemByte(Map.BaseAddress + 1), Iz.EqualTo(AscW("B"))
End Sub

Public Sub CreateView_WhenCalled_ReturnsFileContentAsByteArray()
    Dim FileName As String
    FileName = mIO.CreateFile(Data:="AB")
    Dim Map As MemoryMappedFile
    Set Map = Cor.NewMemoryMappedFile(FileName)

    Dim View()      As Byte
    Dim Actual()    As Byte
    View = Map.CreateView
    Actual = View
    Map.DeleteView View
    
    Assert.That CorArray.IsNull(Actual), Iz.False
    Assert.That Actual, Iz.EqualTo(NewBytes(AscW("A"), AscW("B")))
End Sub

