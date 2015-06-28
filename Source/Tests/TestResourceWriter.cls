VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResourceWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestResourceWriter
'
Option Explicit
Implements ICategorizable

Private mWriter As ResourceWriter
Private mStream As Stream

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Setup()
    Set mStream = New MemoryStream
    Set mWriter = Cor.NewResourceWriter(mStream)
End Sub

Public Sub TestWrite1ByteEntry()
    Dim Data() As Byte
    Data = NewBytes(25)
    
    mWriter.AddResource Data, 101, "Custom"
    mWriter.Generate
    
    ' 32 bytes for empty resource entry
    ' 24 bytes for the minimum header size.
    ' 4 bytes for the resource ID ((101 * 16) or &HFFFF).
    ' 14 bytes for the resource type "Custom" (2 bytes per character + 2 byte terminator).
    ' 2 bytes for DWord alignment.
    ' 1 byte of data.
    ' 3 bytes for DWord alignment.
    Assert.AreEqual 80@, mStream.Length
    
    mStream.Position = 0
    Dim Resources As ResourceSet
    Set Resources = Cor.NewResourceSet(1033, mStream)
    
    Dim b As Variant
    b = Resources.GetObject(101, "Custom")
    Assert.IsTrue IsArray(b), "Should return an array."
    Assert.AreEqual 1, CorArray.Length(b)
    Assert.AreEqual 25, b(0)
End Sub

Public Sub TestWrite1StringEntry()
    mWriter.AddResource "hello", 101
    mWriter.Generate

    ' 32 bytes for a blank resource entry.
    '
    ' 24 bytes for the minimum header size.
    ' 4 bytes for the resource ID (((101 \ 16) + 1) * 16) or &HFFFF.
    ' 4 bytes for the resource type &H0006FFFF.
    ' 0 bytes for DWord alignment.
    ' 42 bytes of data.
    ' 2 bytes for DWord alignment.
    Assert.AreEqual 108@, mStream.Length

    mStream.Position = 0
    Dim Resources As ResourceSet
    Set Resources = Cor.NewResourceSet(1033, mStream)
    
    Assert.AreEqual "hello", Resources.GetString(101)
End Sub

Public Sub TestWrite3String1ByteEntry()
    ' 32 bytes for a blank resource entry.
    '
    ' Both of these strings end up in the
    ' same string resource and share the header.
    '
    ' 24 bytes for minimum header size.
    ' 4 bytes for resource ID (((101 \ 16) + 1) * 16) or &HFFFF.
    ' 4 bytes for resource type &H0006FFFF.
    ' 0 bytes for DWord alignment.
    ' 48 bytes of data.
    ' 0 bytes for DWord alignment.
    mWriter.AddResource "hello", 101
    mWriter.AddResource "bye", 102
    
    ' 24 bytes for minimum header size.
    ' 4 bytes for resource ID (((500 \ 16) + 1) * 16) or &HFFFF.
    ' 4 bytes for resource type &H0006FFFF.
    ' 0 bytes for DWord alignment.
    ' 48 bytes of data.
    ' 0 bytes for DWord alignment.
    mWriter.AddResource "Far Away", 500
    
    Dim Data(0) As Byte
    Data(0) = 25
    
    ' 24 bytes for minimum header size.
    ' 4 bytes for resource ID ((101 * 16) or &HFFFF)
    ' 14 bytes for the resource type "Custom" (2 bytes per character + 2 byte terminator).
    ' 2 bytes for DWord alignment.
    ' 1 byte of data.
    ' 3 bytes for DWord alignment.
    mWriter.AddResource Data, 101, "Custom"
    
    mWriter.Generate
    
    Assert.AreEqual 240@, mStream.Length
    
    mStream.Position = 0
    Dim Resources As ResourceSet
    Set Resources = Cor.NewResourceSet(1033, mStream)
    
    Assert.AreEqual "hello", Resources.GetString(101)
    Assert.AreEqual "bye", Resources.GetString(102)
    Assert.AreEqual "Far Away", Resources.GetString(500)
    
    Dim b As Variant
    b = Resources.GetObject(101, "Custom")
    Assert.IsTrue IsArray(b), "Should return an array."
    Assert.AreEqual 1, CorArray.Length(b)
    Assert.AreEqual 25, b(0)
End Sub
