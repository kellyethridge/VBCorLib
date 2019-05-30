VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinResourceReaderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' WinResourceReaderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub TestIteration()
    Const FileName As String = "c:\windows\system32\dxdiag.exe"
    Dim Reader As WinResourceReader
    Dim Resources As New Hashtable
    
    If Not File.Exists(FileName) Then
        Assert.Ignore "Skipping because file does not exist."
    End If

    Set Reader = Cor.NewWinResourceReader(FileName)
     
    Dim Entry As DictionaryEntry
    For Each Entry In Reader
        Resources.Add Entry.Key, Entry.Value
    Next Entry
    
    Dim Expected As ResourceKey
    Set Expected = Cor.NewResourceKey(7000, ResourceTypes.GroupIcon, 1033)
    Assert.That Resources.ContainsKey(Expected), Iz.True
End Sub

