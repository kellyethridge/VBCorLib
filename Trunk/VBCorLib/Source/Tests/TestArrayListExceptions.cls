VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayListExceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Argument As String = "ArgumentException"
Private Const ArgumentOutOfRange As String = "ArgumentOutOfRangeException"

Private list As ArrayList
Private ex As Exception

Public Sub Setup()
    Set list = New ArrayList
    list.Add 1
End Sub

Public Sub TestBinarySearch()
    On Error Resume Next
    list.BinarySearch 1, -1, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.BinarySearch 1, 0, 2
    On Error GoTo 0
    AssertException Argument
End Sub

Public Sub TestCapacity()
    On Error Resume Next
    list.capacity = 0
    On Error GoTo 0
    AssertException ArgumentOutOfRange
End Sub

Public Sub TestGetEnumerator()
    On Error Resume Next
    list.GetEnumerator -1, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.GetEnumerator 0, 2
    On Error GoTo 0
    AssertException Argument
End Sub

Public Sub TestGetRange()
    On Error Resume Next
    list.GetRange -1, 0
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.GetRange 0, -1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.GetRange 0, 2
    On Error GoTo 0
    AssertException Argument
End Sub

Public Sub TestIndexOf()
    On Error Resume Next
    list.IndexOf 1, -1, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.IndexOf 0, 2, 1
    On Error GoTo 0
    AssertException Argument
End Sub

Public Sub TestInsert()
    On Error Resume Next
    list.Insert -1, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.Insert 2, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
End Sub

Public Sub TestInsertRange()
    On Error Resume Next
    list.Insertrange -1, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.Insertrange 2, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
End Sub

Public Sub TestLastIndexOf()
    On Error Resume Next
    list.LastIndexOf 1, 1, 1
    On Error GoTo 0
    AssertException ArgumentOutOfRange
    On Error Resume Next
    list.LastIndexOf 1, 0, 2
    On Error GoTo 0
    AssertException Argument
End Sub

' can't catch this error with vbcorlib uncompiled.
'
'Public Sub TestSerialize()
'    Dim pp As PropertyBag
'    Dim c As Collection
'    Set pp = New PropertyBag
'    Set c = New Collection
'    list.Add c
'    On Error Resume Next
'    pp.WriteProperty "list", list
'    On Error GoTo 0
'    AssertException "serializationexception"
'End Sub






Private Sub AssertException(ByVal exName As String)
    Assert.IsTrue Catch(ex), "Exception should be caught."
    Assert.IsTrue LCase$(TypeName(ex)) = LCase$(exName), "Wrong exception caught. Caught " & TypeName(ex) & " instead of " & exName & "."
End Sub
