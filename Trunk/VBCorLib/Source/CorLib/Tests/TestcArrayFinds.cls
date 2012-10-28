VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestcArrayFinds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestcArrayFinds
'
Option Explicit

Public Sub TestExists()
    Dim a(-100 To -5) As String
    Dim i As Long
    
    For i = LBound(a) To UBound(a)
        a(i) = "Item " & i
    Next i
    
    a(UBound(a) - 20) = "Hello"
    Assert.IsTrue cArray.Exists(a, AddressOf FindHelloString)
    
    a(UBound(a) - 20) = "Bye"
    Assert.IsFalse cArray.Exists(a, AddressOf FindHelloString)
End Sub

Public Sub TestFind()
    Dim a(-50 To 50) As String
    Dim i As Long
    
    For i = LBound(a) To UBound(a)
        a(i) = "Item " & i
    Next i
    
    a(30) = "Hello"
    Assert.AreEqual "Hello", cArray.find(a, AddressOf FindHelloString)
    
    a(30) = "Bye"
    Assert.AreEqual "", cArray.find(a, AddressOf FindHelloString)
End Sub

Public Sub TestFindAllStrings()
    Dim a(10 To 30) As String
    Dim i As Long
    
    For i = 10 To 30
        a(i) = "Item " & i
    Next i
    a(12) = "Hello"
    a(16) = "Hello"
    a(24) = "Hello"
    a(25) = "Hello"
    
    Dim s() As String
    s = cArray.FindAll(a, AddressOf FindHelloString)
    Assert.AreEqual 4, cArray.GetLength(s)
    
    For i = 0 To 3
        Assert.AreEqual "Hello", s(i)
    Next i
End Sub

Public Sub TestFindNoStrings()
    Dim a(10) As String
    Dim i As Long
    Dim s() As String
    
    s = cArray.FindAll(a, AddressOf FindHelloString)
    
    Assert.AreEqual 0, cArray.GetLength(s)
End Sub

Public Sub TextObjectExists()
    Dim a(1 To 50) As Int32
    Dim i As Long
    
    For i = 1 To 50
        Set a(i) = New Int32
        a(i).mValue = i
    Next i
    
    Assert.IsTrue cArray.Exists(a, AddressOf FindInt32)
    
    a(32).mValue = 0
    Assert.IsFalse cArray.Exists(a, AddressOf FindInt32)
End Sub

Public Sub TestFindObject()
    Dim a(1 To 50) As Int32
    Dim i As Long
    Dim o As Int32
    
    For i = 1 To 50
        Set a(i) = New Int32
        a(i).mValue = i
    Next i
    
    Set o = a(32)
    Assert.AreEqual o, cArray.find(a, AddressOf FindInt32)
    
    a(32).mValue = 0
    Assert.IsNothing cArray.find(a, AddressOf FindInt32)
End Sub

Public Sub TestFindAllObjects()
    Dim a(1 To 50) As Int32
    Dim i As Long
    
    For i = 1 To 50
        Set a(i) = New Int32
        a(i).mValue = i
    Next i
    a(12).mValue = 32
    a(15).mValue = 32
    a(30).mValue = 32
    a(45).mValue = 32
    
    Dim o() As Int32
    o = cArray.FindAll(a, AddressOf FindInt32)
    Assert.AreEqual 5, cArray.GetLength(o)
    
    Assert.AreEqual a(12), o(0)
    Assert.AreEqual a(15), o(1)
    Assert.AreEqual a(30), o(2)
    Assert.AreEqual a(32), o(3)
    Assert.AreEqual a(45), o(4)
End Sub

Public Sub TestFindNoObjects()
    Dim a(10) As Int32
    Dim i As Long
    
    For i = 0 To 10
        Set a(i) = New Int32
    Next i
    
    Dim o() As Int32
    o = cArray.FindAll(a, AddressOf FindInt32)
    
    Assert.AreEqual 0, cArray.GetLength(o)
End Sub

Public Sub TestFindLastIndex()
    Dim a(10) As String
    Dim i As Long
    
    For i = 0 To 10
        a(i) = "Item " & i
    Next i
    
    a(3) = "Hello"
    a(4) = "Hello"
    Assert.AreEqual 4, cArray.FindLastIndex(a, AddressOf FindHelloString)
    Assert.AreEqual "Hello", cArray.FindLast(a, AddressOf FindHelloString)
    
    a(3) = "Bye"
    a(4) = "Bye"
    Assert.AreEqual -1, cArray.FindLastIndex(a, AddressOf FindHelloString)
    Assert.AreEqual "", cArray.FindLast(a, AddressOf FindHelloString)
End Sub

Public Sub TestTrueForAll()
    Dim a(10) As String
    Dim i As Long
    
    For i = 0 To 10
        a(i) = "Hello"
    Next i
    
    Assert.IsTrue cArray.TrueForAll(a, AddressOf FindHelloString)
    
    a(4) = "Bye"
    Assert.IsFalse cArray.TrueForAll(a, AddressOf FindHelloString)
End Sub

Public Sub TestFindIndexUDT()
    Dim a(10) As VBGUID
    Dim i As Long
    Dim v As VBGUID
    
    For i = 0 To 10
        a(i).Data1 = i
    Next i
    
    Assert.AreEqual 7, cArray.FindIndex(a, AddressOf FindVBGuid)
End Sub

Public Sub TestFindUDT()
    Dim a(10) As VBGUID
    Dim G As VBGUID
    
    a(8).Data1 = 7
    G = cArray.find(a, AddressOf FindVBGuid)
    Assert.AreEqual 7, G.Data1
End Sub

Public Sub TestFindAllUDTs()
    Dim a(20) As VBGUID
    Dim i As Long
    Dim G As Variant
    
    a(5).Data1 = 7
    a(9).Data1 = 7
    a(10).Data1 = 7
    a(16).Data1 = 7
    a(20).Data1 = 7
    
    G = cArray.FindAll(a, AddressOf FindVBGuid)
    Assert.AreEqual 5, cArray.GetLength(G)
    
    For i = 0 To 4
        Assert.AreEqual 7, G(i).Data1
    Next i
End Sub

Public Sub TestFindNoUDTs()
    Dim a(20) As VBGUID
    Dim G As VBGUID
    
    G = cArray.find(a, AddressOf FindVBGuid)
    Assert.IsTrue TypeOf G Is VBGUID
    Assert.AreEqual 0, G.Data1
End Sub
