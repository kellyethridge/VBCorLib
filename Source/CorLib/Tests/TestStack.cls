VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private st As Stack

Public Sub Setup()
    Set st = New Stack
End Sub

Public Sub TestEmptyStack()
    Assert.AreEqual 0, st.Count, "Count should return 0."
End Sub

Public Sub TestCountAfterPushingOne()
    st.Push 1
    Assert.AreEqual 1, st.Count, "Count should return 1."
End Sub

Public Sub TestCountAfterPushingTwo()
    st.Push 1
    st.Push 2
    Assert.AreEqual 2, st.Count, "Count should return 2."
End Sub

Public Sub TestCountAfterPoppingOne()
    st.Push 1
    st.Push 2
    st.Push 3
    Assert.AreEqual 3, st.Count, "Count should return 3."
    st.Pop
    Assert.AreEqual 2, st.Count, "Count should return 2 after popping one."
End Sub

Public Sub TestContents()
    st.Push 1
    st.Push 2
    
    Assert.AreEqual 2, st.Pop, "Pop should return 2."
    Assert.AreEqual 1, st.Pop, "Pop should return 1."
End Sub

Public Sub TestContains()
    st.Push 1
    st.Push 2
    
    Assert.IsTrue st.Contains(2), "Contains should return True."
End Sub

Public Sub TestNotContains()
    st.Push 1
    st.Push 2
    
    Assert.IsFalse st.Contains(3), "Contains should return False."
End Sub

Public Sub TestCopyTo()
    Dim i As Long
    For i = 1 To 10
        st.Push i
    Next i
    
    Dim a(0 To 9) As Long
    st.CopyTo a, 0
    For i = 0 To 9
        Assert.AreEqual 10 - i, a(i)
    Next i
End Sub

Public Sub TestPeek()
    st.Push 1
    st.Push 2
    st.Push 3
    
    Assert.AreEqual 3, st.Peek, "Peek should return 3."
    Assert.AreEqual 3, st.Count, "List should not change size."
    Dim i As Long
    For i = 3 To 1 Step -1
        Assert.AreEqual i, st.Pop, "Stack changed."
    Next i
End Sub

Public Sub TestToArray()
    st.Push 1
    st.Push 2
    st.Push 3
    
    Dim a() As Variant
    a = st.ToArray
    Assert.AreEqual 3, cArray.GetLength(a), "Length should be 3."
    Assert.AreEqual 3, a(0)
    Assert.AreEqual 2, a(1)
    Assert.AreEqual 1, a(2)
End Sub

Public Sub TestEnumerator()
    Dim i As Long
    For i = 1 To 10
        st.Push i
    Next i
    
    Dim v As Variant
    i = 10
    For Each v In st
        Assert.AreEqual i, v
        i = i - 1
    Next v
End Sub

Public Sub TestPersitence()
    Dim i As Long
    For i = 1 To 10
        st.Push i
    Next i
    
    Dim Copy As Stack
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    pp1.WriteProperty "stack", st
    pp2.Contents = pp1.Contents
    
    Set Copy = pp2.ReadProperty("stack", Nothing)
    For i = 1 To 10
        Assert.AreEqual st.Pop, Copy.Pop
    Next i
End Sub

Public Sub TestBigStack()
    Dim i As Long
    For i = 1 To 100
        st.Push i
    Next i
    
    For i = 100 To 1 Step -1
        Assert.AreEqual i, st.Pop
    Next i
End Sub

Public Sub TestInitWithElements()
    Dim a() As Long
    a = NewArray(vbLong, 1, 2, 3, 4, 5)
    
    Set st = NewStack(a)
    Assert.AreEqual 5, st.Count
    
    Dim i As Long
    For i = 1 To 5
        Assert.AreEqual 6 - i, st.Pop
    Next i

End Sub
