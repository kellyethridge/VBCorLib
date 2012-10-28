VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDESWeakKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDESWeakKeys
'
Option Explicit

Public Sub TestNonWeakKey()
    Dim b() As Byte
    b = NewArray(ciByte, 1, 2, 3, 4, 5, 6, 7, 8)
    Assert.IsFalse DES.IsWeakKey(b)
    
    b = NewArray(ciByte, 2, 2, 2, 2, 0, 0, 0, 0)
    Assert.IsFalse DES.IsWeakKey(b)
End Sub

Public Sub Test4WeakKeys()
    Dim b() As Byte
    b = NewArray(ciByte, 0, 0, 0, 0, 0, 0, 0, 0)
    Assert.IsTrue DES.IsWeakKey(b), "All zeros should be weak."
    
    b = NewArray(ciByte, &HFE, &HFE, &HFE, &HFE, &HFE, &HFE, &HFE, &HFE)
    Assert.IsTrue DES.IsWeakKey(b), "All &HFE should be weak."

    b = NewArray(ciByte, &H1F, &H1F, &H1F, &H1F, &HE, &HE, &HE, &HE)
    Assert.IsTrue DES.IsWeakKey(b), "1F1F1F1F0E0E0E0E should be weak."
    
    b = NewArray(ciByte, &HE0, &HE0, &HE0, &HE0, &HF1, &HF1, &HF1, &HF1)
    Assert.IsTrue DES.IsWeakKey(b), "E0E0E0E0F1F1F1F1 should be weak."
End Sub

Public Sub TestNonSemiWeakKeys()
    Dim b() As Byte
    b = NewArray(ciByte, 123, 0, 0, 0, 0, 0, 0, 0)
    Assert.IsFalse DES.IsSemiWeakKey(b), "Should not be semi-weak."
End Sub

Public Sub Test12SemiWeakKeys()
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("E001E001F101F101")), "E001E001F101F101 should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("01E001E001F101F1")), "01E001E001F101F1 should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("FE1FFE1FFE0EFE0E")), "FE1FFE1FFE0EFE0E should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("1FFE1FFE0EFE0EFE")), "1FFE1FFE0EFE0EFE should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("E01FE01FF10EF10E")), "E01FE01FF10EF10E should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("1FE01FE00EF10EF1")), "1FE01FE00EF10EF1 should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("01FE01FE01FE01FE")), "01FE01FE01FE01FE should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("FE01FE01FE01FE01")), "FE01FE01FE01FE01 should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("011F011F010E010E")), "011F011F010E010E should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("1F011F010E010E01")), "1F011F010E010E01 should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("E0FEE0FEF1FEF1FE")), "E0FEE0FEF1FEF1FE should be weak."
    Assert.IsTrue DES.IsSemiWeakKey(GetBytes("FEE0FEE0FEF1FEF1")), "FEE0FEE0FEF1FEF1 should be weak."
End Sub



Private Function GetBytes(ByVal s As String) As Byte()
    Dim i As Long
    Dim b() As Byte
    
    ReDim b(7)
    For i = 0 To 7
        b(i) = CByte("&h" & Mid$(s, i * 2 + 1, 2))
    Next i
    
    GetBytes = b
End Function
