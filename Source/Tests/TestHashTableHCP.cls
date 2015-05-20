VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHashTableHCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHashtableHCP
'
Option Explicit
'Implements ICategorizable

'Private mHt As Hashtable
'
'Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
'    Cat.ForFixture "System.Collections"
'End Sub
'
'Public Sub Setup()
'    Set mHt = Cor.NewHashtable(hcp:=New CaseInsensitiveHashCodePrvdr, comparer:=New CaseInsensitiveComparer)
'End Sub
'
'Public Sub TestStringKeys()
'    mHt.Add "Key", "Key"
'    Assert.IsTrue mHt.Contains("Key")
'    Assert.IsTrue mHt.Contains("KEY"), "Key and KEY should be same."
'    Assert.IsTrue mHt.Contains("kEY"), "Key and kEY should be same."
'End Sub
'
'
