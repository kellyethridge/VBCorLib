VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIEntryComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IComparer

''
' Compare the elements of the ArrayList. This comparer class is designed
' specifically to compare DictionaryEntry objects holding key-value
' pairs from an INI file section.
'
Private Function IComparer_Compare(x As Variant, y As Variant) As Long
    Dim Entry1 As DictionaryEntry
    Dim Entry2 As DictionaryEntry
    
    ' We could just use late-binding, but what the heck.
    Set Entry1 = x
    Set Entry2 = y
    
    ' We know that an INI key is a string, so just use StrComp.
    IComparer_Compare = StrComp(Entry1.Key, Entry2.Key, vbTextCompare)
End Function
