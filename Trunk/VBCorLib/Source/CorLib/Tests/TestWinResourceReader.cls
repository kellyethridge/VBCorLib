VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestWinResourceReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestWinResourceReader
'
Option Explicit

Private Reader As WinResourceReader

Public Sub Setup()
    Set Reader = Cor.NewWinResourceReader("source\resexplorer\resexplorer.exe")
End Sub

Public Sub TestIteration()
    Dim Entry As DictionaryEntry
    For Each Entry In Reader
        Dim Key As ResourceKey
        Set Key = Entry.Key
        Debug.Print Key.ToString
'        If Key.ResourceType = ResourceTypes.StringResource Then
'            Debug.Print Entry.value
'        End If
    Next Entry
End Sub
