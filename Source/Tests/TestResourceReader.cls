VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResourceReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestResourceReader
'
Option Explicit
Implements ICategorizable

Private mReader As ResourceReader

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Setup()
    Set mReader = Cor.NewResourceReader(Path.Combine(App.Path, "VBCorLibTests.res"))
End Sub

Public Sub Teardown()
    mReader.CloseReader
End Sub

Public Sub TestEnumStrings()
    Dim Key As ResourceKey
    Dim s As String
    Dim de As DictionaryEntry
    
    For Each de In mReader
        Assert.IsTrue VarType(de.Key) = vbObject
        Assert.IsTrue TypeOf de.Key Is ResourceKey
        
        Set Key = de.Key
        Debug.Print "{" & Key.ToString & "} = ";
        Select Case Key.ResourceType
            Case 6
                Debug.Print de.Value
            Case 2
                Debug.Print "Bitmap"
            Case 14
                Debug.Print "Icon Group"
            Case 3
                Debug.Print "Icon"
            Case 1
                Debug.Print "Cursor"
            Case 12
                Debug.Print "Cursor Group"
        End Select
    Next de
    
End Sub


