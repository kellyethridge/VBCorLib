VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityElementTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SecurityElementTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "IsValidTag_WithValues_ReturnsExpected"
            Test.Use "name", True
            Test.Use "na<me", False
            Test.Use "na>me", False
            Test.Use "na me", False
            Test.Use "", True
            Test.Use vbNullString, True
    
        Case "IsValidText_WithValues_ReturnsExpected"
            Test.Use "val<ue", False
            Test.Use "val>ue", False
            Test.Use "value", True
            Test.Use "val ue", True
            Test.Use "", True
            Test.Use vbNullString, True
            
        Case "IsValidAttributeName_WithValues_ReturnsExpected"
            Test.Use "name", True
            Test.Use "na<me", False
            Test.Use "na>me", False
            Test.Use "na me", False
            Test.Use "", True
            Test.Use vbNullString, True
        
        Case "IsValidAttributeValue_WithValues_ReturnsExpected"
            Test.Use "val<ue", False
            Test.Use "val>ue", False
            Test.Use "val""ue", False
            Test.Use "value", True
            Test.Use "val ue", True
            Test.Use "", True
            Test.Use vbNullString, True
        
        Case "Escape_WithValues_ReturnsExpected"
            Test.Use "val<ue", "val&lt;ue"
            Test.Use "val>ue", "val&gt;ue"
            Test.Use "val""ue", "val&quot;ue"
            Test.Use "val'ue", "val&apos;ue"
            Test.Use "val&ue", "val&amp;ue"
            Test.Use "value", "value"
            
        Case "Constructor_WithTag_SetsDefaultValues"
            Test.Use "tag"
            Test.Use ""
            
        Case "Constructor_WithInvalidTag_ThrowsArgumentException"
            Test.Use "na<me"
            Test.Use "nam>e"
            
    End Select
End Sub

Public Sub IsValidTag_WithValues_ReturnsExpected(ByVal Tag As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = SecurityElement.IsValidTag(Tag)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsValidText_WithValues_ReturnsExpected(ByVal Text As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = SecurityElement.IsValidText(Text)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsValidAttributeName_WithValues_ReturnsExpected(ByVal Name As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = SecurityElement.IsValidAttributeName(Name)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsValidAttributeValue_WithValues_ReturnsExpected(ByVal Value As String, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = SecurityElement.IsValidAttributeValue(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Escape_WithValues_ReturnsExpected(ByVal Value As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = SecurityElement.Escape(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Constructor_WithTag_SetsDefaultValues(ByVal Tag As String)
    Dim se As SecurityElement
    
    Set se = Cor.NewSecurityElement(Tag)
    
    Assert.That se.Attributes, Iz.Nothing
    Assert.That se.Children, Iz.Nothing
    Assert.That se.Tag, Iz.EqualTo(Tag)
    Assert.That se.Text, Iz.Empty
End Sub

Public Sub Constructor_WithInvalidTag_ThrowsArgumentException(ByVal InvalidTag As String)
    On Error Resume Next
    
    Cor.NewSecurityElement InvalidTag
    
    Dim Ex As ArgumentException
    Set Ex = AssertArgumentException(Err)
    Assert.That Ex.Message, Iz.StringContaining(InvalidTag)
End Sub

Public Sub Constructor_WithTagAndText_SetsDefaultValues()
    Dim se As SecurityElement
    
    Set se = Cor.NewSecurityElement("tag", "text")
    
    Assert.That se.Attributes, Iz.Nothing
    Assert.That se.Children, Iz.Nothing
    Assert.That se.Tag, Iz.EqualTo("tag")
    Assert.That se.Text, Iz.EqualTo("text")
End Sub

Public Sub Constructor_WithInvalidTagAndValidText_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewSecurityElement "na<me", "text"
    
    Dim Ex As ArgumentException
    Set Ex = AssertArgumentException(Err)
    Assert.That Ex.Message, Iz.StringContaining("na<me")
End Sub

Public Sub Attributes_WhenSetToNothing_ReturnsNothing()
    Dim Element As New SecurityElement
    
    Set Element.Attributes = Nothing
    
    Assert.That Element.Attributes, Iz.Nothing
End Sub

Public Sub Attributes_WhenSetWithEmptyTable_ReturnsNothing()
    Dim Element As New SecurityElement
    
    Set Element.Attributes = New Hashtable
    
    Assert.That Element.Attributes, Iz.Nothing
End Sub

Public Sub Attributes_WhenSetWithValidValues_ReturnsValidValues()
    Dim Element As New SecurityElement
    Dim Values As New Hashtable
    Values.Add "name", "kelly"
    
    Set Element.Attributes = Values
    
    Dim Actual As Hashtable
    Set Actual = Element.Attributes
    Assert.That Actual, Iz.Not.Nothing
    Assert.That Actual.Count, Iz.EqualTo(1)
    Assert.That Actual("name"), Iz.EqualTo("kelly")
End Sub

Public Sub Attributes_WhenSetWithInvalidKey_ThrowsArgumentException()
    Dim Element As New SecurityElement
    Dim Values As New Hashtable
    Values.Add "The>Key", "value"
    
    On Error Resume Next
    Set Element.Attributes = Values
    
    AssertArgumentException Err
End Sub

Public Sub Attributes_WhenSetWithInvalidValue_ThrowsArgumentException()
    Dim Element As New SecurityElement
    Dim Values As New Hashtable
    Values.Add "TheKey", "val>ue"
    
    On Error Resume Next
    Set Element.Attributes = Values
    
    AssertArgumentException Err
End Sub

Public Sub Attributes_WhenChangeExisting_DoesNotAlterValue()
    Dim Element As New SecurityElement
    Dim Values As New Hashtable
    Values.Add "Key", "Value"
    Set Element.Attributes = Values
    
    Element.Attributes.Add "NewKey", "NewValue"
    
    Assert.That Element.Attributes.ContainsKey("NewKey"), Iz.False
End Sub

Public Sub Attributes_WhenChangingOriginalHashtable_DoesNotAlterValue()
    Dim Element As New SecurityElement
    Dim Values As New Hashtable
    Values.Add "Key", "Value"
    Set Element.Attributes = Values

    Values.Add "NewKey", "NewValue"
    
    Assert.That Element.Attributes.ContainsKey("NewKey"), Iz.False
End Sub

Public Sub Tag_WhenSetWithInvalidValue_ThrowsArgumentException()
    Dim Element As New SecurityElement
    
    On Error Resume Next
    Element.Tag = "Ta<g"
    
    AssertArgumentException Err
End Sub

Public Sub Text_WhenSetWithInvalidValue_ThrowsArgumentException()
    Dim Element As New SecurityElement
    
    On Error Resume Next
    Element.Text = "Te<xt"
    
    AssertArgumentException Err
End Sub











































