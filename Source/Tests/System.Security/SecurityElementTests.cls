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
            
        Case "GetAttribute_WithValueContainingEscapedSequence_UnescapesSequence"
            Test.Use "val&lt;ue", "val<ue"
            Test.Use "val&gt;ue", "val>ue"
            Test.Use "val&quot;ue", "val""ue"
            Test.Use "val&apos;ue", "val'ue"
            Test.Use "val&amp;ue", "val&ue"
            
        Case "Equal_WithTagAndTextValues_ReturnsExpected"
            Test.Use "Tag", "Text", True
            Test.Use "tag", "Text", False
            Test.Use "Tag", "text", False
            
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

Public Sub Constructor_WithValidTagAndInvalidText_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewSecurityElement "name", "te>xt"
    
    AssertArgumentException Err
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

Public Sub AddAttribute_WithValidValues_AddsToAttributes()
    Dim Element As New SecurityElement
    
    Element.AddAttribute "version", "1.0"
    
    Assert.That Element.Attributes("version"), Iz.EqualTo("1.0")
End Sub

Public Sub AddAttribute_WithInvalidName_ThrowsArgumentException()
    Dim Element As New SecurityElement
    
    On Error Resume Next
    Element.AddAttribute "ver<sion", "1.0"
    
    AssertArgumentException Err
End Sub

Public Sub AddAttribute_WithInvalidValue_ThrowsArgumentException()
    Dim Element As New SecurityElement
    
    On Error Resume Next
    Element.AddAttribute "version", "1<0"
    
    AssertArgumentException Err
End Sub

Public Sub AddAttribute_WithExistingAttribute_ThrowsArgumentException()
    Dim Element As New SecurityElement
    Element.AddAttribute "version", "1.0"
    
    On Error Resume Next
    Element.AddAttribute "version", "2.0"
    
    AssertArgumentException Err
End Sub

Public Sub GetAttribute_WithNonExistingName_ReturnsEmpty()
    Dim Element As New SecurityElement
    Dim Actual As String
    
    Actual = Element.GetAttribute("version")
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub GetAttribute_WithExistingName_ReturnsValue()
    Dim Element As New SecurityElement
    Element.AddAttribute "version", "1.0"
    
    Dim Actual As String
    Actual = Element.GetAttribute("version")
    
    Assert.That Actual, Iz.EqualTo("1.0")
End Sub

Public Sub GetAttribute_WithValueContainingEscapedSequence_UnescapesSequence(ByVal EscapedValue As String, ByVal Expected As String)
    Dim Element As New SecurityElement
    Element.AddAttribute "version", EscapedValue
        
    Dim Actual As String
    Actual = Element.GetAttribute("version")
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Children_WhenNewlyConstructed_ReturnsNothing()
    Dim Element As New SecurityElement
    
    Dim Actual As ArrayList
    Set Actual = Element.Children
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub Children_WhenSet_StoresChildren()
    Dim Children As New ArrayList
    Dim Child As New SecurityElement
    Children.Add Child
    Dim Element As New SecurityElement
    
    Set Element.Children = Children
    
    Assert.That Element.Children, Iz.Not.Nothing
    Assert.That Element.Children(0), Iz.SameAs(Child)
End Sub

Public Sub Children_WhenSetAndContainsNullChild_ThrowsArgumentException()
    Dim Children As New ArrayList
    Children.Add Nothing
    Dim Element As New SecurityElement
    
    On Error Resume Next
    Set Element.Children = Children
    
    AssertArgumentException Err
End Sub

Public Sub AddChild_WithNothing_ThrowsArgumentNullException()
    Dim Element As New SecurityElement
    
    On Error Resume Next
    Element.AddChild Nothing
    
    AssertArgumentNullException Err, "Child"
End Sub

Public Sub AddChild_WithValue_AddsToChildrent()
    Dim Element As New SecurityElement
    Dim Child As New SecurityElement
    
    Element.AddChild Child
    
    Assert.That Element.Children(0), Iz.SameAs(Child)
End Sub

Public Sub Equal_WithTagAndTextValues_ReturnsExpected(ByVal Tag As String, ByVal Text As String, ByVal Expected As Boolean)
    Dim Element1 As SecurityElement
    Dim Element2 As SecurityElement
    Set Element1 = Cor.NewSecurityElement("Tag", "Text")
    Set Element2 = Cor.NewSecurityElement(Tag, Text)
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equal_WithDifferingAttributeCount_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddAttribute "a1", "a1"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithDifferingAttributeCount_ReturnsFalse2()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element2.AddAttribute "a1", "a1"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithDifferingAttributeCount_ReturnsFalse3()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddAttribute "a1", "a1"
    Element1.AddAttribute "a2", "a2"
    Element2.AddAttribute "a1", "a1"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithSameAttributes_ReturnsTrue()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddAttribute "a1", "a1"
    Element2.AddAttribute "a1", "a1"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Equal_WithDifferingAttributeNames_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddAttribute "a1", "a1"
    Element2.AddAttribute "a2", "a1"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithDifferingAttributeValues_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddAttribute "a1", "a1"
    Element2.AddAttribute "a1", "a2"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithSameAttributesInDifferingOrder_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddAttribute "a1", "a1"
    Element1.AddAttribute "a2", "a2"
    Element2.AddAttribute "a2", "a2"
    Element2.AddAttribute "a1", "a1"
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithDifferingChildCount_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddChild New SecurityElement
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithDifferingChildCount_ReturnsFalse2()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element2.AddChild New SecurityElement
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithDifferingChildCount_ReturnsFalse3()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddChild New SecurityElement
    Element1.AddChild New SecurityElement
    Element2.AddChild New SecurityElement
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithEqualChildren_ReturnsTrue()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddChild Cor.NewSecurityElement("tag", "text")
    Element2.AddChild Cor.NewSecurityElement("tag", "text")
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.True
End Sub

Public Sub Equal_WithUnequalChildren_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddChild Cor.NewSecurityElement("tag1", "text")
    Element2.AddChild Cor.NewSecurityElement("tag2", "text")
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Equal_WithSameChildrenInDifferingOrder_ReturnsFalse()
    Dim Element1 As New SecurityElement
    Dim Element2 As New SecurityElement
    Element1.AddChild Cor.NewSecurityElement("tag1", "text")
    Element1.AddChild Cor.NewSecurityElement("tag2", "text")
    Element2.AddChild Cor.NewSecurityElement("tag2", "text")
    Element2.AddChild Cor.NewSecurityElement("tag1", "text")
    
    Dim Actual As Boolean
    Actual = Element1.Equal(Element2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Copy_WithNoAttributesOrChildren_ReturnsShallowCopy()
    Dim Element As SecurityElement
    Set Element = Cor.NewSecurityElement("tag", "text")
    
    Dim Actual As SecurityElement
    Set Actual = Element.Copy
    
    Assert.That Actual, Iz.Not.Nothing
    Assert.That Actual.Tag, Iz.EqualTo("tag")
    Assert.That Actual.Text, Iz.EqualTo("text")
    Assert.That Actual.Attributes, Iz.Nothing
    Assert.That Actual.Children, Iz.Nothing
End Sub

Public Sub Copy_WithAttributesAndChildren_ReturnsWithShallowCopy()
    Dim Element As SecurityElement
    Set Element = Cor.NewSecurityElement("tag", "text")
    Element.AddAttribute "att", "value"
    Dim Child As New SecurityElement
    Element.AddChild Child
    
    Dim Actual As SecurityElement
    Set Actual = Element.Copy
    
    Assert.That Actual, Iz.Not.Nothing
    Assert.That Actual.Tag, Iz.EqualTo("tag")
    Assert.That Actual.Text, Iz.EqualTo("text")
    Assert.That Actual.Attributes, Iz.Not.Nothing, "Attributes"
    Assert.That Actual.Attributes("att"), Iz.EqualTo("value")
    Assert.That Actual.Children, Iz.Not.Nothing, "Children"
    Assert.That Actual.Children(0), Iz.SameAs(Child)
End Sub

Public Sub SearchForTextOfTag_WithNoChildren_ReturnsEmpty()
    Dim Element As New SecurityElement
    
    Dim Actual As String
    Actual = Element.SearchForTextOfTag("tag")
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub SearchForTextOfTag_WithSelfHavingMatchingTag_ReturnsText()
    Dim Element As SecurityElement
    Set Element = Cor.NewSecurityElement("tag", "text")
    
    Dim Actual As String
    Actual = Element.SearchForTextOfTag("tag")
    
    Assert.That Actual, Iz.EqualTo("text")
End Sub

Public Sub SearchForTextOfTag_WithChildMatchingTag_ReturnsChildText()
    Dim Element As New SecurityElement
    Element.AddChild Cor.NewSecurityElement("tag", "child_text")
    
    Dim Actual As String
    Actual = Element.SearchForTextOfTag("tag")
    
    Assert.That Actual, Iz.EqualTo("child_text")
End Sub

Public Sub SearchForTextOfTag_WithGrandchildMatchingTag_ReturnsGrandchildText()
    Dim Element As New SecurityElement
    Dim Child As New SecurityElement
    Element.AddChild Child
    Child.AddChild Cor.NewSecurityElement("tag", "grandchild_text")
    
    Dim Actual As String
    Actual = Element.SearchForTextOfTag("tag")
    
    Assert.That Actual, Iz.EqualTo("grandchild_text")
End Sub

Public Sub SearchForChildByTag_WithNoChildren_ReturnsNothing()
    Dim Element As New SecurityElement
    
    Dim Actual As SecurityElement
    Set Actual = Element.SearchForChildByTag("tag")
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub SearchForChildByTag_WithNoMatchingTag_ReturnsNothing()
    Dim Element As New SecurityElement
    Element.AddChild Cor.NewSecurityElement("tag2", "text")
    
    Dim Actual As SecurityElement
    Set Actual = Element.SearchForChildByTag("tag")
    
    Assert.That Actual, Iz.Nothing
End Sub

Public Sub SearchForChildByTag_WithMatchingTag_ReturnsChild()
    Dim Element As New SecurityElement
    Element.AddChild Cor.NewSecurityElement("tag", "text")
    
    Dim Actual As SecurityElement
    Set Actual = Element.SearchForChildByTag("tag")
    
    Assert.That Actual.Text, Iz.EqualTo("text")
End Sub

Public Sub Serialization_WhenCalled_SerializesObject()
    Dim Element As SecurityElement
    Set Element = Cor.NewSecurityElement("elem", "elem_text")
    Element.AddAttribute "name", "value"
    Dim Child As SecurityElement
    Set Child = Cor.NewSecurityElement("child", "child_text")
    Child.AddAttribute "child_name", "child_value"
    Element.AddChild Child
    Dim Grand As SecurityElement
    Set Grand = Cor.NewSecurityElement("grand", "grand_text")
    Grand.AddAttribute "grand_name", "grand_value"
    Child.AddChild Grand
    
    Dim Actual As SecurityElement
    Set Actual = Persist(Element)
    
    Assert.That Actual.Equal(Element), Iz.True
End Sub

Public Sub ToString_WithTagOnly_ReturnsExpected()
    Dim Element As New SecurityElement
    Element.Tag = "tag"
    
    Dim Actual As String
    Actual = Element.ToString
    
    Assert.That Actual, Iz.EqualTo(t("<tag/>\r\n"))
End Sub

Public Sub ToString_WithTagAndText_ReturnsExpected()
    Dim Element As SecurityElement
    Set Element = Cor.NewSecurityElement("tag", "text")
    
    Dim Actual As String
    Actual = Element.ToString
    
    Assert.That Actual, Iz.EqualTo(t("<tag>text</tag>\r\n"))
End Sub

Public Sub ToString_WithTagAndAttribute_ReturnsExpected()
    Dim Element As New SecurityElement
    Element.Tag = "tag"
    Element.AddAttribute "att", "val"
    Element.AddAttribute "att2", "val2"
    
    Dim Actual As String
    Actual = Element.ToString
    
    Assert.That Actual, Iz.EqualTo(t("<tag att=""val""\r\natt2=""val2""/>\r\n"))
End Sub

Public Sub ToString_WithTagAndChild_ReturnsExpected()
    Dim Element As New SecurityElement
    Element.Tag = "parent_tag"
    Dim Child As New SecurityElement
    Child.Tag = "child_tag"
    Element.AddChild Child
    
    Dim Actual As String
    Actual = Element.ToString
    
    Assert.That Actual, Iz.EqualTo(t("<parent_tag>\r\n<child_tag/>\r\n</parent_tag>\r\n"))
End Sub

Public Sub ToString_WithTagTextAndChild_ReturnsExpected()
    Dim Element As New SecurityElement
    Element.Tag = "parent_tag"
    Element.Text = "parent_text"
    Dim Child As New SecurityElement
    Child.Tag = "child_tag"
    Element.AddChild Child
    
    Dim Actual As String
    Actual = Element.ToString
    
    Assert.That Actual, Iz.EqualTo(t("<parent_tag>parent_text\r\n<child_tag/>\r\n</parent_tag>\r\n"))
End Sub

Public Sub ToString_WithTagTextAndChildTagText_ReturnsExpected()
    Dim Element As New SecurityElement
    Element.Tag = "parent_tag"
    Element.Text = "parent_text"
    Dim Child As New SecurityElement
    Child.Tag = "child_tag"
    Child.Text = "child_text"
    Element.AddChild Child
    
    Dim Actual As String
    Actual = Element.ToString
    
    Assert.That Actual, Iz.EqualTo(t("<parent_tag>parent_text\r\n<child_tag>child_text</child_tag>\r\n</parent_tag>\r\n"))
End Sub
























