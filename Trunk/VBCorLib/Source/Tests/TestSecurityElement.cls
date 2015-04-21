VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSecurityElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestSecurityElement
'
Option Explicit

Public Sub TestEmptyTag()
    Dim se As New SecurityElement
    se.Tag = "Name"
    Assert.AreEqual "Name", se.Tag
    Assert.AreEqual "<Name/>", Replace$(se.ToString, vbCrLf, "")
End Sub

Public Sub TestWithText()
    Dim se As New SecurityElement
    se.Tag = "Name"
    se.Text = "Kelly"
    Assert.AreEqual "Kelly", se.Text
    Assert.AreEqual "<Name>Kelly</Name>", Replace$(se.ToString, vbCrLf, "")
End Sub

Public Sub TestWithAttributesOnly()
    Dim se As New SecurityElement
    se.Tag = "Name"
    se.AddAttribute "Age", 33
    se.AddAttribute "Sex", "Male"
    Assert.AreEqual "33", se.GetAttribute("Age")
    Assert.AreEqual "Male", se.GetAttribute("Sex")
    Assert.AreEqual "<Name Age=" & Chr$(34) & "33" & Chr$(34) & " Sex=" & Chr$(34) & "Male" & Chr$(34) & "/>", Replace$(se.ToString, vbCrLf, "")
End Sub

Public Sub TestWithAttributesAndText()
    Dim se As New SecurityElement
    se.Tag = "Name"
    se.Text = "Kelly"
    se.AddAttribute "Age", 33
    se.AddAttribute "Sex", "Male"
    Assert.AreEqual "33", se.GetAttribute("Age")
    Assert.AreEqual "Male", se.GetAttribute("Sex")
    Assert.AreEqual "<Name Age=" & Chr$(34) & "33" & Chr$(34) & " Sex=" & Chr$(34) & "Male" & Chr$(34) & ">Kelly</Name>", Replace$(se.ToString, vbCrLf, "")
End Sub

Public Sub TestWithOneChild()
    Dim se As New SecurityElement
    se.Tag = "Main"
    
    Dim se2 As New SecurityElement
    se2.Tag = "Child"
    se.AddChild se2
    
    Assert.IsNotNothing se.Children
    Assert.AreEqual 1, se.Children.Count
    Assert.AreSame se2, se.SearchForChildByTag("Child")
    Assert.AreEqual "<Main><Child/></Main>", Replace$(se.ToString, vbCrLf, "")
End Sub

Public Sub TestEscape()
    Assert.AreEqual "&gt;&lt;&quot;&apos;&amp;", SecurityElement.Escape("><" & Chr$(34) & "'&")
End Sub

Public Sub TestFromSimpleTag()
    Dim se As SecurityElement
    Set se = SecurityElement.FromString("<tag/>")
    Assert.IsNotNothing se
    Assert.AreEqual "tag", se.Tag
End Sub

Public Sub TestFromSimpleTagWithOneAttribute()
    Dim se As SecurityElement
    Set se = SecurityElement.FromString("<tag name=kelly/>")
    Assert.IsNotNothing se
    Assert.AreEqual "tag", se.Tag
    Assert.IsNotNothing se.Attributes
    Assert.AreEqual 1, se.Attributes.Count
    Assert.IsTrue se.Attributes.ContainsKey("name")
    Assert.AreEqual "kelly", se.GetAttribute("name")
End Sub

Public Sub TestFromComplexTagNoChildren()
    Dim se As SecurityElement
    Set se = SecurityElement.FromString("<tag name=" & Chr$(34) & "kelly" & Chr$(34) & " age = 33>wow</tag>")
    Assert.IsNotNothing se
    Assert.AreEqual "tag", se.Tag
    Assert.AreEqual "kelly", se.GetAttribute("name")
    Assert.AreEqual "33", se.GetAttribute("age")
    Assert.AreEqual "wow", se.Text
End Sub

Public Sub TestWithChildren()
    Dim sb As New StringBuilder
    sb.AppendFormat "<parent name={0}some  ing{0}>I'm the parent", Chr$(34)
    sb.Append "<child1 name=1>One</child1>"
    sb.Append "<child2 name=2>Two</child2>"
    sb.Append "<child3 name=3>Three</child3>"
    sb.Append "</parent>"

    Dim se As SecurityElement
    Set se = SecurityElement.FromString(sb.ToString)
    
    Assert.IsNotNothing se
    Assert.AreEqual "parent", se.Tag
    Assert.AreEqual "I'm the parent", se.Text
    Assert.AreEqual "some  ing", se.GetAttribute("name")
    Assert.IsNotNothing se.Children
    
    Dim se2 As SecurityElement
    Set se2 = se.SearchForChildByTag("child1")
    Assert.IsNotNothing se2
    Assert.AreEqual "1", se2.GetAttribute("name")
    Assert.AreEqual "One", se2.Text
    
    Set se2 = se.SearchForChildByTag("child2")
    Assert.IsNotNothing se2
    Assert.AreEqual "2", se2.GetAttribute("name")
    Assert.AreEqual "Two", se2.Text
    
    Set se2 = se.SearchForChildByTag("child3")
    Assert.IsNotNothing se2
    Assert.AreEqual "3", se2.GetAttribute("name")
    Assert.AreEqual "Three", se2.Text
End Sub

Public Sub TestWithGrandChildren()
    Dim Root As SecurityElement
    Set Root = Cor.NewSecurityElement("Root", "I'm at the root.")
    Root.AddChild Cor.NewSecurityElement("Child1", "I'm a child.")
    
    Dim Child As SecurityElement
    Set Child = Cor.NewSecurityElement("Child2", "I'm also a child.")
    Child.AddChild Cor.NewSecurityElement("Grand1", "I'm a grandchild.")
    Child.AddChild Cor.NewSecurityElement("Grand2", "I'm a granderchild.")
    
    Root.AddChild Child
    
    Dim Root2 As SecurityElement
    Set Root2 = SecurityElement.FromString(Root.ToString)
    
    Assert.AreEqual Root.ToString, Root2.ToString
End Sub
