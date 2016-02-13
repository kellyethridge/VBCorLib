VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: SecurityElement
'

''
' Provides simple XML text management capabilities.
'
' @remarks This class is primarily used to assist with the RSACryptoServiceProvider and DSACryptoServiceProvider
' classes when dealing with the keys as xml parameters. It is not meant for serious xml text manipulation.
' @see SecurityElementStatic
'
Option Explicit
Implements IObject

Private Const PROP_TAG          As String = "Tag"
Private Const PROP_TEXT         As String = "Text"
Private Const PROP_CHILDREN     As String = "Children"
Private Const PROP_ATTRIBUTES   As String = "Attributes"


Private mTag        As String
Private mText       As String
Private mChildren   As ArrayList
Private mAttributes As Hashtable


''
' Returns any attributes contained within the xml element.
'
' @return An IDictionary object containing key/value pairs of attributes or Nothing.
' @remarks The attribute names are case sensitive.
'
Public Property Get Attributes() As Hashtable
    If Not mAttributes Is Nothing Then
        Set Attributes = Cor.NewHashtable(mAttributes)
    End If
End Property

''
' Sets the attributes for this xml element.
'
' @param Value An IDictionary of key/value pairs as attributes.
' @remarks The attribute names and values are validated before the list is accepted.
'
Public Property Set Attributes(ByVal Value As Hashtable)
    If Value Is Nothing Then
        Set mAttributes = Nothing
    ElseIf Value.Count = 0 Then
        Set mAttributes = Nothing
    Else
        Dim Item    As DictionaryEntry
        Dim Table   As Hashtable
        Set Table = Cor.NewHashtable2(Value.Count)
        
        For Each Item In Value
            If Not SecurityElement.IsValidAttributeName(Item.Key) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementName, Item.Key))
            If Not SecurityElement.IsValidAttributeValue(Item.Value) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementValue, Item.Value))
                
            Table.Add Item.Key, Item.Value
        Next
        
        Set mAttributes = Table
    End If
End Property

''
' Returns a list of child SecurityElement objects.
'
' @return A list of children objects or Nothing.
'
Public Property Get Children() As ArrayList
    Set Children = mChildren
End Property

''
' Sets the list of children SecurityElements for this element.
'
' @param RHS The list of children to be set to this SecurityElement.
' @remarks DotNET doesn't seem to validate this list prior to setting it,
' so we won't bother to check either.
'
Public Property Set Children(ByVal RHS As ArrayList)
    Set mChildren = RHS
End Property

''
' Returns the tag name for this SecurityElement.
'
' @return The tag name for this SecurityElement.
'
Public Property Get Tag() As String
    Tag = mTag
End Property

''
' Sets the tag name for this SecurityElement.
'
' @param Value The new tag name.
' @remarks The tag is validated prior to setting.
'
Public Property Let Tag(ByRef Value As String)
    If Not SecurityElement.IsValidTag(Value) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementTag, Value))
        
    mTag = Value
End Property

''
' Returns the text between the opening and closing element tags.
'
' @return The element text.
'
Public Property Get Text() As String
    Text = mText
End Property

''
' Sets the text between the opening and closing element tags.
'
' @param Value The new element text.
' @remarks The text is validated prior to being set.
'
Public Property Let Text(ByRef Value As String)
    If Not SecurityElement.IsValidText(Value) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementText, Value))
    
    mText = Value
End Property

''
' Adds an attribute to the set of attributes for this SecurityElement.
'
' @param Name The name of the new attribute. This is case sensitive.
' @param Value The value of the new attribute.
'
Public Sub AddAttribute(ByRef Name As String, ByRef Value As String)
    If mAttributes Is Nothing Then
        Set mAttributes = New Hashtable
    End If
    
    If Not SecurityElement.IsValidAttributeName(Name) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementName, Name))
    If Not SecurityElement.IsValidAttributeValue(Value) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementValue, Value))
    If mAttributes.ContainsKey(Name) Then _
        Error.Argument Argument_AttributeNamesMustBeUnique
        
    mAttributes.Add Name, Value
End Sub

''
' Adds a new child to the SecurityElement.
'
' @param Child The child to be added.
'
Public Sub AddChild(ByVal Child As SecurityElement)
    If Child Is Nothing Then _
        Throw New ArgumentNullException
    
    If mChildren Is Nothing Then
        Set mChildren = New ArrayList
    End If
    
    Call mChildren.Add(Child)
End Sub

''
' Searches for the specified attribute and returns the value or an empty string.
'
' @param Name The name of the attribute to retrieve.
' @return Returns the attribute value or an empty string if not found.
' @remarks This is the same as the Attribute method in the .NET SecurityElement class.
'
Public Function GetAttribute(ByRef Name As String) As String
    If Not mAttributes Is Nothing Then
        If mAttributes.ContainsKey(Name) Then
            Dim Result As String
            Result = mAttributes(Name)
            
            If InStr(Result, "&") Then
                GetAttribute = Unescape(Result)
            Else
                GetAttribute = Result
            End If
        End If
    End If
End Function

''
' Returns a copy of the SecurityElement object.
'
' @return A copy of the object.
'
Public Function Copy() As SecurityElement
    Dim Bag As New PropertyBag
    Call Bag.WriteProperty("Copy", Me)
    Set Copy = Bag.ReadProperty("Copy")
End Function

''
' Searches for a child element by tag name.
'
' @param Tag The tag name of the child element.
' @return Returns the child element if found, or Nothing if not.
' @remarks Only the first level of children for this instance are searched,
' no grandchildren are searched.
'
Public Function SearchForChildByTag(ByVal Tag As String) As SecurityElement
    If mChildren Is Nothing Then Exit Function
    
    Dim Child As SecurityElement
    For Each Child In mChildren
        If Child.Tag = Tag Then
            Set SearchForChildByTag = Child
            Exit Function
        End If
    Next Child
End Function

''
' Searches for the text of a child with the specified tag name.
'
' @param Tag The tag name of the child being searched for.
' @Return Returns the text for the found child, or an empty string if not found.
' @remarks Only the first level of children for this instance are searched,
' no grandchildren are searched.
'
Public Function SearchForTextOfTag(ByVal Tag As String) As String
    Dim Child As SecurityElement
    Set Child = SearchForChildByTag(Tag)
    
    If Not Child Is Nothing Then
        SearchForTextOfTag = Child.Text
    End If
End Function

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to test for equality.
' @return Returns True if they are equal, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return Returns an integer.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return A string representation of this instance.
'
Public Function ToString() As String
    Dim sb As New StringBuilder
    Call AddSelf(sb, True)
    ToString = sb.ToString
End Function

Friend Sub AddSelf(ByVal sb As StringBuilder, Optional ByVal IsRoot As Boolean)
    Call sb.AppendFormat("<{0}", mTag)
    Call AddAttributes(sb)
    
    If HasChildren Or HasText Then
        Call sb.Append(">")
        Call sb.Append(mText)
        If HasChildren Then Call sb.AppendLine
        Call AddChildren(sb)
        Call sb.AppendFormat("</{0}>{1}", mTag, vbCrLf)
    Else
        Call sb.Append("/>")
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Tag As String, ByRef Text As String)
    If Not SecurityElement.IsValidTag(Tag) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementTag, Tag))
    If Not SecurityElement.IsValidText(Text) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementText, Text))
    
    mTag = Tag
    mText = Text
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Unescape(ByRef s As String) As String
    Dim sb As StringBuilder
    Set sb = StringBuilderCache.Acquire
    sb.AppendString s
    sb.Replace "&quot;", """"
    sb.Replace "&lt;", "<"
    sb.Replace "&gt;", ">"
    sb.Replace "&apos;", "'"
    sb.Replace "&amp;", "&"
    
    Unescape = StringBuilderCache.GetStringAndRelease(sb)
End Function

Private Sub AddAttributes(ByVal sb As StringBuilder)
    If mAttributes Is Nothing Then Exit Sub
    
    Dim de As DictionaryEntry
    For Each de In mAttributes
        Call sb.AppendFormat(" {0}={2}{1}{2}", de.Key, de.Value, Chr$(corQuoteChar))
    Next de
End Sub

Private Sub AddChildren(ByVal sb As StringBuilder)
    If mChildren Is Nothing Then Exit Sub
    
    Dim Child As SecurityElement
    For Each Child In mChildren
        Call Child.AddSelf(sb)
    Next Child
End Sub

Private Function HasChildren() As Boolean
    If Not mChildren Is Nothing Then
        HasChildren = (mChildren.Count > 0)
    End If
End Function

Private Function HasText() As Boolean
    HasText = (Len(mText) > 0)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mTag = .ReadProperty(PROP_TAG)
        mText = .ReadProperty(PROP_TEXT)
        Set mChildren = .ReadProperty(PROP_CHILDREN, Nothing)
        Set mAttributes = .ReadProperty(PROP_ATTRIBUTES, Nothing)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_TAG, mTag)
        Call .WriteProperty(PROP_TEXT, mText)
        Call .WriteProperty(PROP_CHILDREN, mChildren, Nothing)
        Call .WriteProperty(PROP_ATTRIBUTES, mAttributes, Nothing)
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
