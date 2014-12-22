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
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: SecurityElement
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
' @return A IDictionary object containing key/value pairs of attributes or Nothing.
' @remarks The attribute names are case sensitive.
'
Public Property Get Attributes() As Hashtable
    Set Attributes = mAttributes
End Property

''
' Sets the attributes for this xml element.
'
' @param RHS An IDictionary of key/value pairs as attributes.
' @remarks The attribute names and values are validated before the list is accepted.
'
Public Property Set Attributes(ByVal RHS As Hashtable)
    If Not RHS Is Nothing Then
        Dim de As DictionaryEntry
        For Each de In RHS
            If Not SecurityElement.IsValidAttributeName(de.Key) Then _
                Throw Cor.NewArgumentException("Invalid attribute name found.")
            If Not SecurityElement.IsValidAttributeValue(de.Value) Then _
                Throw Cor.NewArgumentException("Invalid attribute value found.")
        Next de
    End If

    Set mAttributes = RHS
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
' @param RHS The new tag name.
' @remarks The tag is validated prior to setting.
'
Public Property Let Tag(ByVal RHS As String)
    If Not SecurityElement.IsValidTag(RHS) Then _
        Throw Cor.NewArgumentException("Invalid Tag.", "Tag")
    
    mTag = RHS
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
' @param RHS The new element text.
' @remarks The text is validated prior to being set.
'
Public Property Let Text(ByVal RHS As String)
    If Not SecurityElement.IsValidText(RHS) Then _
        Throw Cor.NewArgumentException("Invalid Text.", "Text")
    
    mText = RHS
End Property

''
' Adds an attribute to the set of attributes for this SecurityElement.
'
' @param Name The name of the new attribute. This is case sensitive.
' @param Value The value of the new attribute.
'
Public Sub AddAttribute(ByVal Name As String, ByVal Value As String)
    If mAttributes Is Nothing Then
        Set mAttributes = New Hashtable
    End If
    
    If Not SecurityElement.IsValidAttributeName(Name) Then _
        Throw Cor.NewArgumentException("Invalid attribute name.", "Name")
    If Not SecurityElement.IsValidAttributeValue(Value) Then _
        Throw Cor.NewArgumentException("Invalid attribute value.", "Value")
    
    mAttributes(Name) = Value
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
' @remarks I so detest the underscore. It was either that or use an abbreviated
' version of the word such as Attr.
'
Public Function GetAttribute(ByVal Name As String) As String
    If mAttributes.ContainsKey(Name) Then
        GetAttribute = mAttributes(Name)
    End If
End Function

''
' Returns a copy of the SecurityElement object.
'
' @return A copy of the object.
'
Public Function Copy() As SecurityElement
    Dim bag As New PropertyBag
    Call bag.WriteProperty("Copy", Me)
    Set Copy = bag.ReadProperty("Copy")
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
    Equals = Object.Equals(Me, Value)
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
    GetHashCode = ObjPtr(CUnk(Me))
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Tag As String, ByVal Text As String)
    mTag = Tag
    mText = Text
End Sub

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
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddAttributes(ByVal sb As StringBuilder)
    If mAttributes Is Nothing Then Exit Sub
    
    Dim de As DictionaryEntry
    For Each de In mAttributes
        Call sb.AppendFormat(" {0}={2}{1}{2}", de.Key, de.Value, Chr$(vbQuoteChar))
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

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
