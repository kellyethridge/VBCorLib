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

Private Const PropTag          As String = "Tag"
Private Const PropText         As String = "Text"
Private Const PropChildren     As String = "Children"
Private Const PropAttributes   As String = "Attributes"

Private mTag        As String
Private mText       As String
Private mChildren   As ArrayList
Private mAttributes As ArrayList


''
' Returns any attributes contained within the xml element.
'
' @return An IDictionary object containing key/value pairs of attributes or Nothing.
' @remarks The attribute names are case sensitive.
'
Public Property Get Attributes() As Hashtable
    If Not mAttributes Is Nothing Then
        Set Attributes = Cor.NewHashtable2(mAttributes.Count)
        
        Dim i As Long
        For i = 0 To mAttributes.Count - 1 Step 2
            Attributes.Add mAttributes(i), mAttributes(i + 1)
        Next
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
        Dim Item As DictionaryEntry
        Dim List As New ArrayList
        
        For Each Item In Value
            If Not SecurityElement.IsValidAttributeName(Item.Key) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementName, Item.Key))
            If Not SecurityElement.IsValidAttributeValue(Item.Value) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementValue, Item.Value))
                
            List.Add Item.Key
            List.Add Item.Value
        Next
        
        Set mAttributes = List
    End If
End Property

Friend Property Get InternalAttributes() As ArrayList
    Set InternalAttributes = mAttributes
End Property

Friend Property Set InternalAttributes(ByVal Value As ArrayList)
    Set mAttributes = Value
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
' @param Value The list of children to be set to this SecurityElement.
'
Public Property Set Children(ByVal Value As ArrayList)
    If Not Value Is Nothing Then
        Dim Item As Variant
        For Each Item In Value
            If IsObject(Item) Then
                If Item Is Nothing Then
                    Error.Argument ArgumentNull_Child
                End If
            End If
        Next
    End If
    
    Set mChildren = Value
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
    Text = Unescape(mText)
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
        Set mAttributes = New ArrayList
    End If
    
    If Not SecurityElement.IsValidAttributeName(Name) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementName, Name))
    If Not SecurityElement.IsValidAttributeValue(Value) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidElementValue, Value))
    
    Dim i As Long
    For i = 0 To mAttributes.Count - 1 Step 2
        If mAttributes(i) = Name Then
            Error.Argument Argument_AttributeNamesMustBeUnique
        End If
    Next
        
    mAttributes.Add Name
    mAttributes.Add Value
End Sub

''
' Adds a new child to the SecurityElement.
'
' @param Child The child to be added.
'
Public Sub AddChild(ByVal Child As SecurityElement)
    If Child Is Nothing Then _
        Error.ArgumentNull "Child"
    
    If mChildren Is Nothing Then
        Set mChildren = New ArrayList
    End If
    
    mChildren.Add Child
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
        Dim i As Long
        For i = 0 To mAttributes.Count - 1 Step 2
            If mAttributes(i) = Name Then
                Dim Result As String
                Result = mAttributes(i + 1)
                
                If InStr(Result, "&") Then
                    GetAttribute = Unescape(Result)
                Else
                    GetAttribute = Result
                End If
            End If
        Next
    End If
End Function

''
' Returns a copy of the SecurityElement object.
'
' @return A copy of the object.
'
Public Function Copy() As SecurityElement
    Set Copy = New SecurityElement
    Copy.InitCopy Me
End Function

''
' Searches for a child element by tag name.
'
' @param Tag The tag name of the child element.
' @return Returns the child element if found, or Nothing if not.
' @remarks Only the first level of children for this instance are searched,
' no grandchildren are searched.
'
Public Function SearchForChildByTag(ByRef Tag As String) As SecurityElement
    If mChildren Is Nothing Then
        Exit Function
    End If
    
    Dim Child As SecurityElement
    For Each Child In mChildren
        If Child.Tag = Tag Then
            Set SearchForChildByTag = Child
            Exit Function
        End If
    Next
End Function

''
' Searches for the text of a child with the specified tag name.
'
' @param Tag The tag name of the child being searched for.
' @Return Returns the text for the found child, or an empty string if not found.
' @remarks Only the first level of children for this instance are searched,
' no grandchildren are searched.
'
Public Function SearchForTextOfTag(ByRef Tag As String) As String
    If Tag = mTag Then
        SearchForTextOfTag = Unescape(mText)
        Exit Function
    End If
    
    Dim Child As SecurityElement
    Set Child = FindChildByTag(Tag)
    
    If Not Child Is Nothing Then
        SearchForTextOfTag = Child.Text
    End If
End Function

Friend Function FindChildByTag(ByRef Tag As String) As SecurityElement
    If mChildren Is Nothing Then
        Exit Function
    End If
        
    Dim Result As SecurityElement
    Dim Child As SecurityElement
    For Each Child In mChildren
        If Child.Tag = Tag Then
            Set FindChildByTag = Child
            Exit Function
        End If
        
        Set Result = Child.FindChildByTag(Tag)
        If Not Result Is Nothing Then
            Set FindChildByTag = Result
            Exit Function
        End If
    Next
End Function

''
' Compares two SecurityElement instances for equality.
'
' @param Other The instance to compare equality to.
' @return Returns if both instances contain the same properties, the
' same attributes (in the same order) and equal children (in the same order).
'
Public Function Equal(ByVal Other As SecurityElement) As Boolean
    If Other Is Nothing Then _
        Exit Function
    If mTag <> Other.Tag Then _
        Exit Function
    If mText <> Other.Text Then _
        Exit Function
    If Not EqualAttributes(Other) Then _
        Exit Function
    If Not EqualChildren(Other) Then _
        Exit Function
    
    Equal = True
End Function

Private Function EqualAttributes(ByVal Other As SecurityElement) As Boolean
    If Other.InternalAttributes Is Nothing Then
        If Not mAttributes Is Nothing Then
            Exit Function
        End If
    ElseIf mAttributes Is Nothing Then
        Exit Function
    ElseIf Other.InternalAttributes.Count <> mAttributes.Count Then
        Exit Function
    Else
        With Other.InternalAttributes
            Dim i As Long
            For i = 0 To mAttributes.Count - 1
                If mAttributes(i) <> .Item(i) Then
                    Exit Function
                End If
            Next
        End With
    End If
    
    EqualAttributes = True
End Function

Private Function EqualChildren(ByVal Other As SecurityElement) As Boolean
    If Other.Children Is Nothing Then
        If Not mChildren Is Nothing Then
            Exit Function
        End If
    ElseIf mChildren Is Nothing Then
        Exit Function
    ElseIf Other.Children.Count <> mChildren.Count Then
        Exit Function
    Else
        With Other.Children
            Dim i As Long
            Dim Child As SecurityElement
            For i = 0 To mChildren.Count - 1
                Set Child = mChildren(i)
                If Not Child.Equal(.Item(i)) Then
                    Exit Function
                End If
            Next
        End With
    End If

    EqualChildren = True
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
    Dim sb As StringBuilder
    Set sb = StringBuilderCache.Acquire
    
    WriteTo sb
    
    ToString = StringBuilderCache.GetStringAndRelease(sb)
End Function

Friend Sub WriteTo(ByVal sb As StringBuilder)
    sb.AppendFormat "<{0}", mTag
    AddAttributes sb
    
    If HasChildren Or HasText Then
        sb.Append ">"
        sb.Append mText
        
        If HasChildren Then
            sb.AppendLine
            AddChildren sb
        End If
        
        sb.AppendFormat "</{0}>{1}", mTag, vbCrLf
    Else
        sb.AppendLine "/>"
    End If
End Sub

Private Sub AddAttributes(ByVal sb As StringBuilder)
    Dim IsFirst As Boolean
    Dim i       As Long
    
    If mAttributes Is Nothing Then
        Exit Sub
    End If
    
    IsFirst = True
    For i = 0 To mAttributes.Count - 1 Step 2
        If IsFirst Then
            sb.AppendFormat " {0}=""{1}""", mAttributes(i), mAttributes(i + 1)
            IsFirst = False
        Else
            sb.AppendFormat "{0}{1}=""{2}""", vbCrLf, mAttributes(i), mAttributes(i + 1)
        End If
    Next
End Sub

Private Sub AddChildren(ByVal sb As StringBuilder)
    Dim Child As SecurityElement
    For Each Child In mChildren
        Child.WriteTo sb
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

Friend Sub InitCopy(ByVal Source As SecurityElement)
    mTag = Source.Tag
    mText = Source.Text
    Set Attributes = Source.Attributes
    If Not Source.Children Is Nothing Then
        Set mChildren = Cor.NewArrayList(Source:=Source.Children)
    End If
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mTag = .ReadProperty(PropTag)
        mText = .ReadProperty(PropText)
        Set mChildren = .ReadProperty(PropChildren, Nothing)
        Set mAttributes = .ReadProperty(PropAttributes, Nothing)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropTag, mTag
        .WriteProperty PropText, mText
        .WriteProperty PropChildren, mChildren, Nothing
        .WriteProperty PropAttributes, mAttributes, Nothing
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
