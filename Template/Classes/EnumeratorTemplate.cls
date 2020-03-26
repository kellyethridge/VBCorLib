VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumeratorTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This is the base implementation of an enumerator class.
'
' To use this class in a For..Each loop, it must be passed
' into the CreateEnumerator function. The result of that
' function is then returned through the NewEnum function.
'
' Public Function NewEnum() As IUnknown
'     Dim En As New MyEnumerator
'     En.Init Me                           ' Me would be the outer object, such as a list to be enumerated.
'     Set NewEnum = CreateEnumerator(En)
' End Function
'
Option Explicit
Implements IObject
Implements IEnumerator

Private mBase As EnumeratorBase

' Example items to enumerate
Private mItems() As Variant



''
' Returns the current value of the enumeration. This property
' is called automatically if MoveNext succeeds.
'
Public Property Get Current() As Variant
    CopyVariant Current, mItems(mBase.Index)
End Property

''
' Attempts to move the index to the next valid value in the
' object being enumerated. If the move was successful, then
' True is returned, otherwise False is returned.
'
' If True is returned then Current will be called.
'
Public Function MoveNext() As Boolean
    MoveNext = mBase.MoveNext
End Function

''
' Resets the enumerator to the beginning of the object being enumerated.
'
Public Sub Reset()
    mBase.Reset
End Sub

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' This should be overriden for different forms of equality.
' An example would be if this class contained a Name property
' and equality is determined by the names.
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
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Items() As Variant)
    ' Example initialization
    mItems = Items
    Set mBase = NewEnumeratorBase(LBound(Items), CorArray.Length(Items))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    MoveVariant IEnumerator_Current, Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext
End Function

Private Sub IEnumerator_Reset()
    Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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

