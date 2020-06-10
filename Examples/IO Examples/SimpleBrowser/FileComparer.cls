VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' This is a typical implementation for a custom comparer object.
' This comparer compares a specific aspect of the FileInfo objects.
' By setting the properties, we can make this comparer perform
' specific comparisons depending on the situation.
'
Option Explicit
Implements IComparer

Public Enum SortDirection
    Ascending
    Descending
End Enum

' This is the name of the column we are sorting by.
Public SortColumn As String

' This is the direction of the sort we want. We don't
' always have to go in Ascending order. We can make
' it compare however we choose to, as long as we are
' consistent for the entire sort process.
Public SortOrder As SortDirection


''
' This is called by any function that supports an IComparer object.
'
' x and y are two elements from a list being sorted. They are passed
' as variants because the interface must be generic, so we cast them
' to specific datatypes within the funtion.
'
Private Function IComparer_Compare(x As Variant, y As Variant) As Long
    Dim File1 As FileInfo
    Dim File2 As FileInfo
    
    ' Cast the two elements to a datatype we can query easily.
    Set File1 = x
    Set File2 = y
    
    ' Select the field we are sorting by.
    '
    ' We return a negative value if x is less than y, or a positive value
    ' if x is greater than y. We return zero if x equals y. How x and y are
    ' compared is totally up to the implementor of the IComparer class.
    Select Case SortColumn
        Case "Name"
            IComparer_Compare = CorString.Compare(File1.Name, File2.Name, OrdinalIgnoreCase)
            
        Case "Size"
            IComparer_Compare = Sgn(File2.Length - File1.Length)
            
        Case "Modified"
            IComparer_Compare = File1.LastAccessTime.CompareTo(File2.LastAccessTime)
    End Select
    
    ' All comparisons expect a return of a negative number if the left (first)
    ' parameter is less than the right (second) parameter, therefore
    ' if we return the opposite, then the sort order will also be the opposite. In
    ' this case, we will create a Descending sort order if we reverse the value.
    If SortOrder = Descending Then IComparer_Compare = -IComparer_Compare
End Function
