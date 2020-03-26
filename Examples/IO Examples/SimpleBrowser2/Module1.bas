Attribute VB_Name = "Module1"
Option Explicit

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
' This is a callback function that is supported by such functions
' as the cArray.Sort method and its variations. This function is
' directly called back by VBCorLib, passing a reference to two of
' the array elements that need to be compared. Since the elements
' are all the same in the array being sorted, both parameters in the
' callback must be of the same datatype as the array. Both elements
' must also be declared as ByRef, or horrible things will happen.
' Once the callback has been set up, it works identical to the
' IComparer class used in the SimpleBrowser project.
'
' The purpose of a callback is to allow direct access to the elements
' being sorted, or searched, depending on the function being called.
'
' You pass in the address of this function using the AddressOf operator
' where an IComparer object is requested and allowed.
'
Public Function FileInfoComparer(ByRef File1 As FileInfo, ByRef File2 As FileInfo) As Long
    ' Select the field we are sorting by.
    '
    ' We return a negative value if x is less than y, or a positive value
    ' if x is greater than y. We return zero if x equals y. How x and y are
    ' compared is totally up to the implementor of the callback function.
    Select Case SortColumn
        Case "Name"
            FileInfoComparer = CorString.Compare(File1.Name, File2.Name, OrdinalIgnoreCase)
            
        Case "Size"
            FileInfoComparer = Sgn(File2.Length - File1.Length)
            
        Case "Modified"
            FileInfoComparer = File1.LastAccessTime.CompareTo(File2.LastAccessTime)
    End Select
    
    ' All comparisons expect a return of a negative number if the left (first)
    ' parameter is less than the right (second) parameter, therefore
    ' if we return the opposite, then the sort order will also be the opposite. In
    ' this case, we will create a Descending sort order if we reverse the value.
    If SortOrder = Descending Then FileInfoComparer = -FileInfoComparer
End Function
