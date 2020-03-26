VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple INI File Browser"
   ClientHeight    =   8715
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   5880
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8295
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   14631
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox List1 
      Height          =   8250
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Sections"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This simply demonstrates viewing an INI file's sections
' and key-value pairs within a section.
'
Option Explicit
Private mFile As INIFile

''
' Init the form.
Private Sub Form_Load()
    With ListView1
        .ColumnHeaders.Add , , "Key", 2500
        .ColumnHeaders.Add , , "Value", 2500
    End With
End Sub

''
' When a Section is selected from the list, we
' display the key-value pairs in that Section.
'
' On NT machines, some values come from the Registry, aswell.
'
Private Sub List1_Click()
    ' Get an IDictionary object containing all
    ' of the key-value pairs in a Section.
    Dim Values As IDictionary
    Set Values = mFile.GetValues(List1.Text)
    
    ListView1.ListItems.Clear
    
    ' Sort the values
    Dim SortedEntries As ArrayList
    Set SortedEntries = SortDictionary(Values)
    
    ' Iterate over each entry and add the
    ' Key and Value for each entry to the list.
    Dim Entry As DictionaryEntry
    For Each Entry In SortedEntries
        Dim Item As ListItem
        Set Item = ListView1.ListItems.Add(, , Entry.Key)
        Item.SubItems(1) = Entry.Value
    Next Entry
End Sub

''
' Selects a file to be opened.
Private Sub mnuFileOpen_Click()
    On Error GoTo errTrap
    With CD
        .CancelError = True
        .Filter = "INI Files|*.ini|All Files|*.*"
        .ShowOpen
        Set mFile = NewINIFile(.FileName)
        Caption = "Simple INI File Browser - " & mFile.FileName
        
        ' And now show the sections in the INI file.
        ShowSections
        
        ListView1.ListItems.Clear
    End With
errTrap:
End Sub

''
' This will simply diplay a list of all the INI sections in the file.
'
Private Sub ShowSections()
    ' Get all of the Section names in the file.
    Dim Sections() As String
    Sections = mFile.GetSectionNames
    
    ' Sort the section names with a case-insensitive string comparer.
    CorArray.Sort Sections, CaseInsensitiveComparer.Default
    
    List1.Clear
    
    ' Simply add them to the list.
    Dim i As Long
    For i = 0 To UBound(Sections)
        List1.AddItem Sections(i)
    Next i
End Sub

''
' Returns the IDictionary contents in a sorted list.
'
Private Function SortDictionary(ByVal Dictionary As IDictionary) As ArrayList
    Dim ret As New ArrayList
    
    ' We can use AddRange because any class that implements
    ' the IDictionary interface must (should) also implement the
    ' ICollection interface. Well, you could leave the ICollection
    ' interface out, but then consistency would be lost.
    ret.AddRange Dictionary
    
    ' Ok, how can we sort them now?
    ' An IDictionary object contains DictionaryEntry objects,
    ' and those are what was added to the ArrayList.
    '
    ' This looks like a job for a custom comparer to allow the
    ' sorting routine in the ArrayList to perform properly.
    ret.Sort Comparer:=New INIEntryComparer
    
    Set SortDictionary = ret
End Function
