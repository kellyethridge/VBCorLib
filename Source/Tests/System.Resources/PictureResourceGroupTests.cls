VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureResourceGroupTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' PictureResourceGroupTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Add_WithNothing_ThrowsArgumentNullException()
    Dim Group As PictureResourceGroup
    Set Group = Cor.NewPictureResourceGroup(CursorGroup)
    
    On Error Resume Next
    Group.Add Nothing
    
    AssertArgumentNullException Err, "Resource"
End Sub

Public Sub Add_WhenCursorGroupWithIcon_ThrowsArgumentException()
    Dim Group As PictureResourceGroup
    Set Group = Cor.NewPictureResourceGroup(CursorGroup)
    
    On Error Resume Next
    Group.Add Cor.NewPictureResourceInfo(1, LoadIcon)

    AssertArgumentException Err, "Resource"
End Sub

Public Sub Add_WhenIconGroupWithCursor_ThrowsArgumentException()
    Dim Group As PictureResourceGroup
    Set Group = Cor.NewPictureResourceGroup(IconGroup)
    
    On Error Resume Next
    Group.Add Cor.NewPictureResourceInfo(1, LoadCursor)

    AssertArgumentException Err, "Resource"
End Sub
