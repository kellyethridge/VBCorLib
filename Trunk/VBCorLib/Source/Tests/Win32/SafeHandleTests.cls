VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SafeHandleTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SafeHandleTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "Microsoft.Win32.SafeHandles"
End Sub

Public Sub NewSafeHandle_WhenCalled_CreatesAnOpenHandle()
    Dim SUT As SafeHandle
    
    Set SUT = Cor.NewSafeHandle(1, True)
    
    Assert.That SUT.DangerousGetHandle, Iz.EqualTo(1)
    Assert.That SUT.IsClosed, Iz.False
    Assert.That SUT.IsInvalid, Iz.False
End Sub

Public Sub CloseHandle_WhenClosed_NotifiesSubclass()
    Dim SUT As New FakeSafeHandle
    SUT.Init 2, True
    
    SUT.CloseHandle
    
    Assert.That SUT.ReleaseHandleCallCount, Iz.EqualTo(1)
    Assert.That SUT.IsClosed, Iz.True
End Sub

Public Sub CloseHandle_WhenNotHandleOwner_DoesNotReleaseHandle()
    Dim SUT As New FakeSafeHandle
    SUT.Init 2, False
    
    SUT.CloseHandle
    
    Assert.That SUT.ReleaseHandleCallCount, Iz.EqualTo(0)
    Assert.That SUT.IsClosed, Iz.False
End Sub

