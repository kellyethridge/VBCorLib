VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorEqualsConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IConstraint

Private mExpected   As Variant
Private mActual     As Variant

Friend Sub Init(ByRef Expected As Variant)
    CopyVariant mExpected, Expected
End Sub

Private Function IConstraint_Matches(Actual As Variant) As Boolean
    If Object.Equals(Actual, mExpected) Then
        IConstraint_Matches = True
    Else
        CopyVariant mActual, Actual
    End If
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As SimplyVBComp.TextMessageWriter)
    Writer.WriteActualValue Object.ToString(mActual)
End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As SimplyVBComp.TextMessageWriter)
    Writer.WriteExpectedValue Object.ToString(mExpected)
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As SimplyVBComp.TextMessageWriter)
    Writer.DisplayConstraintDifferences Me
End Sub
