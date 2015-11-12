Attribute VB_Name = "ConstraintsHelper"
Option Explicit

Public Function Equals(ByRef Expected As Variant) As CorEqualsConstraint
    Set Equals = New CorEqualsConstraint
    Equals.Init Expected
End Function


