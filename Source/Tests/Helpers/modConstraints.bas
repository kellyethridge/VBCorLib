Attribute VB_Name = "modConstraints"
Option Explicit

Public Function Equals(ByRef Expected As Variant) As CorEqualsConstraint
    Set Equals = New CorEqualsConstraint
    Equals.Init Expected
End Function
