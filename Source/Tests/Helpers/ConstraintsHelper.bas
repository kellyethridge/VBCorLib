Attribute VB_Name = "ConstraintsHelper"
Option Explicit

Public Function Equals(ByRef Expected As Variant) As IConstraint
    Dim Constraint As New CorEqualsConstraint
    Constraint.Init Expected
    
    Set Equals = Constraint
End Function

Public Function NotEquals(ByRef Expected As Variant) As IConstraint
    Set NotEquals = Sim.NewNotConstraint(Equals(Expected))
End Function
