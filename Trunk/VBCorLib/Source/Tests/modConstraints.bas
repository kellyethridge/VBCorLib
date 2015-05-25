Attribute VB_Name = "modConstraints"
Option Explicit

Private mObjectComparer As New ObjectComparer

Public Function Equals(ByRef Other As Variant) As IConstraint
    Set Equals = Iz.EqualTo(Other).Using(mObjectComparer)
End Function
