Attribute VB_Name = "modConstraints"
Option Explicit

Private mObjectComparer As New ObjectComparer

Public Function Equals(ByVal Other As IObject) As IConstraint
    Set Equals = Iz.EqualTo(Other).Using(mObjectComparer)
End Function
