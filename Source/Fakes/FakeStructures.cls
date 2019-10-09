VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeStructures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type SimpleStructure
    Value As Long
End Type

Public Type ComplexStructure
    Value As Long
    FixedLong(0) As Long
    Inner As SimpleStructure
End Type
    
Public Type StringStructure
    Value As String
End Type

Public Type InvalidComplexStructure
    Value As Long
    Inner As StringStructure
End Type

