Attribute VB_Name = "PersistenceHelper"
Option Explicit

Public Function Persist(ByVal Source As Object) As Object
    Dim Serializer      As New PropertyBag
    Dim Deserializer    As New PropertyBag
    
    On Error GoTo SerializationError
    Serializer.WriteProperty "SUT", Source, Nothing
    
    Deserializer.Contents = Serializer.Contents
    
    On Error GoTo DeserializationError
    Set Persist = Deserializer.ReadProperty("SUT", Nothing)
    Exit Function
    
SerializationError:
    Assert.Fail "'" & TypeName(Source) & "' failed to serialize: " & Err.Description
    Exit Function
    
DeserializationError:
    Assert.Fail "'" & TypeName(Source) & "' failed to deserialize: " & Err.Description
End Function


