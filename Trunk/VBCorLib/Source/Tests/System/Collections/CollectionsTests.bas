Attribute VB_Name = "CollectionsTests"
Option Explicit

Public Function CreateCollectionsTests() As TestSuite
    With Sim.NewTestSuite("System.Collections")
        .Add New BitArrayTests
        .Add New TestSortedList
        .Add New ComparerTests
        .Add New TestCaseInsensitiveHCP
        
        AddArrayListTestsTo .This
        AddQueueTestsTo .This
        AddStackTestsTo .This
        AddHashtableTestsTo .This
    
        Set CreateCollectionsTests = .This
    End With
End Function

Private Sub AddArrayListTestsTo(ByVal Suite As TestSuite)
    With Sim.NewTestSuite("ArrayList")
        .Add New ArrayListTests
        .Add New ArrayListAdapterTests
        .Add New ArrayListRangedTests
        .Add New ArrayListRepeatTests
        .Add New ArrayListEnumeratorTests
        .Add New ArrayListReadOnlyTests
        Suite.Add .This
    End With
End Sub

Private Sub AddQueueTestsTo(ByVal Suite As TestSuite)
    With Sim.NewTestSuite("Queue")
        .Add New QueueTests
        .Add New QueueEnumeratorTests
        Suite.Add .This
    End With
End Sub

Private Sub AddStackTestsTo(ByVal Suite As TestSuite)
    With Sim.NewTestSuite("Stack")
        .Add New StackTests
        .Add New StackEnumeratorTests
        Suite.Add .This
    End With
End Sub

Private Sub AddHashtableTestsTo(ByVal Suite As TestSuite)
    With Sim.NewTestSuite("Hashtable")
        .Add New TestHashTable
        .Add New TestHashTableHCP
        .Add New TestDictionaryEntry
        Suite.Add .This
    End With
End Sub
