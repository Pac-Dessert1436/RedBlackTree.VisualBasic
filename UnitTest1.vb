Imports NUnit.Framework

''' <summary>
''' Unit tests for the <see cref="RedBlackTree(Of T)"/> class.
''' </summary>
Public NotInheritable Class UnitTest1

    <Test>
    Public Sub Insert_MaintainsSortedTraversal_AndValidRedBlackProperties()
        Dim tree As New RedBlackTree(Of Integer)
        Dim values As Integer() = {10, 5, 15, 3, 7, 12, 18, 1, 4, 6, 8, 11, 13, 17, 19}
        Array.ForEach(values, AddressOf tree.Insert)

        Assert.That(tree.Count, [Is].EqualTo(values.Length))
        Assert.That(tree.IsValidRBT(), [Is].True)
        Assert.That(tree.InOrderSequence.ToArray(), [Is].EqualTo(values.Order().ToArray()))

        For Each val As Integer In values
            Assert.That(tree.Contains(val), [Is].True)
            Assert.That(tree.Find(val), [Is].EqualTo(val))
        Next val
    End Sub

    <Test>
    Public Sub DuplicateInsert_IgnoresDuplicates_AndDoesNotChangeCount()
        Dim tree As New RedBlackTree(Of Integer)
        tree.Insert(5)
        tree.Insert(5)
        tree.Insert(5)

        Assert.That(tree.Count, [Is].EqualTo(1))
        Assert.That(tree.Contains(5), [Is].True)
        Dim elem As Integer() = {5}
        Assert.That(tree.InOrderSequence.ToArray(), [Is].EqualTo(elem))
        Assert.That(tree.IsValidRBT(), [Is].True)
    End Sub

    <Test>
    Public Sub Root_AndMissingLookup_AreHandledCorrectly()
        Dim emptyTree As New RedBlackTree(Of Integer)
        Assert.That(Function() emptyTree.Root, Throws.TypeOf(Of InvalidOperationException)())

        Dim tree As New RedBlackTree(Of Integer)
        tree.Insert(42)

        Assert.That(tree.Root, [Is].EqualTo(42))
        Assert.That(tree.Remove(99), [Is].False)
        Assert.That(tree.Contains(99), [Is].False)
        Assert.That(Function() tree.Find(99), Throws.TypeOf(Of KeyNotFoundException)())
    End Sub

    <Test>
    Public Sub Remove_DeletesExistingValues_AndKeepsTreeBalanced()
        Dim tree As New RedBlackTree(Of Integer)
        Dim allValues As Integer() = Enumerable.Range(1, 30).ToArray()
        Array.ForEach(allValues, AddressOf tree.Insert)

        Dim removedValues As Integer() = {1, 2, 5, 14, 15, 28, 30}
        For Each val As Integer In removedValues
            Assert.That(tree.Remove(val), [Is].True)
        Next val

        Assert.That(tree.Count, [Is].EqualTo(allValues.Length - removedValues.Length))
        Assert.That(tree.IsValidRBT(), [Is].True)
        For Each val As Integer In removedValues
            Assert.That(tree.Contains(val), [Is].False)
        Next val

        Dim remainingValues = allValues.Except(removedValues).Order().ToArray()
        Assert.That(tree.InOrderSequence.ToArray(), [Is].EqualTo(remainingValues))
    End Sub

    <Test>
    Public Sub CustomComparer_WorksWithReverseOrdering()
        Dim comp = Comparer(Of Integer).Create(Function(left, right) right.CompareTo(left))
        Dim tree As New RedBlackTree(Of Integer)(comp)
        Dim values As Integer() = {5, 1, 4, 2, 3}
        Array.ForEach(values, AddressOf tree.Insert)

        Assert.That(tree.IsValidRBT(), [Is].True)
        Dim arr As Integer() = {5, 4, 3, 2, 1}
        Assert.That(tree.InOrderSequence.ToArray(), [Is].EqualTo(arr))
    End Sub

    <Test>
    Public Sub Enumerator_IteratesInOrder()
        Dim tree As New RedBlackTree(Of Integer)
        Dim values = {9, 2, 7, 1, 8, 3, 4, 6, 5}
        Array.ForEach(values, AddressOf tree.Insert)
        Assert.That(tree.ToArray(), [Is].EqualTo(values.Order().ToArray()))
    End Sub
End Class
