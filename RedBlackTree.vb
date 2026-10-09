Option Explicit On
Option Strict On
Option Infer On
Option Compare Binary

''' <summary>
''' A generic Red-Black Tree implementation for VB.NET.
''' </summary>
Public NotInheritable Class RedBlackTree(Of T As IComparable(Of T))
    Implements IEnumerable(Of T)

    Private Enum NodeColor As Byte
        Red = 0
        Black = 1
    End Enum

    Private NotInheritable Class RBTNode
        Public Property Value As T
        Public Property Color As NodeColor
        Public Property Left As RBTNode
        Public Property Right As RBTNode
        Public Property Parent As RBTNode

        Public Sub New(value As T)
            Me.Value = value
            Color = NodeColor.Red
        End Sub
    End Class

    Private _root As RBTNode
    Private _count As Integer
    Private ReadOnly _comparer As IComparer(Of T)

    ''' <summary>
    ''' Initializes a new instance of the <see cref="RedBlackTree(Of T)"/> class.
    ''' </summary>
    Public Sub New()
        _comparer = Comparer(Of T).Default
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="RedBlackTree(Of T)"/> class with a custom comparer.
    ''' </summary>
    ''' <param name="comparer">The custom comparer to use for comparison.</param>
    Public Sub New(comparer As IComparer(Of T))
        ArgumentNullException.ThrowIfNull(comparer)
        _comparer = comparer
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return _count
        End Get
    End Property

    Public ReadOnly Property Root As T
        Get
            If _root Is Nothing Then Throw New InvalidOperationException("Tree is empty.")
            Return _root.Value
        End Get
    End Property

    ''' <summary>
    ''' Inserts a new value into the Red-Black Tree.
    ''' </summary>
    ''' <param name="value">The value to insert.</param>
    Public Sub Insert(value As T)
        ArgumentNullException.ThrowIfNull(value)

        If _root Is Nothing Then
            _root = New RBTNode(value) With {
                .Color = NodeColor.Black
            }
            _count = 1
            Exit Sub
        End If

        Dim current As RBTNode = _root
        Dim parent As RBTNode = Nothing

        While current IsNot Nothing
            parent = current
            Dim comparison As Integer = _comparer.Compare(value, current.Value)
            If comparison < 0 Then
                current = current.Left
            ElseIf comparison > 0 Then
                current = current.Right
            Else
                Exit Sub
            End If
        End While

        Dim newNode As New RBTNode(value) With {.Parent = parent}
        If _comparer.Compare(value, parent.Value) < 0 Then
            parent.Left = newNode
        Else
            parent.Right = newNode
        End If

        _count += 1
        FixInsert(newNode)
    End Sub

    Private Sub FixInsert(node As RBTNode)
        Do While node.Parent IsNot Nothing AndAlso node.Parent.Color = NodeColor.Red
            Dim parent As RBTNode = node.Parent
            Dim grandParent As RBTNode = parent.Parent
            If grandParent Is Nothing Then Exit Do

            If parent Is grandParent.Left Then
                Dim uncle As RBTNode = grandParent.Right

                If uncle IsNot Nothing AndAlso uncle.Color = NodeColor.Red Then
                    parent.Color = NodeColor.Black
                    uncle.Color = NodeColor.Black
                    grandParent.Color = NodeColor.Red
                    node = grandParent
                Else
                    If node Is parent.Right Then
                        node = parent
                        RotateLeft(node)
                    End If

                    parent = node.Parent
                    If parent Is Nothing Then Exit Do
                    grandParent = parent.Parent
                    If grandParent Is Nothing Then Exit Do

                    parent.Color = NodeColor.Black
                    grandParent.Color = NodeColor.Red
                    RotateRight(grandParent)
                End If
            Else
                Dim uncle As RBTNode = grandParent.Left

                If uncle IsNot Nothing AndAlso uncle.Color = NodeColor.Red Then
                    parent.Color = NodeColor.Black
                    uncle.Color = NodeColor.Black
                    grandParent.Color = NodeColor.Red
                    node = grandParent
                Else
                    If node Is parent.Left Then
                        node = parent
                        RotateRight(node)
                    End If

                    parent = node.Parent
                    If parent Is Nothing Then Exit Do
                    grandParent = parent.Parent
                    If grandParent Is Nothing Then Exit Do

                    parent.Color = NodeColor.Black
                    grandParent.Color = NodeColor.Red
                    RotateLeft(grandParent)
                End If
            End If
        Loop

        _root.Color = NodeColor.Black
    End Sub

    Private Sub RotateLeft(x As RBTNode)
        Dim y As RBTNode = x.Right
        If y Is Nothing Then Exit Sub

        x.Right = y.Left
        If y.Left IsNot Nothing Then y.Left.Parent = x

        y.Parent = x.Parent

        If x.Parent Is Nothing Then
            _root = y
        ElseIf x Is x.Parent.Left Then
            x.Parent.Left = y
        Else
            x.Parent.Right = y
        End If

        y.Left = x
        x.Parent = y
    End Sub

    Private Sub RotateRight(x As RBTNode)
        Dim y As RBTNode = x.Left
        If y Is Nothing Then Exit Sub

        x.Left = y.Right
        If y.Right IsNot Nothing Then y.Right.Parent = x

        y.Parent = x.Parent

        If x.Parent Is Nothing Then
            _root = y
        ElseIf x Is x.Parent.Right Then
            x.Parent.Right = y
        Else
            x.Parent.Left = y
        End If

        y.Right = x
        x.Parent = y
    End Sub

    ''' <summary>
    ''' Determines whether the Red-Black Tree contains a specific value.
    ''' </summary>
    ''' <param name="value">The value to search for.</param>
    ''' <returns>True if the value is found; otherwise, false.</returns>
    Public Function Contains(value As T) As Boolean
        Return FindNode(value) IsNot Nothing
    End Function

    ''' <summary>
    ''' Finds the value in the Red-Black Tree.
    ''' </summary>
    ''' <param name="value">The value to search for.</param>
    ''' <returns>The value if found; otherwise, throws a <see cref="KeyNotFoundException"/>.</returns>
    Public Function Find(value As T) As T
        Dim node As RBTNode = FindNode(value)
        If node Is Nothing Then Throw New KeyNotFoundException("Value not found.")
        Return node.Value
    End Function

    ''' <summary>
    ''' Finds the node in the Red-Black Tree.
    ''' </summary>
    ''' <param name="value">The value to search for.</param>
    ''' <returns>The node if found; otherwise, returns Nothing.</returns>
    Private Function FindNode(value As T) As RBTNode
        Dim current As RBTNode = _root
        While current IsNot Nothing
            Dim comparison As Integer = _comparer.Compare(value, current.Value)
            If comparison = 0 Then Return current
            current = If(comparison < 0, current.Left, current.Right)
        End While
        Return Nothing
    End Function

    ''' <summary>
    ''' Removes a value from the Red-Black Tree.
    ''' </summary>
    ''' <param name="value">The value to remove.</param>
    ''' <returns>True if the value is found and removed; otherwise, false.</returns>
    Public Function Remove(value As T) As Boolean
        Dim target As RBTNode = FindNode(value)
        If target Is Nothing Then Return False

        DeleteNode(target)
        _count -= 1
        Return True
    End Function

    Private Shared Function Minimum(node As RBTNode) As RBTNode
        While node.Left IsNot Nothing
            node = node.Left
        End While
        Return node
    End Function

    Private Sub DeleteNode(node As RBTNode)
        Dim replacement As RBTNode
        Dim replacementParent As RBTNode
        Dim originalColor As NodeColor = node.Color

        If node.Left Is Nothing Then
            replacement = node.Right
            replacementParent = node.Parent
            Transplant(node, replacement)
        ElseIf node.Right Is Nothing Then
            replacement = node.Left
            replacementParent = node.Parent
            Transplant(node, replacement)
        Else
            Dim successor As RBTNode = Minimum(node.Right)
            originalColor = successor.Color
            replacement = successor.Right

            If successor.Parent Is node Then
                replacementParent = successor
                If replacement IsNot Nothing Then replacement.Parent = successor
            Else
                replacementParent = successor.Parent
                Transplant(successor, successor.Right)
                successor.Right = node.Right
                If successor.Right IsNot Nothing Then successor.Right.Parent = successor
            End If

            Transplant(node, successor)
            successor.Left = node.Left
            If successor.Left IsNot Nothing Then successor.Left.Parent = successor
            successor.Color = node.Color
        End If

        If originalColor = NodeColor.Black Then FixDelete(replacement, replacementParent)
    End Sub

    Private Sub FixDelete(x As RBTNode, parent As RBTNode)
        Do While x IsNot _root AndAlso (x Is Nothing OrElse x.Color = NodeColor.Black)
            If parent Is Nothing Then Exit Do

            If x Is parent.Left Then
                Dim sibling As RBTNode = parent.Right

                If sibling IsNot Nothing AndAlso sibling.Color = NodeColor.Red Then
                    sibling.Color = NodeColor.Black
                    parent.Color = NodeColor.Red
                    RotateLeft(parent)
                    sibling = parent.Right
                End If

                If sibling Is Nothing Then
                    x = parent
                    parent = x?.Parent
                    Continue Do
                End If

                If (sibling.Left Is Nothing OrElse sibling.Left.Color = NodeColor.Black) AndAlso
                   (sibling.Right Is Nothing OrElse sibling.Right.Color = NodeColor.Black) Then
                    sibling.Color = NodeColor.Red
                    x = parent
                    parent = x?.Parent
                Else
                    If sibling.Right Is Nothing OrElse sibling.Right.Color = NodeColor.Black Then
                        If sibling.Left IsNot Nothing Then sibling.Left.Color = NodeColor.Black
                        sibling.Color = NodeColor.Red
                        RotateRight(sibling)
                        sibling = parent.Right
                    End If

                    sibling.Color = parent.Color
                    parent.Color = NodeColor.Black
                    If sibling.Right IsNot Nothing Then sibling.Right.Color = NodeColor.Black
                    RotateLeft(parent)
                    x = _root
                    parent = Nothing
                End If
            Else
                Dim sibling As RBTNode = parent.Left

                If sibling IsNot Nothing AndAlso sibling.Color = NodeColor.Red Then
                    sibling.Color = NodeColor.Black
                    parent.Color = NodeColor.Red
                    RotateRight(parent)
                    sibling = parent.Left
                End If

                If sibling Is Nothing Then
                    x = parent
                    parent = x?.Parent
                    Continue Do
                End If

                If (sibling.Left Is Nothing OrElse sibling.Left.Color = NodeColor.Black) AndAlso
                   (sibling.Right Is Nothing OrElse sibling.Right.Color = NodeColor.Black) Then
                    sibling.Color = NodeColor.Red
                    x = parent
                    parent = x?.Parent
                Else
                    If sibling.Left Is Nothing OrElse sibling.Left.Color = NodeColor.Black Then
                        If sibling.Right IsNot Nothing Then sibling.Right.Color = NodeColor.Black
                        sibling.Color = NodeColor.Red
                        RotateLeft(sibling)
                        sibling = parent.Left
                    End If

                    sibling.Color = parent.Color
                    parent.Color = NodeColor.Black
                    If sibling.Left IsNot Nothing Then sibling.Left.Color = NodeColor.Black
                    RotateRight(parent)
                    x = _root
                    parent = Nothing
                End If
            End If
        Loop

        If x IsNot Nothing Then x.Color = NodeColor.Black
    End Sub

    Private Sub Transplant(u As RBTNode, v As RBTNode)
        If u.Parent Is Nothing Then
            _root = v
        ElseIf u Is u.Parent.Left Then
            u.Parent.Left = v
        Else
            u.Parent.Right = v
        End If

        If v IsNot Nothing Then v.Parent = u.Parent
    End Sub

    ''' <summary>
    ''' Gets the in-order sequence of values in the Red-Black Tree.
    ''' </summary>
    ''' <returns>An enumerable collection of values in in-order sequence.</returns>
    Public ReadOnly Property InOrderSequence As IEnumerable(Of T)
        Get
            Dim result As New List(Of T)
            InOrderWalk(_root, result)
            Return result
        End Get
    End Property

    Private Shared Sub InOrderWalk(node As RBTNode, acc As List(Of T))
        If node Is Nothing Then Exit Sub
        InOrderWalk(node.Left, acc)
        acc.Add(node.Value)
        InOrderWalk(node.Right, acc)
    End Sub

    ''' <summary>
    ''' Returns an enumerator that iterates through the enumerable collection.
    ''' </summary>
    ''' <returns>An enumerator that can be used to iterate through the enumerable collection.</returns>
    Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
        Return InOrderSequence.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function

    ''' <summary>
    ''' Determines whether the Red-Black Tree is valid.
    ''' </summary>
    ''' <returns>True if the Red-Black Tree is valid; otherwise, false.</returns>
    Public Function IsValidRBT() As Boolean
        If _root Is Nothing Then Return True
        If _root.Color <> NodeColor.Black Then Return False
        If _root.Parent IsNot Nothing Then Return False

        Dim blackHeight As Integer = -1
        Return ValidateRBTNode(_root, 0, blackHeight)
    End Function

    Private Shared Function ValidateRBTNode(node As RBTNode, currentBlack As Integer, ByRef expectedBlack As Integer) As Boolean
        If node Is Nothing Then
            If expectedBlack = -1 Then
                expectedBlack = currentBlack
                Return True
            End If
            Return currentBlack = expectedBlack
        End If

        If node.Color = NodeColor.Red Then
            If (node.Left IsNot Nothing AndAlso node.Left.Color = NodeColor.Red) OrElse
               (node.Right IsNot Nothing AndAlso node.Right.Color = NodeColor.Red) Then
                Return False
            End If
        End If

        If node.Left IsNot Nothing AndAlso node.Left.Parent IsNot node Then Return False
        If node.Right IsNot Nothing AndAlso node.Right.Parent IsNot node Then Return False

        Dim nextBlack As Integer = currentBlack + If(node.Color = NodeColor.Black, 1, 0)
        Return ValidateRBTNode(node.Left, nextBlack, expectedBlack) AndAlso
               ValidateRBTNode(node.Right, nextBlack, expectedBlack)
    End Function
End Class