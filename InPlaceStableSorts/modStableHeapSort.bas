Attribute VB_Name = "modStableHeapSort"
Option Explicit

'Stable Heap Sort
'----------------
'
'Summary
'-------
'
'Stable Heap Sort works in a similar manner to traditional heap sort.
'There are some differences:
'- The heaps are arranged so that the parent of a heap is the least value of the key.
'  Where there are equal key values, then the parent is also the earlier value in natural order.
'  The mapping of the heaps is different to traditional heap sort and supports this by default.
'- When pushing a value down the heaps, the left side of the heap must be processed before the
'  right hand side.  This sometimes means that both left and right sides of a heap must be processed.
'- When pushing a value down to the right hand child, any sequence of equal values must be rotated
'  so that the last of the equal values is brought to the top and all others are shuffled down one.
'  It is then the last of the equal values that gets pushed down to the right child.
'- When sorted items are removed from the heap, they are taken from the top.  The subsequent orphan
'  elements then become part of a virtual heap.
'
'Performance
'-----------
'
'While promising to be NLogN like Heap Sort, the rebuilding process sometimes needs to rebuild
'both children.  This is an exponential process and the combination of 2^N and Log2(N) = N.
'As a result this sort is of order O(N^2).
'
'Concerns over large numbers of equivalent elements causing extra load via the rotation requirements
'are unfounded with any extra load being more than compensated from fewer swapping operations.  In fact
'performance improves where there are large numbers of equivalent elements.
'
'Similar to Smooth Sort, performance becomes order O(N) when the data is already sorted.
'
'Heap Structure
'--------------
'
'The heap is arranged so that the indexes into the array go deep before they go across.
'Unlike heap-sort where the indexes go across instead of deep
'
'Stable Heap Sort:                           Traditional Heap Sort:
'
'                      0                                         0
'                     / \                                       / \
'                    /   \                                     /   \
'                   /     \                                   /     \
'                  /       \                                 /       \
'                 /         \                               /         \
'                /           \                             /           \
'               1             8                           1             2
'              / \           / \                         / \           / \
'             /   \         /   \                       /   \         /   \
'            2     5       9     12                    3     4       5     6
'           / \   / \     / \   / \                   / \   / \     / \   / \
'          3   4 6   7   10 11 13  14                7   8 9   10  11 12 13  14
'
'The reason for this is that the earlier data is to the left.  This arrangement can be
'maintained and stability ensured.
'
'
'Similar to traditional heap sort, the only guaranteed item is the top item.  We can be
'sure that it comes first.
'
'Unlike traditional heap sort, the last element is not swapped with the first then pushed down the heap.
'Stable heap sort removes the top item from the heap and restructures the remainder.
'
'Once elements are removed from the top, its children become orphans.
'Indexes to the orphans are stored in a Log(N) array in order.
'A virtual heap is constructed where the first (last numerically) orphan is the very top.
'Subsequent orphans then become children of their immediate predecessor.
'
'If elements 0 and 1 are removed from the heap structure above, the new virtual heap appears:
'
'                         ____
'                        /    8
'                       /    / \
'                __    /    /   \
'            2__/  5__/    9     12
'           / \   / \     / \   / \
'          3   4 6   7   10 11 13  14
'
'
'
'An Example
'----------
'
'Consider the following example:
'
'
'                       4a    (The keys are the numbers only - eg 4a = 4b)
'                      / \
'                     /   \
'                    /     \
'                   /       \
'                  /         \
'                 /           \
'                /             \
'               4b              3c
'              / \             / \
'             /   \           /   \
'            /     \         /     \
'           3a      3b      3d      2a
'          / \     / \     / \     / \
'         8   7a  4c  9   3e  6a  7b  6b
'
'In the above example, the bottom level of the heap is already arranged into heaps.
'When doing subsequent levels, swapping a value with its left child is trivial, for example
'swapping '4b' with the '3a' does not affect stability.
'However, swapping a value with its right child needs an extra step.  In the above example
'swapping '3c' with '2a' would place the '3d' before the '3c' and destroy stability.
'In this case it is necessary to rotate the equivalent values first: eg:
'
'                       4a    (The keys are the numbers only)
'                      / \
'                     /   \
'                    /     \
'                   /       \
'                  /         \
'                 /           \
'                /             \
'               3a              3e
'              / \             / \
'             /   \           /   \
'            /     \         /     \
'           4b      3b      3c      2
'          / \     / \     / \     / \
'         8   7a  4c  9   3d  6a  7b  6b
'
'Note that the '3a' has been wapped with '4b' without issue.
'Note that the '3c', '3d' and '3e' have been rotated.
'It is now possible to swap the '3e' with the '2'.
'Fortunately, all of the values needing rotating will be adjacent to the value needing rotating.
'
'The result:
'                       4a    (The keys are the numbers only)
'                      / \
'                     /   \
'                    /     \
'                   /       \
'                  /         \
'                 /           \
'                /             \
'               3a              2
'              / \             / \
'             /   \           /   \
'            /     \         /     \
'           4b      3b      3c      3e
'          / \     / \     / \     / \
'         8   7a  4c  9   3d  6a  7b  6b
'
'
'Another other difference occurs when a node can be swapped with either of its children.
'In this case it must first be swapped with its left child.  Then it may be necessary
'the reswap the result with the right child.  In the above example '4a' is first swapped with '3a'.
'
'                       3a    (The keys are the numbers only)
'                      / \
'                     /   \
'                    /     \
'                   /       \
'                  /         \
'                 /           \
'                /             \
'               4a              2
'              / \             / \
'             /   \           /   \
'            /     \         /     \
'           4b      3b      3c      3e
'          / \     / \     / \     / \
'         8   7a  4c  9   3d  6a  7b  6b
'
''4a' is rotated with '4b' and then '4b' is subsequently rotated with '3b' to give:
'
'                       3a    (The keys are the numbers only)
'                      / \
'                     /   \
'                    /     \
'                   /       \
'                  /         \
'                 /           \
'                /             \
'               3b              2
'              / \             / \
'             /   \           /   \
'            /     \         /     \
'           4a      4b      3c      3e
'          / \     / \     / \     / \
'         8   7a  4c  9   3d  6a  7b  6b
'
'Then the node in question is swapped with its right child if necessary.  Our example becomes:
'After '3a' is rotated with '3b'.
'
'                       2    (The keys are the numbers only)
'                      / \
'                     /   \
'                    /     \
'                   /       \
'                  /         \
'                 /           \
'                /             \
'               3a              3b
'              / \             / \
'             /   \           /   \
'            /     \         /     \
'           4a      4b      3c      3e
'          / \     / \     / \     / \
'         8   7a  4c  9   3d  6a  7b  6b
'
'
'The next step is to iteratively remove the top of the heap and restructure.
'
'
'                        2    (The keys are the numbers only)
'
'                 3a_______________
'                / \               \
'               /   \               3b
'              /     \             / \
'             /       \           /   \
'            /         \         /     \
'           4a          4b      3c      3e
'          / \         / \     / \     / \
'         8   7a      4c  9   3d  6a  7b  6b
'
'The heap topped by '3a' is restructured and the new parent (3a) is rebuilt.
'
'                        2    (The keys are the numbers only)
'                        3a
'
'                   4a_______
'                  / \       \
'                 8   7a      4b_________
'                            / \         \
'                           4c  9         3b
'                                        / \
'                                       /   \
'                                      /     \
'                                     3c      3e
'                                    / \     / \
'                                   3d  6a  7b  6b
'
'
'Then the heap is restructured
':2, 3a
'
'                   4a_______
'                  / \       \
'                 8   7a      4b_________
'                            / \         \
'                           4c  9         3b
'                                        / \
'                                       /   \
'                                      /     \
'                                     3c      3e
'                                    / \     / \
'                                   3d  6a  7b  6b
'
'
'Then the heaps topped by the new parents (4a, 4b) rebuilt
':2, 3a
'
'                   3b_______
'                  / \       \
'                 8   7a      3c_________
'                            / \         \
'                           4a  9         3d
'                                        / \
'                                       /   \
'                                      /     \
'                                     4b      3e
'                                    / \     / \
'                                   4c  6a  7b  6b
'
'
'Then the 3b is taken off the top, the heap is restructured
':2, 3a, 3b
'
'                 8
'                  \
'                   7a_______
'                            \
'                             3c_________
'                            / \         \
'                           4a  9         3d
'                                        / \
'                                       /   \
'                                      /     \
'                                     4b      3e
'                                    / \     / \
'                                   4c  6a  7b  6b
'
'
'The top of the new heap is rebuilt
':2, 3a, 3b
'
'                 3c
'                  \
'                   3d_______
'                            \
'                             3e_________
'                            / \         \
'                           8   9         4a
'                                        / \
'                                       /   \
'                                      /     \
'                                     4b      4c
'                                    / \     / \
'                                   7a  6a  7b  6b
'
'
'And so on.
'
'
'
'Navigating the Heap
'-------------------
'It is always necessary to know the size of the left half of the heap.
'This is a number
'No Items     Size of Left Side of Tree
'1            0
'2-3          1
'4-7          3
'8-15         7
'16-31        15
'
'The formula is     lngLeftSize = 2 ^ Int(Log(lngNoElements) / Log(2)) - 1
'
'Other important parameters are:
'- The total number of elements, this determines the bottom of any non-complete sub-heaps.
'- The Top of the heap (starts at 0 but increments as items are removed from the top of the heap).
'- The left/right integer.
'
'The left right integer contains bits that indicate whether the item is the left or right child
'of the parent.  The the bits are rolled on and off the integer as necessary and index 0 is
'always assumed to be the very top of the heap (even when it has been removed).
'
'Initially, before the heap is deformed the following rules apply:
'
'The left child of Item X = X + 1.  The lngLeftSize = (lngLeftSize - 1) / 2.  The left/right integer is rolled up one bit.
'The right child of Item X = X + 1 + lngLeftSize.The lngLeftSize = (lngLeftSize - 1) / 2.  The left/right integer is rolled up one bit and a 1 bit is added.
'If the lowest bit of the left/right integer = 0 then the parent of Item X is X-1.   The lngLeftSize = lngLeftSize * 2 + 1.  The left/right integer is rolled down one bit.
'If the lowest bit of the left/right integer = 1 then the parent of Item X is X-1-(lngLeftSize * 2 + 1).   The lngLeftSize = lngLeftSize * 2 + 1.  The left/right integer is rolled down one bit.
'The top of the tree is at position 0.
'The bottom of the tree is where lngLeftSize = 0 or the child index is past the end of the array.
'
'
'Once the heap begins to deform, these additional rules come into place.
'An array of size Log2(N) of heaps is maintained.  Initially this array contains only 1 entry for the 1
'heap where the top is zero.  Deforming the heap removes the 1 heap top and creates 2 heap tops.
'In our example above:
'
'There is initially 1 heap top the array appears as: [0][][][]
'(where zero is the index of the top of the heap.)
'The first deformation removes the 0 heap top and creates two: [8][1][][]
'(the item in element 1 is an additional child of element 8.)
'The second deformation removes 1 and creates 2: [8][5][2][]
'(the item in element 5 is an additional child of element 8 and the element 2 is an additional child of element 5.
'The next deformation goes to: [8][5][4][3]
'Element 3 begin a bottom level node is simply removed.
'The next deformation goes to: [8][5][][]
'Similarly Element 4 begin a bottom level node is simply removed.
'The next deformation goes to: [8][7][6][]
'Then [8][7][][]
'Then [8][][][]
'Then [12][9][][]
'Then [12][11][10][]
'Then [12][11][][]
'Then [12][][][]
'Then [14][13][][]
'Then [14][][][]
'Then [][][][]
'And the sort is complete


Sub StableHeapSort(ByRef theData() As DataElement, lngNoElements As Long)

    Dim lngLeftSize As Long
    
    Dim heapHeadsIndexes() As Long
    Dim heapHeadsLeftSizes() As Long
    Dim heapHeadsLeftRightBits() As Long
    Dim lngNoHeapHeads As Long
    
    Dim lngMaxNoHeaps As Long
    
    'Calculate the size of the left part
    '4-7 = 3
    '8-15 = 4
    lngLeftSize = Int(Log(lngNoElements) / Log(2))
    lngLeftSize = 2 ^ lngLeftSize - 1
    
    'Allocate space for the heap tops (orphan heaps)
    lngMaxNoHeaps = Int(Log(lngNoElements) / Log(2)) + 1
    ReDim heapHeadsIndexes(lngMaxNoHeaps) As Long
    ReDim heapHeadsLeftSizes(lngMaxNoHeaps) As Long
    ReDim heapHeadsLeftRightBits(lngMaxNoHeaps) As Long
        
    'Start with a single heap (only one orphan top)
    heapHeadsIndexes(0) = 0
    heapHeadsLeftSizes(0) = lngLeftSize
    heapHeadsLeftRightBits(0) = 0
    lngNoHeapHeads = 1
    
    'Build the single orphan heap
    BuildSubHeap theData, lngNoElements, 0, lngLeftSize, 0, 0, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
    
    'While there is at least one orphan heap
    While lngNoHeapHeads > 0
    
        'Because the virtual heap is a heap, the first (and top) element is the next ordered element.
        'We need to remove this top item from the virtual heap.  This will make orphans of its children.
        'We then need to reconstruct the virtual heap with the first child being the top of the new
        'virtual heap, the second child being an extra child of the first, and all other orphans becoming
        'children in the same manner.
    
        'If the first orphan heap has any children at all, then
        'removing this heap will create orphans of the children
        If heapHeadsLeftSizes(lngNoHeapHeads - 1) > 0 Then
        
            'The first orphan heap has two children
            If heapHeadsIndexes(lngNoHeapHeads - 1) + 1 + heapHeadsLeftSizes(lngNoHeapHeads - 1) < lngNoElements Then

                'Add both children as tops of heaps
                heapHeadsIndexes(lngNoHeapHeads) = heapHeadsIndexes(lngNoHeapHeads - 1) + 1
                heapHeadsLeftSizes(lngNoHeapHeads) = (heapHeadsLeftSizes(lngNoHeapHeads - 1) - 1) / 2
                heapHeadsLeftRightBits(lngNoHeapHeads) = heapHeadsLeftRightBits(lngNoHeapHeads - 1) * 2
                
                heapHeadsIndexes(lngNoHeapHeads - 1) = heapHeadsIndexes(lngNoHeapHeads - 1) + 1 + heapHeadsLeftSizes(lngNoHeapHeads - 1)
                heapHeadsLeftSizes(lngNoHeapHeads - 1) = (heapHeadsLeftSizes(lngNoHeapHeads - 1) - 1) / 2
                heapHeadsLeftRightBits(lngNoHeapHeads - 1) = heapHeadsLeftRightBits(lngNoHeapHeads - 1) * 2 + 1
            
                lngNoHeapHeads = lngNoHeapHeads + 1
                
                If lngNoHeapHeads > 2 Then
                    'Build the sub-heap where the second child is at the top
                    PushDownLeftFirst theData, lngNoElements, heapHeadsIndexes(lngNoHeapHeads - 2), heapHeadsLeftSizes(lngNoHeapHeads - 2), heapHeadsLeftRightBits(lngNoHeapHeads - 2), lngNoHeapHeads - 2, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
                End If
                'Build the sub-heap where the first child is at the top including the second child
                PushDownLeftFirst theData, lngNoElements, heapHeadsIndexes(lngNoHeapHeads - 1), heapHeadsLeftSizes(lngNoHeapHeads - 1), heapHeadsLeftRightBits(lngNoHeapHeads - 1), lngNoHeapHeads - 1, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
                
            Else
            
                'The top orphan element only has one child
                If heapHeadsIndexes(lngNoHeapHeads - 1) + 1 < lngNoElements Then
            
                    'Only add the left child as the top of a heap
                    heapHeadsIndexes(lngNoHeapHeads - 1) = heapHeadsIndexes(lngNoHeapHeads - 1) + 1
                    heapHeadsLeftSizes(lngNoHeapHeads - 1) = (heapHeadsLeftSizes(lngNoHeapHeads - 1) - 1) / 2
                    heapHeadsLeftRightBits(lngNoHeapHeads - 1) = heapHeadsLeftRightBits(lngNoHeapHeads - 1) * 2
                                
                    If lngNoHeapHeads > 2 Then
                        PushDownLeftFirst theData, lngNoElements, heapHeadsIndexes(lngNoHeapHeads - 1), heapHeadsLeftSizes(lngNoHeapHeads - 1), heapHeadsLeftRightBits(lngNoHeapHeads - 1), lngNoHeapHeads - 1, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
                    End If
                                
                Else
                
                    'The first heap has no children because we have run out of data
                    'Do not add any children as the tops of heaps
                    lngNoHeapHeads = lngNoHeapHeads - 1
                
                End If
                
            End If
            
        Else
            'The first heap has no children because it is at the very bottom possible level.
            'The first heap has no children so removing it will not create orphans
            lngNoHeapHeads = lngNoHeapHeads - 1
        End If
        
    Wend
    
End Sub

Function BuildSubHeap(theData() As DataElement, lngNoElements As Long, lngHeapTop As Long, lngLeftSize As Long, lngLeftRightBits As Long, lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long, heapHeadsLeftRightBits() As Long)

    'Build sub heap builds a heap where the element on top of the heap is to come before the two children.
    'It does this by first recursively building the children and then building the top level.
    
    If lngLeftSize >= 1 And lngHeapTop + 1 < lngNoElements Then
    
        'This function is only called when making the initial heap so dont worry about third children
    
        'has at least a left child
        BuildSubHeap theData, lngNoElements, lngHeapTop + 1, (lngLeftSize - 1) / 2, lngLeftRightBits * 2, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
        If lngHeapTop + lngLeftSize + 1 < lngNoElements Then
            'has a right child
            BuildSubHeap theData, lngNoElements, lngHeapTop + 1 + lngLeftSize, (lngLeftSize - 1) / 2, lngLeftRightBits * 2 + 1, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
        End If

        PushDownLeftFirst theData, lngNoElements, lngHeapTop, lngLeftSize, lngLeftRightBits, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
    End If


End Function

Function PushDownLeftFirst(theData() As DataElement, lngNoElements As Long, lngHeapTop As Long, lngLeftSize As Long, lngLeftRightBits As Long, lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long, heapHeadsLeftRightBits() As Long)

    'PushDownLeftFirst builds the top level of a sub-heap.
    '
    'Swapping the top element with its left child is ok.
    'Swapping the top element with its right child also means placing it afer all of the items in the left sub heap.
    'Before we swap an element with its right child, we need to be sure that there are no equally keyed elements in the left half.
    'If there is an equally keyed element in the left half, we need to swap the last element down the right child.

    Dim swapItem As DataElement
    Dim blNeedsRotate As Boolean

    'Swap the top element with the left if possible
    If lngLeftSize > 0 And lngHeapTop + 1 < lngNoElements Then
        'It has at least a left child
        
        'Swap on the left
        If theData(lngHeapTop).theKey > theData(lngHeapTop + 1).theKey Then
            swapItem.theKey = theData(lngHeapTop).theKey
            swapItem.originalOrder = theData(lngHeapTop).originalOrder
        
            theData(lngHeapTop).theKey = theData(lngHeapTop + 1).theKey
            theData(lngHeapTop).originalOrder = theData(lngHeapTop + 1).originalOrder
        
            theData(lngHeapTop + 1).theKey = swapItem.theKey
            theData(lngHeapTop + 1).originalOrder = swapItem.originalOrder
            
            PushDownLeftFirst theData, lngNoElements, lngHeapTop + 1, (lngLeftSize - 1) / 2, lngLeftRightBits * 2, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
        End If
    End If
        
    If lngLeftSize > 0 And lngHeapTop + lngLeftSize + 1 < lngNoElements Then
        'Swap on the right
        If theData(lngHeapTop).theKey > theData(lngHeapTop + 1 + lngLeftSize).theKey Then
        
            'If the item at the top of this sub heap leads a number
            'of equivalent values, then rotate them about so that the last
            'of the equivalent values is passed down the right
            If theData(lngHeapTop).theKey = theData(lngHeapTop + 1).theKey Then
                RotateEqualValues theData, lngNoElements, lngHeapTop, lngLeftSize, lngLeftRightBits, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
            End If
        
            swapItem.theKey = theData(lngHeapTop).theKey
            swapItem.originalOrder = theData(lngHeapTop).originalOrder
        
            theData(lngHeapTop).theKey = theData(lngHeapTop + 1 + lngLeftSize).theKey
            theData(lngHeapTop).originalOrder = theData(lngHeapTop + 1 + lngLeftSize).originalOrder
        
            theData(lngHeapTop + 1 + lngLeftSize).theKey = swapItem.theKey
            theData(lngHeapTop + 1 + lngLeftSize).originalOrder = swapItem.originalOrder
            
            PushDownLeftFirst theData, lngNoElements, lngHeapTop + 1 + lngLeftSize, (lngLeftSize - 1) / 2, lngLeftRightBits * 2 + 1, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
        End If

    End If
    
    If lngHeapNo <> 0 And lngHeapTop = heapHeadsIndexes(lngHeapNo) Then
        'This is the top of a heap, it will have a third child
        
        If theData(lngHeapTop).theKey > theData(heapHeadsIndexes(lngHeapNo - 1)).theKey Then
            'Needs to be swapped
        
            'See if we need to rotate equal values
            blNeedsRotate = False
            If lngLeftSize > 0 And lngHeapTop + 1 < lngNoElements Then
                If theData(lngHeapTop).theKey = theData(lngHeapTop + 1).theKey Then
                    blNeedsRotate = True
                End If
            End If
            If lngLeftSize > 0 And lngHeapTop + lngLeftSize + 1 < lngNoElements Then
                If theData(lngHeapTop).theKey = theData(lngHeapTop + 1 + lngLeftSize).theKey Then
                    blNeedsRotate = True
                End If
            End If
            
            If blNeedsRotate Then
                RotateEqualValues theData, lngNoElements, lngHeapTop, lngLeftSize, lngLeftRightBits, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
            End If
            
            'Swap it
            swapItem.theKey = theData(lngHeapTop).theKey
            swapItem.originalOrder = theData(lngHeapTop).originalOrder
        
            theData(lngHeapTop).theKey = theData(heapHeadsIndexes(lngHeapNo - 1)).theKey
            theData(lngHeapTop).originalOrder = theData(heapHeadsIndexes(lngHeapNo - 1)).originalOrder
        
            theData(heapHeadsIndexes(lngHeapNo - 1)).theKey = swapItem.theKey
            theData(heapHeadsIndexes(lngHeapNo - 1)).originalOrder = swapItem.originalOrder
            
            PushDownLeftFirst theData, lngNoElements, heapHeadsIndexes(lngHeapNo - 1), heapHeadsLeftSizes(lngHeapNo - 1), heapHeadsLeftRightBits(lngHeapNo - 1), lngHeapNo - 1, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits
        End If
            
    End If

End Function



Function RotateEqualValues(ByRef theData() As DataElement, ByVal lngNoElements As Long, ByVal lngHeapTop As Long, ByVal lngLeftSize As Long, lngLeftRightBits As Long, lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long, heapHeadsLeftRightBits() As Long)
    'This function is called when a node needs to be swapped with its right child
    'If the node leads a run of equally valued items on the left child then this swap
    'would wreck the stability of the sort.
    '
    'This function rotates the run of equally valued items by one so that it is the last
    'item that is at the top and is in turn swapped down the right child.
    
    Dim swapItem As DataElement
    Dim lngItemIndex As Long
    Dim lngItemLeftSize As Long
    Dim lngItemLeftRightBits As Long
    Dim lngItemHeapNo As Long
    
    Dim lngNextIndex As Long
    Dim lngNextLeftSize As Long
    Dim lngNextLeftRightBits As Long
    Dim lngNextHeapNo As Long
    
    'Find the last equally valued item
    lngItemIndex = FindLastEqualValue(theData, lngNoElements, lngHeapTop, lngLeftSize, lngLeftRightBits, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits, lngItemLeftSize, lngItemLeftRightBits, lngItemHeapNo)

    If lngItemIndex <> lngHeapTop Then
    
        'Remember the last item for later
        swapItem.theKey = theData(lngItemIndex).theKey
        swapItem.originalOrder = theData(lngItemIndex).originalOrder
    
        'Now rotate the values
        While lngItemIndex <> lngHeapTop
        
            lngNextIndex = FindNextEqualValue(theData, lngNoElements, lngItemIndex, lngItemLeftSize, lngItemLeftRightBits, lngItemHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits, lngNextLeftSize, lngNextLeftRightBits, lngNextHeapNo)
            
            theData(lngItemIndex).theKey = theData(lngNextIndex).theKey
            theData(lngItemIndex).originalOrder = theData(lngNextIndex).originalOrder
            
            lngItemIndex = lngNextIndex
            lngItemLeftSize = lngNextLeftSize
            lngItemLeftRightBits = lngNextLeftRightBits
            lngItemHeapNo = lngNextHeapNo
            
        Wend
        
        theData(lngHeapTop).theKey = swapItem.theKey
        theData(lngHeapTop).originalOrder = swapItem.originalOrder
        
    End If

End Function
Function FindLastEqualValue(ByRef theData() As DataElement, ByVal lngNoElements As Long, ByVal lngHeapTop As Long, ByVal lngLeftSize As Long, ByVal lngLeftRightBits As Long, ByVal lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long, heapHeadsLeftRightBits() As Long, ByRef lngLastLeftSize As Long, ByRef lngLastLeftRightBits As Long, ByRef lngLastHeapNo As Long) As Long
    'Function that returns the index into the data of the last element
    'in a subheap that has the same key as the top item in the sub-heap.
    'Also returns lngPath, the bits indicate the left/right path to the last item

    Dim blAtEnd As Boolean
    
    Dim lngLastIndex As Long
    
    blAtEnd = False
        
    lngLastIndex = lngHeapTop
    lngLastLeftSize = lngLeftSize
    lngLastLeftRightBits = lngLeftRightBits
    lngLastHeapNo = lngHeapNo
        
    While Not blAtEnd
        
        blAtEnd = True
        
        'Check for third child
        If lngLastHeapNo > 0 And lngLastIndex = heapHeadsIndexes(lngLastHeapNo) Then
            'Has a third child
            If theData(heapHeadsIndexes(lngLastHeapNo - 1)).theKey = theData(lngHeapTop).theKey Then
                blAtEnd = False
                lngLastIndex = heapHeadsIndexes(lngLastHeapNo - 1)
                lngLastLeftSize = heapHeadsLeftSizes(lngLastHeapNo - 1)
                lngLastLeftRightBits = heapHeadsLeftRightBits(lngLastHeapNo - 1)
                lngLastHeapNo = lngLastHeapNo - 1
                blAtEnd = False
            End If
        End If
        
        If lngLastLeftSize > 0 Then
            'Check if the right child has the same key
            If blAtEnd And lngLastIndex + 1 + lngLastLeftSize < lngNoElements Then
                If theData(lngLastIndex + 1 + lngLastLeftSize).theKey = theData(lngLastIndex).theKey Then
                    lngLastIndex = lngLastIndex + 1 + lngLastLeftSize
                    lngLastLeftSize = (lngLastLeftSize - 1) / 2
                    lngLastLeftRightBits = lngLastLeftRightBits * 2 + 1
                    blAtEnd = False
                End If
            End If
        
            'If the right child is different then check the left child
            If blAtEnd And lngLastIndex + 1 < lngNoElements Then
                If theData(lngLastIndex + 1).theKey = theData(lngLastIndex).theKey Then
                    lngLastIndex = lngLastIndex + 1
                    lngLastLeftSize = (lngLastLeftSize - 1) / 2
                    lngLastLeftRightBits = lngLastLeftRightBits * 2
                    blAtEnd = False
                End If
            End If
        End If
        
    Wend


    FindLastEqualValue = lngLastIndex

End Function

Function FindNextEqualValue(ByRef theData() As DataElement, ByVal lngNoElements As Long, ByVal lngHeapTop As Long, ByVal lngLeftSize As Long, lngLeftRightBits As Long, lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long, heapHeadsLeftRightBits() As Long, ByRef lngNextLeftSize As Long, ByRef lngNextLeftRightBits As Long, ByRef lngNextHeapNo As Long) As Long

    Dim lngNextIndex As Long
    Dim blGotNext As Boolean
    Dim blCheckRightChild As Boolean
    
    blGotNext = False
    blCheckRightChild = False
    
    If lngHeapTop = heapHeadsIndexes(lngHeapNo) Then
        If lngHeapNo = 0 Then
            lngNextIndex = -1 'Should never happen
            blGotNext = True
        Else
            'This is the top of one of the heaps
            'Start with the parent
            lngNextIndex = heapHeadsIndexes(lngHeapNo - 1)
            lngNextLeftSize = heapHeadsLeftSizes(lngHeapNo - 1)
            lngNextLeftRightBits = heapHeadsLeftRightBits(lngHeapNo - 1)
            lngNextHeapNo = lngHeapNo - 1
            
            'Check the right child of this heap
            If lngNextIndex + 1 + lngNextLeftSize < lngNoElements Then
                If theData(lngNextIndex + 1 + lngNextLeftSize).theKey = theData(lngHeapTop).theKey Then
                    'This is the one
                    lngNextIndex = lngNextIndex + 1 + lngNextLeftSize
                    lngNextLeftSize = (lngNextLeftSize - 1) / 2
                    lngNextLeftRightBits = lngNextLeftRightBits * 2 + 1
                    blGotNext = True
                    blCheckRightChild = True
                End If
            End If
            
            'Check the left child of this heap
            If (Not blGotNext) And lngNextIndex + 1 < lngNoElements Then
                If theData(lngNextIndex + 1).theKey = theData(lngHeapTop).theKey Then
                    'This is the one
                    lngNextIndex = lngNextIndex + 1
                    lngNextLeftSize = (lngNextLeftSize - 1) / 2
                    lngNextLeftRightBits = lngNextLeftRightBits * 2
                    blGotNext = True
                    blCheckRightChild = True
                End If
            End If
        End If
    Else
        'This is not the top of one of the heaps
        
        If (lngLeftRightBits And 1) = 0 Then
            'Start with the parent
            lngNextIndex = lngHeapTop - 1
            lngNextLeftRightBits = lngLeftRightBits \ 2
            lngNextLeftSize = lngLeftSize * 2 + 1
        Else
            'Start with the parent
            lngNextIndex = lngHeapTop - 1 - (lngLeftSize * 2 + 1)
            lngNextLeftRightBits = lngLeftRightBits \ 2
            lngNextLeftSize = lngLeftSize * 2 + 1
            
            'This was a right child and so check the left child
            If theData(lngNextIndex + 1).theKey = theData(lngNextIndex).theKey Then
                lngNextIndex = lngNextIndex + 1
                lngNextLeftRightBits = lngNextLeftRightBits * 2
                lngNextLeftSize = (lngNextLeftSize - 1) / 2
                blGotNext = True
                blCheckRightChild = True
            End If
        End If
    End If
    
    If blCheckRightChild Then
        lngNextIndex = FindLastEqualValue(theData, lngNoElements, lngNextIndex, lngNextLeftSize, lngNextLeftRightBits, lngNextHeapNo, heapHeadsIndexes, heapHeadsLeftSizes, heapHeadsLeftRightBits, lngNextLeftSize, lngNextLeftRightBits, lngNextHeapNo)
    End If

    FindNextEqualValue = lngNextIndex

End Function



