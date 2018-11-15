Attribute VB_Name = "modCraigsSmoothSort"
Option Explicit

'Craigs Smooth Sort
'------------------
'
'Summary
'-------
'
'Stable Heap Sort works in a similar manner to traditional heap sort.
'There are some differences:
'- The heaps are arranged so that the parent of a heap is the least value of the key.
'  Where there are equal key values, then the parent is also the earlier value in natural order.
'  The mapping of the heaps is different to traditional heap sort and supports this by default.
'- When sorted items are removed from the heap, they are taken from the top.  The subsequent orphan
'  elements then become part of a virtual heap.
'
'Similar to Edsger Dijkstra's Smoothsort, the heap is arranged in natural order.  Except using a
'binary arrangement.
'Also similar to Smoothsort, the heap is destroyed from the top.  Except using a buffer of orphan nodes
'instead of mathematics.
'
'This sort is a version of my stable heap sort but with the stability removed in the sake of performance.
'
'Performance
'-----------
'
'It should be O(NLogN)
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


Sub CraigsSmoothSort(ByRef theData() As DataElement, lngNoElements As Long)

    Dim lngLeftSize As Long
    
    Dim heapHeadsIndexes() As Long
    Dim heapHeadsLeftSizes() As Long
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
        
    'Start with a single heap (only one orphan top)
    heapHeadsIndexes(0) = 0
    heapHeadsLeftSizes(0) = lngLeftSize
    lngNoHeapHeads = 1
    
    'Build the single orphan heap
    BuildCraigsSubHeap theData, lngNoElements, 0, lngLeftSize, 0, heapHeadsIndexes, heapHeadsLeftSizes
    
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
                
                heapHeadsIndexes(lngNoHeapHeads - 1) = heapHeadsIndexes(lngNoHeapHeads - 1) + 1 + heapHeadsLeftSizes(lngNoHeapHeads - 1)
                heapHeadsLeftSizes(lngNoHeapHeads - 1) = (heapHeadsLeftSizes(lngNoHeapHeads - 1) - 1) / 2
            
                lngNoHeapHeads = lngNoHeapHeads + 1
                
                If lngNoHeapHeads > 2 Then
                    'Build the sub-heap where the second child is at the top
                    PushDown theData, lngNoElements, heapHeadsIndexes(lngNoHeapHeads - 2), heapHeadsLeftSizes(lngNoHeapHeads - 2), lngNoHeapHeads - 2, heapHeadsIndexes, heapHeadsLeftSizes
                End If
                'Build the sub-heap where the first child is at the top including the second child
                PushDown theData, lngNoElements, heapHeadsIndexes(lngNoHeapHeads - 1), heapHeadsLeftSizes(lngNoHeapHeads - 1), lngNoHeapHeads - 1, heapHeadsIndexes, heapHeadsLeftSizes
                
            Else
            
                'The top orphan element only has one child
                If heapHeadsIndexes(lngNoHeapHeads - 1) + 1 < lngNoElements Then
            
                    'Only add the left child as the top of a heap
                    heapHeadsIndexes(lngNoHeapHeads - 1) = heapHeadsIndexes(lngNoHeapHeads - 1) + 1
                    heapHeadsLeftSizes(lngNoHeapHeads - 1) = (heapHeadsLeftSizes(lngNoHeapHeads - 1) - 1) / 2
                                
                    If lngNoHeapHeads > 2 Then
                        PushDown theData, lngNoElements, heapHeadsIndexes(lngNoHeapHeads - 1), heapHeadsLeftSizes(lngNoHeapHeads - 1), lngNoHeapHeads - 1, heapHeadsIndexes, heapHeadsLeftSizes
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

Function BuildCraigsSubHeap(theData() As DataElement, lngNoElements As Long, lngHeapTop As Long, lngLeftSize As Long, lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long)

    'Build sub heap builds a heap where the element on top of the heap is to come before the two children.
    'It does this by first recursively building the children and then building the top level.
    
    If lngLeftSize >= 1 And lngHeapTop + 1 < lngNoElements Then
    
        'This function is only called when making the initial heap so dont worry about third children
    
        'has at least a left child
        BuildCraigsSubHeap theData, lngNoElements, lngHeapTop + 1, (lngLeftSize - 1) / 2, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes
        If lngHeapTop + lngLeftSize + 1 < lngNoElements Then
            'has a right child
            BuildCraigsSubHeap theData, lngNoElements, lngHeapTop + 1 + lngLeftSize, (lngLeftSize - 1) / 2, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes
        End If

        PushDown theData, lngNoElements, lngHeapTop, lngLeftSize, lngHeapNo, heapHeadsIndexes, heapHeadsLeftSizes
    End If


End Function

Function PushDown(theData() As DataElement, lngNoElements As Long, lngHeapTop As Long, lngLeftSize As Long, lngHeapNo As Long, heapHeadsIndexes() As Long, heapHeadsLeftSizes() As Long)

    'PushDown builds the top level of a sub-heap.
    '

    Dim swapItem As Long
    Dim lngSwapWith As Long
    Dim lngSwapWithLeftSize As Long
    Dim lngSwapWithHeapNo As Long
    
    'Assume no swapping
    lngSwapWith = lngHeapTop
    lngSwapWithLeftSize = lngLeftSize
    lngSwapWithHeapNo = lngHeapNo
    

    'Swap the top element with the left if possible
    If lngLeftSize > 0 And lngHeapTop + 1 < lngNoElements Then
        'It has at least a left child
        
        'Swap on the left
        If theData(lngSwapWith).theKey > theData(lngHeapTop + 1).theKey Then
            'Yes, it is possible to swap with the left child
            lngSwapWith = lngHeapTop + 1
            lngSwapWithLeftSize = (lngLeftSize - 1) / 2
            lngSwapWithHeapNo = lngHeapNo
        End If
    End If
        
    If lngLeftSize > 0 And lngHeapTop + lngLeftSize + 1 < lngNoElements Then
        'It has a right child
        If theData(lngSwapWith).theKey > theData(lngHeapTop + 1 + lngLeftSize).theKey Then
            'Yes it is possible to swap with the right child
            lngSwapWith = lngHeapTop + 1 + lngLeftSize
            lngSwapWithLeftSize = (lngLeftSize - 1) / 2
            lngSwapWithHeapNo = lngHeapNo
        End If
    End If
    
    If lngHeapNo <> 0 And lngHeapTop = heapHeadsIndexes(lngHeapNo) Then
        'This is the top of a heap, it will have a third child
        If theData(lngSwapWith).theKey > theData(heapHeadsIndexes(lngHeapNo - 1)).theKey Then
            'Needs to be swapped
            lngSwapWith = heapHeadsIndexes(lngHeapNo - 1)
            lngSwapWithLeftSize = heapHeadsLeftSizes(lngHeapNo - 1)
            lngSwapWithHeapNo = lngHeapNo - 1
        End If
    End If
    
    If lngSwapWith <> lngHeapTop Then
        'Swap it
        swapItem = theData(lngHeapTop).theKey
        theData(lngHeapTop).theKey = theData(lngSwapWith).theKey
        theData(lngSwapWith).theKey = swapItem
        
        swapItem = theData(lngHeapTop).originalOrder
        theData(lngHeapTop).originalOrder = theData(lngSwapWith).originalOrder
        theData(lngSwapWith).originalOrder = swapItem
        
        PushDown theData, lngNoElements, lngSwapWith, lngSwapWithLeftSize, lngSwapWithHeapNo, heapHeadsIndexes, heapHeadsLeftSizes
    End If

End Function



