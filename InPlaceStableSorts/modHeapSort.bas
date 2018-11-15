Attribute VB_Name = "modHeapSort"
Option Explicit

Function HeapSort(ByRef theData() As DataElement, ByVal lngNoElements As Long)

    'Lean and mean Heap sort.
    
    Dim lngSubHeapTop As Long
    
    Dim lngThisParent As Long
    Dim lngSwapWith As Long
    
    Dim lngSwapValue As Long
    
    Dim lngNoElementsInHeap As Long
    
    'Build the intial heap
    'Build it from the bottom up, but ignore the very bottom level without children
    lngSubHeapTop = lngNoElements \ 2 - 1
    While lngSubHeapTop >= 0
    
        'Push this value all the way down the heap
        lngThisParent = lngSubHeapTop
        While lngThisParent >= 0
            'Swap it with the greatest child
            lngSwapWith = PickSwapChild(theData, lngNoElements, lngThisParent)
            If lngSwapWith >= 0 Then
                lngSwapValue = theData(lngSwapWith).theKey
                theData(lngSwapWith).theKey = theData(lngThisParent).theKey
                theData(lngThisParent).theKey = lngSwapValue
            
                lngSwapValue = theData(lngSwapWith).originalOrder
                theData(lngSwapWith).originalOrder = theData(lngThisParent).originalOrder
                theData(lngThisParent).originalOrder = lngSwapValue
            End If
            
            lngThisParent = lngSwapWith
        Wend
    
        lngSubHeapTop = lngSubHeapTop - 1
    Wend
    
    'Progressively make the heap smaller
    lngNoElementsInHeap = lngNoElements
    While lngNoElementsInHeap > 1
        
        'We know that the last item is at the top of the heap
        'Put it after the end of the heap and shrink the heap.
        'Then rebuild the heap so that the next last item is at the top.
        'And so on
        
        'Swap the item on the top of the heap with the last element in the heap.
        lngSwapValue = theData(0).theKey
        theData(0).theKey = theData(lngNoElementsInHeap - 1).theKey
        theData(lngNoElementsInHeap - 1).theKey = lngSwapValue
        
        lngSwapValue = theData(0).originalOrder
        theData(0).originalOrder = theData(lngNoElementsInHeap - 1).originalOrder
        theData(lngNoElementsInHeap - 1).originalOrder = lngSwapValue
        
        'Shrink the heap
        lngNoElementsInHeap = lngNoElementsInHeap - 1
        
        'Rebuild the heap
        lngThisParent = 0
        While lngThisParent >= 0
            lngSwapWith = PickSwapChild(theData, lngNoElementsInHeap, lngThisParent)
            If lngSwapWith >= 0 Then
                lngSwapValue = theData(lngSwapWith).theKey
                theData(lngSwapWith).theKey = theData(lngThisParent).theKey
                theData(lngThisParent).theKey = lngSwapValue
            
                lngSwapValue = theData(lngSwapWith).originalOrder
                theData(lngSwapWith).originalOrder = theData(lngThisParent).originalOrder
                theData(lngThisParent).originalOrder = lngSwapValue
            End If
            lngThisParent = lngSwapWith
        Wend
    Wend

End Function

Function PickSwapChild(ByRef theData() As DataElement, ByVal lngNoElements As Long, ByVal lngSubHeapTop As Long)

    Dim lngLeftChild As Long
    Dim lngRightChild As Long
    Dim lngSwappedWith As Long

    'Where are the left and right children
    lngLeftChild = lngSubHeapTop * 2 + 1
    
    If lngLeftChild >= lngNoElements Then
        'No Children, no swapping
        lngSwappedWith = -1
    Else
        'Where is the right child
        lngRightChild = lngSubHeapTop * 2 + 2
        
        'Arrange this sub heap
        If theData(lngLeftChild).theKey > theData(lngSubHeapTop).theKey Then
            'Maybe we should swap with the left child
            If lngRightChild < lngNoElements Then
                If theData(lngRightChild).theKey > theData(lngLeftChild).theKey Then
                    'The right child is even bigger so swap with it
                    lngSwappedWith = lngRightChild
                Else
                    'Yes swap with the left child
                    lngSwappedWith = lngLeftChild
                End If
            Else
                'No right child so definitely swap with the left
                lngSwappedWith = lngLeftChild
            End If
        Else
            If lngRightChild < lngNoElements Then
                If theData(lngRightChild).theKey > theData(lngSubHeapTop).theKey Then
                    'We should swap with the right child
                    lngSwappedWith = lngRightChild
                Else
                    'No swapping
                    lngSwappedWith = -1
                End If
            Else
                'No swapping
                lngSwappedWith = -1
            End If
        End If
    End If

    PickSwapChild = lngSwappedWith

End Function
