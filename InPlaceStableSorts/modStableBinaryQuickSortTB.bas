Attribute VB_Name = "modStableBinaryQuickSortTB"
Option Explicit

'StableBinaryQuickSortTB
'-----------------------
'
'Written by Craig Brown 21/1/13
'
'Using the algorithm used in Thomas Baudel's stable quick-sort.
'
'It differs from traditional quicksort in that instead of swapping elements in the pivot function,
'the array segments is recursively divided into two and each half is pivotted in a stable manner.
'When the two halves are pivotted, they can be stably merged together to complete the whole pivot
'function.


Const QUICKSORTDEPTH2 = 32

Global SMALLSEGMENTSIZETB As Long
Global smallBufferTB() As DataElement


Function StableBinaryQuickSortTB(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long)

    'These two arrays can be dynamically allocated to be log2(number of elements).
    Dim lngStackStarts(QUICKSORTDEPTH2) As Long
    Dim lngStackEnds(QUICKSORTDEPTH2) As Long
    Dim lngStackSize As Long
    
    Dim lngPivotValue(3) As Long
    Dim lngPivotPoint(3) As Long
    Dim lngI As Long
    Dim lngJ As Long
    
    Dim lngNewStart1 As Long
    Dim lngNewEnd1 As Long
    Dim lngNewStart2 As Long
    Dim lngNewEnd2 As Long
    
    Dim blFoundPivot As Boolean
    
    Dim lngTemp As Long
    
    lngStackSize = 0
    lngStackStarts(lngStackSize) = lngFirstElement
    lngStackEnds(lngStackSize) = lngLastElement
    lngStackSize = lngStackSize + 1
    
    Rnd -1
    Randomize 1
    
    While lngStackSize > 0
    
        lngStackSize = lngStackSize - 1
        
        If lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) < 6 Then
            InsertSort lngStackStarts(lngStackSize), lngStackEnds(lngStackSize) + 1, theData
        Else
    
            lngPivotPoint(0) = Int(Rnd() * (lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) + 1)) + lngStackStarts(lngStackSize)
            lngPivotPoint(1) = Int(Rnd() * (lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) + 1)) + lngStackStarts(lngStackSize)
            lngPivotPoint(2) = Int(Rnd() * (lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) + 1)) + lngStackStarts(lngStackSize)
            lngPivotValue(0) = theData(lngPivotPoint(0)).theKey
            lngPivotValue(1) = theData(lngPivotPoint(1)).theKey
            lngPivotValue(2) = theData(lngPivotPoint(2)).theKey
            OrderLongs lngPivotValue(0), lngPivotValue(1)
            OrderLongs lngPivotValue(1), lngPivotValue(2)
            OrderLongs lngPivotValue(0), lngPivotValue(1)
            
            blFoundPivot = False
            If lngPivotValue(1) = lngPivotValue(2) Then
                'Top two pivot values are the same, maybe use the smaller one
                If lngPivotValue(0) = lngPivotValue(1) Then
                    'All three pivot values are the same
                    'Find another value
                    lngI = lngStackStarts(lngStackSize)
                    While lngI <= lngStackEnds(lngStackSize) And Not blFoundPivot
                        If theData(lngI).theKey <> lngPivotValue(1) Then
                            blFoundPivot = True
                            If theData(lngI).theKey < lngPivotValue(1) Then
                                lngPivotValue(1) = theData(lngI).theKey
                            Else
                                lngPivotValue(2) = theData(lngI).theKey
                            End If
                        Else
                            lngI = lngI + 1
                        End If
                    Wend
                    
                Else
                    'Top two pivot values are the same, use the smaller one
                    lngPivotValue(1) = lngPivotValue(0)
                    blFoundPivot = True
                End If
            Else
                'Use the middle pivot
                lngPivotValue(1) = lngPivotValue(1)
                blFoundPivot = True
            End If
            
            'If a pivot cannot be found then all of the data must have the same key
            If blFoundPivot Then
            
                'Swap the data around
                lngPivotPoint(1) = StableBinaryPivotTB(theData, lngStackStarts(lngStackSize), lngStackEnds(lngStackSize), lngPivotValue(1))
                
                'Put the two parts on the stack
                
                'Put the larget part on first so that the stack doesn't overflow
                lngNewStart1 = lngStackStarts(lngStackSize)
                lngNewEnd1 = lngStackStarts(lngStackSize) + lngPivotPoint(1) - 1
                lngNewStart2 = lngStackStarts(lngStackSize) + lngPivotPoint(1)
                lngNewEnd2 = lngStackEnds(lngStackSize)
                
                If lngNewEnd1 - lngNewStart1 > lngNewEnd2 - lngNewStart2 Then
                    'Part 1 is bigger put it on first
                    lngStackStarts(lngStackSize) = lngNewStart1
                    lngStackEnds(lngStackSize) = lngNewEnd1
                    lngStackStarts(lngStackSize + 1) = lngNewStart2
                    lngStackEnds(lngStackSize + 1) = lngNewEnd2
                Else
                    'Part 2 is bigger put it on first
                    lngStackStarts(lngStackSize) = lngNewStart2
                    lngStackEnds(lngStackSize) = lngNewEnd2
                    lngStackStarts(lngStackSize + 1) = lngNewStart1
                    lngStackEnds(lngStackSize + 1) = lngNewEnd1
                End If
                lngStackSize = lngStackSize + 2
                
            End If
            
        End If
    Wend

End Function

Function StableBinaryPivotTB(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

    Dim lngNoUnders As Long
    Dim lngNoUnders1 As Long
    Dim lngNoUnders2 As Long
    
    Dim lngWholeSegmentSize As Long
    Dim lngSegment1Size As Long
    Dim lngSegment2Size As Long
    
    Dim lngTop As Long
    Dim lngBot As Long
    
    Dim lngSwapValue As Long
    
    lngWholeSegmentSize = (lngLastElement - lngFirstElement + 1)
    
    If lngWholeSegmentSize <= SMALLSEGMENTSIZETB Then
        'Only a few elements do it faster with a small buffer
        lngNoUnders = StableBinaryPivotSmallTB(theData(), lngFirstElement, lngLastElement, lngPivotValue)
    Else
        
        If lngWholeSegmentSize <= 2 Then
            'Only a few elements do it faster with a small buffer
            If lngWholeSegmentSize <= 1 Then
                If theData(lngFirstElement).theKey <= lngPivotValue Then
                    lngNoUnders = 1
                Else
                    lngNoUnders = 0
                End If
            Else
                If theData(lngLastElement).theKey <= lngPivotValue And theData(lngFirstElement).theKey > lngPivotValue Then
                    'Swap them around
                    lngSwapValue = theData(lngLastElement).theKey
                    theData(lngLastElement).theKey = theData(lngFirstElement).theKey
                    theData(lngFirstElement).theKey = lngSwapValue
                    lngSwapValue = theData(lngLastElement).originalOrder
                    theData(lngLastElement).originalOrder = theData(lngFirstElement).originalOrder
                    theData(lngFirstElement).originalOrder = lngSwapValue
                    
                    lngNoUnders = 1
                Else
                    If theData(lngFirstElement).theKey <= lngPivotValue Then
                        If theData(lngLastElement).theKey <= lngPivotValue Then
                            lngNoUnders = 2
                        Else
                            lngNoUnders = 1
                        End If
                    Else
                        lngNoUnders = 0
                    End If
                End If
            End If
        Else
        
            'Pivot two halves of this data segment
            lngSegment1Size = lngWholeSegmentSize / 2
            lngSegment2Size = lngWholeSegmentSize - lngSegment1Size
            
            lngNoUnders1 = StableBinaryPivotTB(theData(), lngFirstElement, lngFirstElement + lngSegment1Size - 1, lngPivotValue)
            lngNoUnders2 = StableBinaryPivotTB(theData(), lngFirstElement + lngSegment1Size, lngFirstElement + lngSegment1Size + lngSegment2Size - 1, lngPivotValue)
            
            'Now join the two halves together
            
            'Reverse the order of the overs in the first segment
            lngTop = lngFirstElement + lngSegment1Size - 1
            lngBot = lngFirstElement + lngNoUnders1
            While lngTop > lngBot
                lngSwapValue = theData(lngTop).theKey
                theData(lngTop).theKey = theData(lngBot).theKey
                theData(lngBot).theKey = lngSwapValue
            
                lngSwapValue = theData(lngTop).originalOrder
                theData(lngTop).originalOrder = theData(lngBot).originalOrder
                theData(lngBot).originalOrder = lngSwapValue
                
                lngTop = lngTop - 1
                lngBot = lngBot + 1
            Wend
            
            'Reverse the order of the unders in the second segment
            lngTop = lngFirstElement + lngSegment1Size + lngNoUnders2 - 1
            lngBot = lngFirstElement + lngSegment1Size
            While lngTop > lngBot
                lngSwapValue = theData(lngTop).theKey
                theData(lngTop).theKey = theData(lngBot).theKey
                theData(lngBot).theKey = lngSwapValue
            
                lngSwapValue = theData(lngTop).originalOrder
                theData(lngTop).originalOrder = theData(lngBot).originalOrder
                theData(lngBot).originalOrder = lngSwapValue
                
                lngTop = lngTop - 1
                lngBot = lngBot + 1
            Wend
            
            'Revers the order of both of the above
            lngTop = lngFirstElement + lngSegment1Size + lngNoUnders2 - 1
            lngBot = lngFirstElement + lngNoUnders1
            While lngTop > lngBot
                lngSwapValue = theData(lngTop).theKey
                theData(lngTop).theKey = theData(lngBot).theKey
                theData(lngBot).theKey = lngSwapValue
            
                lngSwapValue = theData(lngTop).originalOrder
                theData(lngTop).originalOrder = theData(lngBot).originalOrder
                theData(lngBot).originalOrder = lngSwapValue
                
                lngTop = lngTop - 1
                lngBot = lngBot + 1
            Wend
            
            
            lngNoUnders = lngNoUnders1 + lngNoUnders2
        End If
        
    End If
    StableBinaryPivotTB = lngNoUnders

End Function


Function StableBinaryPivotSmallTB(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

    'A function to pivot a small segment of data using a buffer.

    Dim lngI As Long
    Dim lngNoUnders As Long
    Dim lngNoOvers As Long
    
    lngNoUnders = 0
    lngNoOvers = 0
    
    'Loop through the data and compact the under values to the front of the original array
    'and store the over values in the buffer
    For lngI = lngFirstElement To lngLastElement
        If theData(lngI).theKey <= lngPivotValue Then
            theData(lngFirstElement + lngNoUnders).theKey = theData(lngI).theKey
            theData(lngFirstElement + lngNoUnders).originalOrder = theData(lngI).originalOrder
            lngNoUnders = lngNoUnders + 1
        Else
            smallBufferTB(lngNoOvers).theKey = theData(lngI).theKey
            smallBufferTB(lngNoOvers).originalOrder = theData(lngI).originalOrder
            lngNoOvers = lngNoOvers + 1
        End If
    Next
    
    'Now put the over values back into the original array after the unders
    lngI = 0
    While lngI < lngNoOvers
        theData(lngFirstElement + lngNoUnders + lngI).theKey = smallBufferTB(lngI).theKey
        theData(lngFirstElement + lngNoUnders + lngI).originalOrder = smallBufferTB(lngI).originalOrder
        lngI = lngI + 1
    Wend

    StableBinaryPivotSmallTB = lngNoUnders

End Function

