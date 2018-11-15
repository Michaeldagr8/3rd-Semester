Attribute VB_Name = "modStableBinaryQuickSortCB"
Option Explicit



Const QUICKSORTDEPTH2 = 32

Global SMALLSEGMENTSIZECB As Long
Global smallBufferCB() As DataElement

'StableBinaryQuickSortCB
'-----------------------
'
'Written by Craig Brown - 21/1/13.
'
'This is a stable version of quicksort.  It differs from traditional quicksort in that
'instead of swapping elements that need to be pivotted, runs of elements are rotated around
'to maintain stability.
'
'The differences are found in the pivot function used.
'


Function StableBinaryQuickSortCB(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long)

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
                lngPivotPoint(1) = StableBinaryPivotCB(theData, lngStackStarts(lngStackSize), lngStackEnds(lngStackSize), lngPivotValue(1))
                
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

Function StableBinaryPivotCB(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

    'The complicated selection of the pivot value
    'has already guaranteed that there will be some values above the pivot
    'and some at or below the pivot


    'This function pivots the block of data in a binary way from the bottom up
    '
    'Eg, if the data starts:
    '
    '     385828702948983 with a pivot value of 5 it will merge the runs together like:
    '      >>><    ><
    '     328788702498983 by swapping every second run, and then do it again
    '       >>>>><<<
    '     320248788798983 then again
    '          >>>>>>>>>
    '     320243878879898 to get the finished pivot
    '

    Dim lngFirstOverValue As Long
    Dim lngLastUnderValue As Long
    
    Dim lngStartIndex As Long
    Dim lngOverEnd1 As Long
    Dim lngUnderEnd1 As Long
    Dim lngNoElementsToProcess As Long
    
    Dim lngOverEnd2 As Long
    Dim lngUnderEnd2 As Long
    
    Dim i As Long
    Dim j As Long
    
    Dim lngSwapValue As Long
    
    Dim blKeepLooping As Long
    
    If (lngLastElement - lngFirstElement + 1) <= SMALLSEGMENTSIZECB Then
        'For small segments use a small buffer for performance
        lngFirstOverValue = StableBinaryPivotSmallCBForward(theData, lngFirstElement, lngLastElement, lngPivotValue) + lngFirstElement
    Else
    
        'Skip things that can be skipped
        lngFirstOverValue = lngFirstElement
        While theData(lngFirstOverValue).theKey <= lngPivotValue
            lngFirstOverValue = lngFirstOverValue + 1
        Wend
        lngLastUnderValue = lngLastElement
        While theData(lngLastUnderValue).theKey > lngPivotValue
            lngLastUnderValue = lngLastUnderValue - 1
        Wend
        
        'Run through the data using the small buffer as much as we can
        If lngFirstOverValue < lngLastUnderValue And SMALLSEGMENTSIZECB > 4 Then
            lngStartIndex = lngFirstOverValue
            lngNoElementsToProcess = lngLastUnderValue - lngStartIndex + 1
            If lngNoElementsToProcess > SMALLSEGMENTSIZECB Then
                lngNoElementsToProcess = SMALLSEGMENTSIZECB
            End If
            lngFirstOverValue = StableBinaryPivotSmallCBForward(theData, lngStartIndex, lngStartIndex + lngNoElementsToProcess - 1, lngPivotValue) + lngFirstOverValue
            lngStartIndex = lngStartIndex + lngNoElementsToProcess
            While lngStartIndex < lngLastUnderValue
                lngNoElementsToProcess = lngLastUnderValue - lngStartIndex + 1
                If lngNoElementsToProcess > SMALLSEGMENTSIZECB Then
                    lngNoElementsToProcess = SMALLSEGMENTSIZECB
                End If
                StableBinaryPivotSmallCBForward theData, lngStartIndex, lngStartIndex + lngNoElementsToProcess - 1, lngPivotValue
                lngStartIndex = lngStartIndex + lngNoElementsToProcess
                
                'Do it again, but do it backwards, this will group runs of overs and unders together
                If lngStartIndex < lngLastUnderValue Then
                    lngNoElementsToProcess = lngLastUnderValue - lngStartIndex + 1
                    If lngNoElementsToProcess > SMALLSEGMENTSIZECB Then
                        lngNoElementsToProcess = SMALLSEGMENTSIZECB
                    End If
                    StableBinaryPivotSmallCBBackward theData, lngStartIndex, lngStartIndex + lngNoElementsToProcess - 1, lngPivotValue
                    lngStartIndex = lngStartIndex + lngNoElementsToProcess
                End If
            Wend
            
            'Revise this if needed
            While theData(lngLastUnderValue).theKey > lngPivotValue
                lngLastUnderValue = lngLastUnderValue - 1
            Wend
        End If
            
        While lngFirstOverValue < lngLastUnderValue
        
            'Scan through the data
            lngStartIndex = lngFirstOverValue
            While lngStartIndex < lngLastUnderValue
            
                'Rotate every second run of overs with the following set of unders
                'Find the run of overs
                lngOverEnd1 = lngStartIndex
                Do
                    If lngOverEnd1 + 1 <= lngLastElement Then
                        If theData(lngOverEnd1 + 1).theKey > lngPivotValue Then
                            lngOverEnd1 = lngOverEnd1 + 1
                            blKeepLooping = True
                        Else
                            blKeepLooping = False
                        End If
                    Else
                        blKeepLooping = False
                    End If
                Loop Until Not blKeepLooping
                'Find the run of unders
                lngUnderEnd1 = lngOverEnd1 + 1
                'Icky visual basic if condition because vb evaluates all parts of the if every time
                Do
                    If lngUnderEnd1 + 1 <= lngLastElement Then
                        If theData(lngUnderEnd1 + 1).theKey <= lngPivotValue Then
                            blKeepLooping = True
                            lngUnderEnd1 = lngUnderEnd1 + 1
                        Else
                            blKeepLooping = False
                        End If
                    Else
                        blKeepLooping = False
                    End If
                Loop Until Not blKeepLooping
                
                'Rotate these two blocks around
                i = lngStartIndex
                j = lngOverEnd1
                While i < j
                    lngSwapValue = theData(i).theKey
                    theData(i).theKey = theData(j).theKey
                    theData(j).theKey = lngSwapValue
                
                    lngSwapValue = theData(i).originalOrder
                    theData(i).originalOrder = theData(j).originalOrder
                    theData(j).originalOrder = lngSwapValue
                    
                    i = i + 1
                    j = j - 1
                Wend
                
                i = lngOverEnd1 + 1
                j = lngUnderEnd1
                While i < j
                    lngSwapValue = theData(i).theKey
                    theData(i).theKey = theData(j).theKey
                    theData(j).theKey = lngSwapValue
                
                    lngSwapValue = theData(i).originalOrder
                    theData(i).originalOrder = theData(j).originalOrder
                    theData(j).originalOrder = lngSwapValue
                    
                    i = i + 1
                    j = j - 1
                Wend
                
                i = lngStartIndex
                j = lngUnderEnd1
                While i < j
                    lngSwapValue = theData(i).theKey
                    theData(i).theKey = theData(j).theKey
                    theData(j).theKey = lngSwapValue
                
                    lngSwapValue = theData(i).originalOrder
                    theData(i).originalOrder = theData(j).originalOrder
                    theData(j).originalOrder = lngSwapValue
                    
                    i = i + 1
                    j = j - 1
                Wend
                
                'Adjust the range of items needing swapping
                If lngStartIndex = lngFirstOverValue Then
                    'If this was the first run of overs than the first run of overs has moved back
                    lngFirstOverValue = lngFirstOverValue + (lngUnderEnd1 - lngOverEnd1)
                End If
                If lngUnderEnd1 = lngLastUnderValue Then
                    'If this was the last run of unders then it has moved forward
                    lngLastUnderValue = lngLastUnderValue - (lngOverEnd1 - lngStartIndex + 1)
                End If
                
                If lngUnderEnd1 < lngLastUnderValue Then
                    'Another run of each, dont swap these yet
                    'this causes the runs to bunch up in a binary way
                    
                    'Find the run of overs
                    lngOverEnd2 = lngUnderEnd1 + 1
                    'Icky visual basic if condition because vb evaluates all parts of the if every time
                    Do
                        If lngOverEnd2 + 1 <= lngLastElement Then
                            If theData(lngOverEnd2 + 1).theKey > lngPivotValue Then
                                blKeepLooping = True
                                lngOverEnd2 = lngOverEnd2 + 1
                            Else
                                blKeepLooping = False
                            End If
                        Else
                            blKeepLooping = False
                        End If
                    Loop Until Not blKeepLooping
                    
                    'Find run of unders
                    lngUnderEnd2 = lngOverEnd2 + 1
                    'Icky visual basic if condition because vb evaluates all parts of the if every time
                    Do
                        If lngUnderEnd2 + 1 <= lngLastElement Then
                            If theData(lngUnderEnd2 + 1).theKey <= lngPivotValue Then
                                blKeepLooping = True
                                lngUnderEnd2 = lngUnderEnd2 + 1
                            Else
                                blKeepLooping = False
                            End If
                        Else
                            blKeepLooping = False
                        End If
                    Loop Until Not blKeepLooping
                                    
                    'Point to the next start of overs
                    lngStartIndex = lngUnderEnd2 + 1
                Else
                    'Point to the next start of overs
                    lngStartIndex = lngUnderEnd1 + 1
                End If
            Wend
        
        Wend
    End If

    StableBinaryPivotCB = lngFirstOverValue - lngFirstElement

End Function



Function StableBinaryPivotSmallCBForward(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

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
            smallBufferCB(lngNoOvers).theKey = theData(lngI).theKey
            smallBufferCB(lngNoOvers).originalOrder = theData(lngI).originalOrder
            lngNoOvers = lngNoOvers + 1
        End If
    Next
    
    'Now put the over values back into the original array after the unders
    lngI = 0
    While lngI < lngNoOvers
        theData(lngFirstElement + lngNoUnders + lngI).theKey = smallBufferCB(lngI).theKey
        theData(lngFirstElement + lngNoUnders + lngI).originalOrder = smallBufferCB(lngI).originalOrder
        lngI = lngI + 1
    Wend

    StableBinaryPivotSmallCBForward = lngNoUnders

End Function

Function StableBinaryPivotSmallCBBackward(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

    'A function to pivot a small segment of data using a buffer.
    'The over values will be placed first

    Dim lngI As Long
    Dim lngNoUnders As Long
    Dim lngNoOvers As Long
    
    lngNoUnders = 0
    lngNoOvers = 0
    
    'Loop through the data and compact the under values to the front of the original array
    'and store the over values in the buffer
    For lngI = lngFirstElement To lngLastElement
        If theData(lngI).theKey <= lngPivotValue Then
            smallBufferCB(lngNoUnders).theKey = theData(lngI).theKey
            smallBufferCB(lngNoUnders).originalOrder = theData(lngI).originalOrder
            lngNoUnders = lngNoUnders + 1
        Else
            theData(lngFirstElement + lngNoOvers).theKey = theData(lngI).theKey
            theData(lngFirstElement + lngNoOvers).originalOrder = theData(lngI).originalOrder
            lngNoOvers = lngNoOvers + 1
        End If
    Next
    
    'Now put the over values back into the original array after the unders
    lngI = 0
    While lngI < lngNoUnders
        theData(lngFirstElement + lngNoOvers + lngI).theKey = smallBufferCB(lngI).theKey
        theData(lngFirstElement + lngNoOvers + lngI).originalOrder = smallBufferCB(lngI).originalOrder
        lngI = lngI + 1
    Wend

    StableBinaryPivotSmallCBBackward = lngNoUnders

End Function



