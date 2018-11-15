Attribute VB_Name = "modStableQuickSort"
Option Explicit

Const QUICKSORTDEPTH = 32

Global PIVOTBUFFERSIZE As Long
Global SHUFFLENOBLOCKS  As Long

'PivotFlexFast Variables
'These variables belong in the PivotFlexFast function but
'this program allows the user to increase its size to be gigantic.
'If this happens then constantly reallocating and freeing the data
'slows down the process
Dim lngBunchStarts() As Long
Dim lngBunchSizes() As Long
Dim lngBunchOrder() As Long
Dim lngNoBunches As Long

'AggregateBunches Variables
'These variables belong in the AggregateBunches function but
'this program allows the user to increase its size to be gigantic.
'If this happens then constantly reallocating and freeing the data
'slows down the process
Dim lngBlockDestStarts() As Long
Dim lngBlockSourceStarts() As Long
Dim lngBlockSourceEnds() As Long
Dim lngBlockMoveDist() As Long
Dim lngBlockMinFromBlock() As Long
Dim lngBlockMaxFromBlock() As Long
Dim lngNoBlocks As Long

Sub RedimStableQuickSortArrays()
    ReDim lngBunchStarts(PIVOTBUFFERSIZE) As Long
    ReDim lngBunchSizes(PIVOTBUFFERSIZE) As Long
    ReDim lngBunchOrder(PIVOTBUFFERSIZE) As Long
    
    ReDim lngBlockDestStarts(SHUFFLENOBLOCKS) As Long
    ReDim lngBlockSourceStarts(SHUFFLENOBLOCKS) As Long
    ReDim lngBlockSourceEnds(SHUFFLENOBLOCKS) As Long
    ReDim lngBlockMoveDist(SHUFFLENOBLOCKS) As Long
    ReDim lngBlockMinFromBlock(SHUFFLENOBLOCKS) As Long
    ReDim lngBlockMaxFromBlock(SHUFFLENOBLOCKS) As Long
End Sub

Function StableQuickSort(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long)

    'This function is a pretty standard iterative version of quicksort,
    'all of the logic for stability is in the pivot function.
    'Quicksort can be done simpler than this but I have used the
    'median of 3 method of picking the pivot and have guaranteed
    'that there is always a value greater than the pivot.

    Dim lngStackStarts(QUICKSORTDEPTH) As Long
    Dim lngStackEnds(QUICKSORTDEPTH) As Long
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
    
            'Starting here is all about picking the pivot
            'using the median of 3 method
            lngPivotPoint(0) = Int(Rnd() * (lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) + 1)) + lngStackStarts(lngStackSize)
            lngPivotPoint(1) = Int(Rnd() * (lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) + 1)) + lngStackStarts(lngStackSize)
            lngPivotPoint(2) = Int(Rnd() * (lngStackEnds(lngStackSize) - lngStackStarts(lngStackSize) + 1)) + lngStackStarts(lngStackSize)
            lngPivotValue(0) = theData(lngPivotPoint(0)).theKey
            lngPivotValue(1) = theData(lngPivotPoint(1)).theKey
            lngPivotValue(2) = theData(lngPivotPoint(2)).theKey
            OrderLongs lngPivotValue(0), lngPivotValue(1)
            OrderLongs lngPivotValue(1), lngPivotValue(2)
            OrderLongs lngPivotValue(0), lngPivotValue(1)
            
            'At this point, we have three values but I want to make sure that they
            'are different and that there is a greater value than our pivot point.
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
                lngPivotPoint(1) = PivotFlexFast(theData, lngStackStarts(lngStackSize), lngStackEnds(lngStackSize), lngPivotValue(1))
                
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

Function PivotFlexFast(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

    'This function scans the data and fills up the bunch buffer with
    'details of runs of data that needs to be left alone
    
    Dim lngIndex As Long
    Dim lngNewBunchStart As Long
    Dim lngNewBunchEnd As Long
    
    Dim blKeepLooping As Boolean
    
    Dim lngRightMark As Long
    
    lngNoBunches = 0
    
    lngIndex = lngFirstElement
    While lngIndex <= lngLastElement
    
        'Skip over data that is greater than the pivot point
        'and needs to be shifted right
        blKeepLooping = False
        If lngIndex <= lngLastElement Then
            If theData(lngIndex).theKey > lngPivotValue Then
                blKeepLooping = True
            End If
        End If
        While blKeepLooping
            lngIndex = lngIndex + 1
            blKeepLooping = False
            If lngIndex <= lngLastElement Then
                If theData(lngIndex).theKey > lngPivotValue Then
                    blKeepLooping = True
                End If
            End If
        Wend
        
        If lngIndex <= lngLastElement Then
            'We are at the start of a bunch, start recording the bunch
            lngNewBunchStart = lngIndex
            'Skip past the bunch data
            blKeepLooping = False
            If lngIndex <= lngLastElement Then
                If theData(lngIndex).theKey <= lngPivotValue Then
                    blKeepLooping = True
                End If
            End If
            While blKeepLooping
                lngIndex = lngIndex + 1
                blKeepLooping = False
                If lngIndex <= lngLastElement Then
                    If theData(lngIndex).theKey <= lngPivotValue Then
                        blKeepLooping = True
                    End If
                End If
            Wend
            lngNewBunchEnd = lngIndex - 1
            
            'Add this bunch on to the list of bunches
            lngBunchStarts(lngNoBunches) = lngNewBunchStart
            lngBunchSizes(lngNoBunches) = lngNewBunchEnd - lngNewBunchStart + 1
            lngBunchOrder(lngNoBunches) = 1
            lngNoBunches = lngNoBunches + 1
            
            If lngNoBunches >= PIVOTBUFFERSIZE Then
                'If we have filled up our buffer of bunches, we need to aggregate them together
                AggregateBunches False, lngFirstElement, lngLastElement, lngBunchStarts, lngBunchSizes, lngBunchOrder, lngNoBunches, theData, lngPivotValue
            End If
            
        End If
        
    Wend
    
    'Aggregate all remaining bunches
    lngRightMark = lngLastElement + 1
    If lngNoBunches > 0 Then
        AggregateBunches True, lngFirstElement, lngLastElement, lngBunchStarts, lngBunchSizes, lngBunchOrder, lngNoBunches, theData, lngPivotValue
        lngRightMark = lngFirstElement + lngBunchSizes(0)
    End If

    PivotFlexFast = lngRightMark - lngFirstElement

End Function

Function AggregateBunches(ByVal blFinal As Boolean, ByVal lngStartLeft As Long, ByVal lngFarRight As Long, ByRef lngBunchStarts() As Long, ByRef lngBunchSizes() As Long, ByRef lngBunchOrder() As Long, ByRef lngNoBunches As Long, ByRef theData() As DataElement, ByVal lngPivotValue As Long)

    Dim lngStartAggregatingAt As Long
    Dim blKeepLooking As Boolean
    Dim lngNextOrder As Long
    
    Dim lngBunchNo As Long
    Dim lngCurrentIndex As Long
    Dim lngLastEnd As Long
    Dim lngThisMove As Long
    Dim lngThisSize As Long
    
    Dim lngHalfNoBlocks As Long
    Dim lngCurrentFromBlockNo As Long
    
    Dim lngBlockNo As Long
    
    Dim lngTestIndex As Long
    Dim lngStartIndex As Long
    Dim lngMiddleIndex As Long
    Dim lngEndIndex As Long
    Dim blNeedsToBeMoved As Boolean
    
    Dim tempData As DataElement
    Dim lngComesFromBlock As Long
    Dim lngComesFromIndex As Long
    Dim lngGoesToBlock As Long
    Dim lngGoesToIndex As Long
    
    'Work out where to start aggregating at
    If blFinal Then
        'If this is the final aggregation then aggregate all of the bunches to the
        'start of the array.  Job done
        lngStartIndex = lngStartLeft
        lngStartAggregatingAt = 0
        lngNextOrder = lngBunchOrder(0) + 1
    Else
        'Otherwise we need to aggregate them in an efficient way.
        'The lngBunchOrder array contains an indication of how many times
        'the bunch has been aggregated.
        
        'We need to aggregate at least two runs together of the same order
        lngStartAggregatingAt = PIVOTBUFFERSIZE - 1
        Do
            If lngStartAggregatingAt = 0 Then
                blKeepLooking = False
            Else
                If lngBunchOrder(lngStartAggregatingAt - 1) = lngBunchOrder(lngStartAggregatingAt) Or (lngBunchOrder(lngStartAggregatingAt - 1) > lngBunchOrder(lngStartAggregatingAt) + 1 And lngBunchOrder(lngStartAggregatingAt) > 1) Then
                    'Found two runs of the same order, or an increment of order > 1
                    blKeepLooking = False
                Else
                    blKeepLooking = True
                    lngStartAggregatingAt = lngStartAggregatingAt - 1
                End If
            End If
        Loop Until Not blKeepLooking
        'Now find the first run of this order
        Do
            If lngStartAggregatingAt = 0 Then
                blKeepLooking = False
            Else
                If lngBunchOrder(lngStartAggregatingAt - 1) <> lngBunchOrder(lngStartAggregatingAt) Then
                    'Found the first run of this order
                    blKeepLooking = False
                Else
                    blKeepLooking = True
                    lngStartAggregatingAt = lngStartAggregatingAt - 1
                End If
            End If
        Loop Until Not blKeepLooking
        
        
        lngNextOrder = lngBunchOrder(lngStartAggregatingAt) + 1
        lngStartIndex = lngBunchStarts(lngStartAggregatingAt)
    End If
    lngEndIndex = lngBunchStarts(lngNoBunches - 1) + lngBunchSizes(lngNoBunches - 1) - 1
    
    'We are going to aggregate all of the bunches starting at lngStartAggregatingAt
    
    'Create a full set of blocks in both directions, blocks are created for runs going left
    'and right.  The blocks are ordered by their destination location.
    'First do the bunches that are going left
    lngCurrentIndex = lngStartIndex
    lngNoBlocks = 0
    lngBunchNo = lngStartAggregatingAt
    While lngBunchNo < lngNoBunches
        lngThisMove = lngCurrentIndex - lngBunchStarts(lngBunchNo)
        If lngThisMove <> 0 And lngBunchSizes(lngBunchNo) <> 0 Then
            lngBlockDestStarts(lngNoBlocks) = lngCurrentIndex
            lngBlockSourceStarts(lngNoBlocks) = lngBunchStarts(lngBunchNo)
            lngBlockSourceEnds(lngNoBlocks) = lngBunchStarts(lngBunchNo) + lngBunchSizes(lngBunchNo) - 1
            lngBlockMoveDist(lngNoBlocks) = lngThisMove
            lngNoBlocks = lngNoBlocks + 1
        End If
        lngCurrentIndex = lngCurrentIndex + lngBunchSizes(lngBunchNo)
        lngBunchNo = lngBunchNo + 1
    Wend
    lngMiddleIndex = lngCurrentIndex
    'Now do the runs between the bunches that are to go right
    lngBunchNo = lngStartAggregatingAt
    lngLastEnd = lngStartIndex
    While lngBunchNo < lngNoBunches
        lngThisSize = lngBunchStarts(lngBunchNo) - lngLastEnd
        lngThisMove = lngCurrentIndex - lngLastEnd
        If lngThisMove <> 0 And lngThisSize <> 0 Then
            lngBlockDestStarts(lngNoBlocks) = lngCurrentIndex
            lngBlockSourceStarts(lngNoBlocks) = lngLastEnd
            lngBlockSourceEnds(lngNoBlocks) = lngBunchStarts(lngBunchNo) - 1
            lngBlockMoveDist(lngNoBlocks) = lngThisMove
            lngNoBlocks = lngNoBlocks + 1
        End If
        lngLastEnd = lngBunchStarts(lngBunchNo) + lngBunchSizes(lngBunchNo)
        lngCurrentIndex = lngCurrentIndex + lngThisSize
        lngBunchNo = lngBunchNo + 1
    Wend
    
    'Create a quick link index between blocks
    lngHalfNoBlocks = lngNoBlocks \ 2
    lngCurrentFromBlockNo = 0
    lngBlockNo = 0
    While lngBlockNo < lngNoBlocks
        If lngBlockNo = lngHalfNoBlocks Then
            lngCurrentFromBlockNo = 0
        End If
        
        'Get the minimum
        While lngBlockDestStarts(lngCurrentFromBlockNo + 1) <= lngBlockSourceStarts(lngBlockNo) And lngCurrentFromBlockNo < (lngNoBlocks - 1)
            lngCurrentFromBlockNo = lngCurrentFromBlockNo + 1
        Wend
        lngBlockMinFromBlock(lngBlockNo) = lngCurrentFromBlockNo
        
        'Get the maximum
        While lngBlockDestStarts(lngCurrentFromBlockNo + 1) <= lngBlockSourceEnds(lngBlockNo) And lngCurrentFromBlockNo < (lngNoBlocks - 1)
            lngCurrentFromBlockNo = lngCurrentFromBlockNo + 1
        Wend
        lngBlockMaxFromBlock(lngBlockNo) = lngCurrentFromBlockNo
        
        lngBlockNo = lngBlockNo + 1
    Wend
    
    'Keep shuffling until everything is done
    lngBlockNo = 0
    While lngBlockNo < lngHalfNoBlocks
        lngTestIndex = lngBlockSourceStarts(lngBlockNo)
        While lngTestIndex <= lngBlockSourceEnds(lngBlockNo)
    
            blNeedsToBeMoved = False
            If lngTestIndex < lngMiddleIndex Then
                If theData(lngTestIndex).theKey > lngPivotValue Then
                    blNeedsToBeMoved = True
                End If
            Else
                If theData(lngTestIndex).theKey <= lngPivotValue Then
                    blNeedsToBeMoved = True
                End If
            End If
            If blNeedsToBeMoved Then
                'Keep a copy of the data
                AssignElement tempData, theData(lngTestIndex)
            
                lngGoesToIndex = lngTestIndex
                
                'Find where it comes from
                lngComesFromBlock = BinarySearchForBlocks(lngGoesToIndex, lngBlockDestStarts, 0, lngNoBlocks - 1)
                lngComesFromIndex = lngBlockSourceStarts(lngComesFromBlock) + (lngGoesToIndex - lngBlockDestStarts(lngComesFromBlock))
                
                
                While lngComesFromIndex <> lngTestIndex
                    AssignElement theData(lngGoesToIndex), theData(lngComesFromIndex)

                    lngGoesToBlock = lngComesFromBlock
                    lngGoesToIndex = lngComesFromIndex
                    
                    lngComesFromBlock = BinarySearchForBlocks(lngGoesToIndex, lngBlockDestStarts, lngBlockMinFromBlock(lngGoesToBlock), lngBlockMaxFromBlock(lngGoesToBlock))
                    lngComesFromIndex = lngBlockSourceStarts(lngComesFromBlock) + (lngGoesToIndex - lngBlockDestStarts(lngComesFromBlock))
                Wend
                
                AssignElement theData(lngGoesToIndex), tempData
                
            End If
        
            lngTestIndex = lngTestIndex + 1
        Wend
        lngBlockNo = lngBlockNo + 1
    Wend

    'Record the bunches
    lngBunchNo = lngStartAggregatingAt + 1
    While lngBunchNo < lngNoBunches
        lngBunchSizes(lngStartAggregatingAt) = lngBunchSizes(lngStartAggregatingAt) + lngBunchSizes(lngBunchNo)
        lngBunchNo = lngBunchNo + 1
    Wend
    lngBunchStarts(lngStartAggregatingAt) = lngStartIndex
    lngBunchOrder(lngStartAggregatingAt) = lngNextOrder
    lngNoBunches = lngStartAggregatingAt + 1
        

End Function

Function BinarySearchForBlocks(ByVal lngValToFind As Long, ByRef lngBlockStarts() As Long, ByVal lngMinIndex As Long, ByVal lngMaxIndex As Long) As Long

    Dim lngMidIndex As Long
    Dim lngResult As Long

    While (lngMaxIndex - lngMinIndex) > 1
        lngMidIndex = (lngMaxIndex + lngMinIndex) \ 2
        If lngValToFind < lngBlockStarts(lngMidIndex) Then
            lngMaxIndex = lngMidIndex
        Else
            lngMinIndex = lngMidIndex
        End If
    Wend

    If lngValToFind < lngBlockStarts(lngMaxIndex) Then
        lngResult = lngMinIndex
    Else
        lngResult = lngMaxIndex
    End If
    
    BinarySearchForBlocks = lngResult

End Function

Function OrderLongs(ByRef lngOne As Long, ByRef lngTwo As Long)
    Dim lngTemp As Long
    
    If lngOne > lngTwo Then
        lngTemp = lngOne
        lngOne = lngTwo
        lngTwo = lngTemp
    End If
    
End Function

