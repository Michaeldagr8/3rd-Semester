Attribute VB_Name = "modTraditionalQuickSort"
Option Explicit

Const QUICKSORTDEPTH2 = 32


Function QuickSort(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long)

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
                lngPivotPoint(1) = TraditionalPivot(theData, lngStackStarts(lngStackSize), lngStackEnds(lngStackSize), lngPivotValue(1))
                
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

Function TraditionalPivot(ByRef theData() As DataElement, ByVal lngFirstElement As Long, ByVal lngLastElement As Long, ByVal lngPivotValue As Long) As Long

    Dim lngRightIndex As Long
    Dim lngLeftIndex As Long
    
    'The complicated selection of the pivot value
    'has already guaranteed that there will be some values above the pivot
    'and some at or below the pivot
    
    lngRightIndex = lngLastElement
    lngLeftIndex = lngFirstElement
    
    While lngRightIndex >= lngLeftIndex
        
        While theData(lngRightIndex).theKey > lngPivotValue
            lngRightIndex = lngRightIndex - 1
        Wend
        While theData(lngLeftIndex).theKey <= lngPivotValue
            lngLeftIndex = lngLeftIndex + 1
        Wend
        
        If lngRightIndex > lngLeftIndex Then
            swapElements theData(lngLeftIndex), theData(lngRightIndex)
        End If
    Wend
    
    TraditionalPivot = lngLeftIndex - lngFirstElement

End Function

