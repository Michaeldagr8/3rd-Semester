Attribute VB_Name = "modTraditionalMergeSort"
Option Explicit

Dim datMergeBuffer() As DataElement


Sub MergeSort(ByVal lngFrom As Long, ByVal lngTo As Long, ByRef myData() As DataElement)

    Dim lngI As Long
    Dim lngTheBufferNo As Long

    ReDim datMergeBuffer(2, lngTo) As DataElement
    
    'Copy all of the data into buffer 1
    'This would not be necessary in C as the original array could be used and just referenced
    lngI = lngFrom
    While lngI < lngTo
        AssignElement datMergeBuffer(0, lngI), myData(lngI)
        lngI = lngI + 1
    Wend
        
    lngTheBufferNo = MergeSortRecur(lngFrom, lngTo)
                
    'Copy all of the data back
    lngI = lngFrom
    While lngI < lngTo
        AssignElement myData(lngI), datMergeBuffer(lngTheBufferNo, lngI)
        lngI = lngI + 1
    Wend
    
    
End Sub

Function MergeSortRecur(ByVal lngFrom As Long, ByVal lngTo As Long) As Long

    
    Dim lngThisMiddle As Long
    Dim lngBufferNo1 As Long
    Dim lngBufferNo2 As Long
    Dim lngDestBuffer As Long
    
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim lngDestIndex As Long
    
    If lngTo - lngFrom <= 2 Then
        If lngTo - lngFrom = 2 Then
            If datMergeBuffer(0, lngFrom).theKey > datMergeBuffer(0, lngFrom + 1).theKey Then
                swapElements datMergeBuffer(0, lngFrom), datMergeBuffer(0, lngFrom + 1)
            End If
        End If
        lngDestBuffer = 0
    Else
        'Sort the two halves of the data
        lngThisMiddle = (lngFrom + lngTo) \ 2
        lngBufferNo1 = MergeSortRecur(lngFrom, lngThisMiddle)
        lngBufferNo2 = MergeSortRecur(lngThisMiddle, lngTo)
        
        'Merge the two halves back together
        'Merge them into the buffer that is not buffer 1
        lngDestBuffer = 1 - lngBufferNo1
        
        'Merge them
        lngIndex1 = lngFrom
        lngIndex2 = lngThisMiddle
        lngDestIndex = lngFrom
        While lngIndex1 < lngThisMiddle And lngIndex2 < lngTo
            If datMergeBuffer(lngBufferNo1, lngIndex1).theKey <= datMergeBuffer(lngBufferNo2, lngIndex2).theKey Then
                AssignElement datMergeBuffer(lngDestBuffer, lngDestIndex), datMergeBuffer(lngBufferNo1, lngIndex1)
                lngIndex1 = lngIndex1 + 1
            Else
                AssignElement datMergeBuffer(lngDestBuffer, lngDestIndex), datMergeBuffer(lngBufferNo2, lngIndex2)
                lngIndex2 = lngIndex2 + 1
            End If
            lngDestIndex = lngDestIndex + 1
        Wend
        While lngIndex1 < lngThisMiddle
            AssignElement datMergeBuffer(lngDestBuffer, lngDestIndex), datMergeBuffer(lngBufferNo1, lngIndex1)
            lngIndex1 = lngIndex1 + 1
            lngDestIndex = lngDestIndex + 1
        Wend
        While lngIndex2 < lngTo
            AssignElement datMergeBuffer(lngDestBuffer, lngDestIndex), datMergeBuffer(lngBufferNo2, lngIndex2)
            lngIndex2 = lngIndex2 + 1
            lngDestIndex = lngDestIndex + 1
        Wend
        
    End If
    
    MergeSortRecur = lngDestBuffer
    
End Function



