Attribute VB_Name = "modInPlaceMergeSort"
'Algorithm taken from Thomas Baudel's recursive C version
'http://thomas.baudel.name/Visualisation/VisuTri/inplacestablesort.html

'This version was re-written to improve clarity.
'It is a bit longer than a direct conversion from Thomas Baudel's code but
'executes with about the same speed.

Option Explicit

Global SMALLSEGMENTSIZEIPMS As Long
Global smallBufferIPMS() As DataElement


Sub InPlaceMergeSort(ByVal lngFrom As Long, ByVal lngTo As Long, ByRef myData() As DataElement)
    
    Dim lngThisMiddle As Long
    
    If lngTo - lngFrom < 12 Then
        InsertSort lngFrom, lngTo, myData
    Else
        lngThisMiddle = (lngFrom + lngTo) \ 2
        InPlaceMergeSort lngFrom, lngThisMiddle, myData
        InPlaceMergeSort lngThisMiddle, lngTo, myData
        InPlaceMerge lngFrom, lngThisMiddle, lngTo - 1, myData
    End If
    
End Sub
 
Sub InPlaceMerge(ByVal lngStartFirstStream As Long, ByVal lngStartSecondStream As Long, ByVal lngEndSecondStream As Long, ByRef myData() As DataElement)
    
    'This function merges two sets of pre-sorted data
    'I have called these streams.  The second stream is located in the
    'array immediately after the first finished.  Both streams are pre-sorted.
    
    'This is how the function works:
    '
    'Say we have a set of data:     Data   |a|b|c|d|e|f|a|b|e|f|
    '                                      ---------------------
    '                            Indexes    0 1 2 3 4 5 6 7 8 9
    '
    'This is in two sorted streams starting at element 0 through to 5 and then 6 through to 9
    '
    'As the first stream is bigger, it takes the second half of it:    Data |d|e|f|
    '                                                                       -------
    '                                                                        3 4 5
    '
    'And calls this the first block.
    '
    'It then identifies that the block:    Data |a|b|
    '                                           -----
    '                                            6 7
    '
    'from the second stream should go before it, and calls this the second block.
    '
    'It then pushes these two blocks around to get this data:
    '
    '                 Data   |a|b|c|a|b|d|e|f|e|f|
    '                        ---------------------
    '       Previous Index    0 1 2 6 7 3 4 5 8 9
    '                Index    0 1 2 3 4 5 6 7 8 9
    '
    'The function then calls itself recursively to merge indexes 0,1,2 with indexes 3,4.
    'The function also calls itself recursively to merge indexes 5,6,7 with indexes 8,9.
    
    Dim lngLengthFirstStream As Long
    Dim lngLengthSecondStream As Long
    
    Dim lngFirstBlockStart As Long
    Dim lngFirstBlockLength As Long
    Dim lngSecondBlockStart As Long
    Dim lngSecondBlockLength As Long
    
    lngLengthFirstStream = lngStartSecondStream - lngStartFirstStream
    lngLengthSecondStream = lngEndSecondStream - lngStartSecondStream + 1
    
    'If there is only one stream that has any size then the job is done, do nothing
    If lngLengthFirstStream <> 0 And lngLengthSecondStream <> 0 Then
        'If there are only two elements, then make sure that they are in order
        If lngLengthFirstStream + lngLengthSecondStream = 2 Then
            If myData(lngStartSecondStream).theKey < myData(lngStartFirstStream).theKey Then
                swapElements myData(lngStartSecondStream), myData(lngStartFirstStream)
            End If
        Else
            If lngLengthFirstStream + lngLengthSecondStream <= SMALLSEGMENTSIZEIPMS Then
                'Use a bit of extra space to speed it up
                InPlaceMergeSmall lngStartFirstStream, lngStartSecondStream, lngEndSecondStream, myData
            Else
                If lngLengthFirstStream > lngLengthSecondStream Then
                    'First block starts half way through the first stream
                    lngFirstBlockStart = lngStartFirstStream + (lngLengthFirstStream \ 2)
                    'And continues to the end of the first stream
                    lngFirstBlockLength = lngStartSecondStream - lngFirstBlockStart
                    
                    'Second block starts at the start of the second stream
                    lngSecondBlockStart = lngStartSecondStream
                    'And ends at a point so that everything in the first block
                    'should come after everything in the second block
                    lngSecondBlockLength = BinarySearchForFirstElementGEValue(myData(lngFirstBlockStart).theKey, lngStartSecondStream, lngEndSecondStream + 1, myData) - lngSecondBlockStart
                Else
                    'Second block starts at the start of the second stream
                    lngSecondBlockStart = lngStartSecondStream
                    'And continues to half way through the second stream
                    lngSecondBlockLength = lngLengthSecondStream \ 2
                    
                    'First block starts so that everything in the first block should
                    'come after everything in the second block
                    lngFirstBlockStart = BinarySearchForFirstElementGTValue(myData(lngSecondBlockStart + lngSecondBlockLength).theKey, lngStartFirstStream, lngStartSecondStream, myData)
                    'First block continues to the end of the first stream
                    lngFirstBlockLength = lngStartSecondStream - lngFirstBlockStart
                End If
                
                'Shuffle block one and block two around
                PushBlock lngFirstBlockStart, lngFirstBlockLength, lngSecondBlockStart + lngSecondBlockLength - lngFirstBlockLength, myData
                
                'Recursively call itself to merge the two groups
                InPlaceMerge lngStartFirstStream, lngFirstBlockStart, lngFirstBlockStart + lngSecondBlockLength - 1, myData
                InPlaceMerge lngSecondBlockStart + lngSecondBlockLength - lngFirstBlockLength, lngSecondBlockStart + lngSecondBlockLength, lngEndSecondStream, myData
            End If
        End If
    End If
End Sub
 
 

Function BinarySearchForFirstElementGEValue(ByVal lngValToFind As Long, ByVal lngFrom As Long, ByVal lngTo As Long, ByRef theData() As DataElement) As Long

    Dim lngLow As Long
    Dim lngHi As Long
    Dim lngMiddle As Long
    Dim lngResult As Long
    
    lngLow = lngFrom
    lngHi = lngTo - 1
    
    While (lngHi - lngLow) > 1
        lngMiddle = (lngLow + lngHi) \ 2
        If theData(lngMiddle).theKey >= lngValToFind Then
            lngHi = lngMiddle
        Else
            lngLow = lngMiddle
        End If
    Wend
    
    If theData(lngLow).theKey >= lngValToFind Then
        lngResult = lngLow
    Else
        If theData(lngHi).theKey >= lngValToFind Then
            lngResult = lngHi
        Else
            lngResult = lngHi + 1
        End If
    End If
    
    BinarySearchForFirstElementGEValue = lngResult

End Function


Function BinarySearchForFirstElementGTValue(ByVal lngValToFind As Long, ByVal lngFrom As Long, ByVal lngTo As Long, ByRef theData() As DataElement) As Long

    Dim lngLow As Long
    Dim lngHi As Long
    Dim lngMiddle As Long
    Dim lngResult As Long
    
    lngLow = lngFrom
    lngHi = lngTo - 1
    
    While (lngHi - lngLow) > 1
        lngMiddle = (lngLow + lngHi) \ 2
        If theData(lngMiddle).theKey > lngValToFind Then
            lngHi = lngMiddle
        Else
            lngLow = lngMiddle
        End If
    Wend
    
    If theData(lngLow).theKey > lngValToFind Then
        lngResult = lngLow
    Else
        If theData(lngHi).theKey > lngValToFind Then
            lngResult = lngHi
        Else
            lngResult = lngHi + 1
        End If
    End If
    
    BinarySearchForFirstElementGTValue = lngResult

End Function

Sub PushBlock(ByVal lngBlock1Start As Long, ByVal lngBlock1Length As Long, ByVal lngBlock1Dest As Long, ByRef myData() As DataElement)

    'Function to move a block of memory forward.
    'In doing so, another block of memory must be moved back.
    Dim lngBlock2Length As Long
    
    Dim lngNumberOfSeperateLoops As Long
    Dim lngLoopNo As Long
    
    Dim lngSavedElement As DataElement
    Dim lngCurrentIndex As Long
    Dim lngSourceIndex As Long
    Dim lngCurrentStartIndex As Long
    
    'If we actually have moving to do
    If lngBlock1Length <> 0 And lngBlock1Dest <> lngBlock1Start Then
        'Work out the details of the block going the opposite direction
        'Moving block 1 forwards
        lngBlock2Length = lngBlock1Dest - lngBlock1Start
        
        'This algorithm follows a path through the data being moved shuffling
        'as it goes.  The path is guarenteed to return to the starting position.
        'It is not guaranteed to hit every single element though.
        'We often need to repeat the process starting at a different start point,
        'the number of these start points required is the greatest common denominator
        'between the total size of the data affected and one of the block sizes.
        lngNumberOfSeperateLoops = GreatestCommonDenominator(lngBlock1Length + lngBlock2Length, lngBlock1Length)
        lngLoopNo = 0
        While lngLoopNo < lngNumberOfSeperateLoops
        
            lngCurrentStartIndex = lngBlock1Start + lngLoopNo
        
            lngCurrentIndex = lngCurrentStartIndex
            lngSourceIndex = lngCurrentIndex + lngBlock1Length
            
            'Continually find where the data is to come from and
            'move the data until we complete a loop
            AssignElement lngSavedElement, myData(lngCurrentIndex)
            While lngSourceIndex <> lngCurrentStartIndex
                AssignElement myData(lngCurrentIndex), myData(lngSourceIndex)
                
                lngCurrentIndex = lngSourceIndex
                If lngCurrentIndex >= lngBlock1Dest Then
                    lngSourceIndex = lngCurrentIndex - lngBlock2Length
                Else
                    lngSourceIndex = lngCurrentIndex + lngBlock1Length
                End If
            Wend
            AssignElement myData(lngCurrentIndex), lngSavedElement
            
            lngLoopNo = lngLoopNo + 1
        Wend
    End If
End Sub

Function GreatestCommonDenominator(ByVal lngM As Long, ByVal lngN As Long) As Long
    
    Dim lngT As Long
   
    While lngN <> 0
        lngT = lngM Mod lngN
        lngM = lngN
        lngN = lngT
    Wend
    
    GreatestCommonDenominator = lngM
    
End Function

Sub InPlaceMergeSmall(ByVal lngStartFirstStream As Long, ByVal lngStartSecondStream As Long, ByVal lngEndSecondStream As Long, ByRef myData() As DataElement)

    Dim lngI As Long
    Dim lngJ As Long
    Dim lngIx As Long

    lngI = lngStartFirstStream
    lngJ = lngStartSecondStream
    lngIx = 0
    
    'Merge them into the small buffer
    While lngI < lngStartSecondStream And lngJ <= lngEndSecondStream
        If myData(lngI).theKey <= myData(lngJ).theKey Then
            smallBufferIPMS(lngIx).theKey = myData(lngI).theKey
            smallBufferIPMS(lngIx).originalOrder = myData(lngI).originalOrder
            lngIx = lngIx + 1
            lngI = lngI + 1
        Else
            smallBufferIPMS(lngIx).theKey = myData(lngJ).theKey
            smallBufferIPMS(lngIx).originalOrder = myData(lngJ).originalOrder
            lngIx = lngIx + 1
            lngJ = lngJ + 1
        End If
    Wend
    While lngI < lngStartSecondStream
        smallBufferIPMS(lngIx).theKey = myData(lngI).theKey
        smallBufferIPMS(lngIx).originalOrder = myData(lngI).originalOrder
        lngIx = lngIx + 1
        lngI = lngI + 1
    Wend
    While lngJ <= lngEndSecondStream
        smallBufferIPMS(lngIx).theKey = myData(lngJ).theKey
        smallBufferIPMS(lngIx).originalOrder = myData(lngJ).originalOrder
        lngIx = lngIx + 1
        lngJ = lngJ + 1
    Wend

    'Copy all of the data back to the array
    lngI = 0
    lngIx = lngStartFirstStream
    While lngIx <= lngEndSecondStream
        myData(lngIx).theKey = smallBufferIPMS(lngI).theKey
        myData(lngIx).originalOrder = smallBufferIPMS(lngI).originalOrder
        lngIx = lngIx + 1
        lngI = lngI + 1
    Wend

End Sub

 

