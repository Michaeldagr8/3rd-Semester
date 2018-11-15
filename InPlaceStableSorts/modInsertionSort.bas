Attribute VB_Name = "modInsertionSort"
Option Explicit

Type DataElement
    theKey As Long
    originalOrder As Long
End Type
    
    
Sub InsertSort(ByVal lngFrom As Long, ByVal lngTo As Long, ByRef myData() As DataElement)

    Dim lngI As Long
    Dim lngJ As Long
    Dim blDone As Boolean
    
    
    If lngTo > lngFrom + 1 Then
        lngI = lngFrom + 1
        While lngI < lngTo
            lngJ = lngI
            blDone = False
            While lngJ > lngFrom And Not blDone
                If myData(lngJ).theKey < myData(lngJ - 1).theKey Then
                    swapElements myData(lngJ), myData(lngJ - 1)
                Else
                    blDone = True
                End If
                lngJ = lngJ - 1
            Wend
            lngI = lngI + 1
        Wend
    End If
    
End Sub
    

Sub AssignElement(ByRef toElement As DataElement, ByRef fromElement As DataElement)
    toElement.theKey = fromElement.theKey
    toElement.originalOrder = fromElement.originalOrder
End Sub

Sub swapElements(ByRef toElement As DataElement, ByRef fromElement As DataElement)

    Dim tempElement As DataElement
    
    tempElement.theKey = toElement.theKey
    tempElement.originalOrder = toElement.originalOrder
    toElement.theKey = fromElement.theKey
    toElement.originalOrder = fromElement.originalOrder
    fromElement.theKey = tempElement.theKey
    fromElement.originalOrder = tempElement.originalOrder
    
End Sub

