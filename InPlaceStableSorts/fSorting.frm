VERSION 5.00
Begin VB.Form frmSorting 
   Caption         =   "Stable In Place Sorting"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox edBufferSizeIPMS 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   40
      Text            =   "64"
      ToolTipText     =   "1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox edBufferSizeSQBinTB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   38
      Text            =   "64"
      ToolTipText     =   "1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox edBufferSizeSQBinCB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   36
      Text            =   "64"
      ToolTipText     =   "1"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox cbxDoStableBinaryQuickSortTB 
      Caption         =   "Do Stable Binary Quick Sort (TB) O(NLogNLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   35
      Top             =   1680
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.CheckBox cbxDoStableBinaryQuickSortCB 
      Caption         =   "Do Stable Binary Quick Sort (CB) O(NLogNLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   34
      Top             =   960
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.CheckBox cbxDoCraigsSmoothSort 
      Caption         =   "Do Craigs Smooth Sort (Not Stable) O(NLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   33
      Top             =   5400
      Width           =   3975
   End
   Begin VB.CheckBox cbxDoHeapSort 
      Caption         =   "Do Traditional Heap Sort (Not Stable) O(NLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   4680
      Width           =   3855
   End
   Begin VB.CheckBox cbxDoSmoothSort 
      Caption         =   "Do Smooth Sort (Not Stable) O(NLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   31
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CheckBox cbxDoStableHeapSort 
      Caption         =   "Do Stable Heap Sort (Stable) O(N^2)"
      Height          =   375
      Left            =   2160
      TabIndex        =   30
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CheckBox cbxMergeSort 
      Caption         =   "DoTraditional Merge Sort (Stable) O(NLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Top             =   3240
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.CheckBox cbxUnstableQuickSort 
      Caption         =   "Do Traditional Quick Sort (Unstable) O(NLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   4320
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CommandButton cmdShowElements 
      Caption         =   "Show"
      Height          =   375
      Left            =   7560
      TabIndex        =   28
      Top             =   6060
      Width           =   1695
   End
   Begin VB.Frame frameShowElements 
      Caption         =   "Show Elements"
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   9255
      Begin VB.TextBox edToElement 
         Height          =   285
         Left            =   4440
         TabIndex        =   27
         Text            =   "200"
         Top             =   210
         Width           =   1095
      End
      Begin VB.TextBox edFromElement 
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Text            =   "0"
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblToElement 
         Alignment       =   1  'Right Justify
         Caption         =   "To Element:"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   260
         Width           =   1335
      End
      Begin VB.Label lblFromElement 
         Alignment       =   1  'Right Justify
         Caption         =   "From Element:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   260
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8040
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox edNoDistinctValues 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   19
      Text            =   "100000"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox edBufferSizeSQB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Text            =   "64"
      ToolTipText     =   "1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox edResults 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6600
      Width           =   9255
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox edRepetitions 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   14
      Text            =   "1"
      ToolTipText     =   "1"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox edDataSetSize 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   13
      Text            =   "30000"
      ToolTipText     =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox edSeed 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   12
      Text            =   "1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CheckBox cbxDoInsertionSort 
      Caption         =   "Do Insertion Sort (Stable) O(N^2)"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CheckBox cbxDoInPlaceMerge 
      Caption         =   "Do In-Place Merge Sort (Stable) O(NLogNLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.CheckBox cbxDoStableQuickSort 
      Caption         =   "Do Stable Buffered Quick Sort O(NLogNLogN)"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.Frame frameOriginalOrder 
      Caption         =   "Original Data Order"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      Begin VB.OptionButton opMostlyInOrder 
         Caption         =   "95% In Order"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton opReverse 
         Caption         =   "Reverse Order"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton opInOrder 
         Caption         =   "In Order"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton opRandom 
         Caption         =   "Random Order"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label lblBufferSizeIPMS 
      Alignment       =   1  'Right Justify
      Caption         =   "Buffer Size:"
      Height          =   255
      Left            =   2160
      TabIndex        =   41
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Label lblBufferSizeSQBinTB 
      Alignment       =   1  'Right Justify
      Caption         =   "Buffer Size:"
      Height          =   255
      Left            =   2160
      TabIndex        =   39
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label lblBufferSizeSQBinCB 
      Alignment       =   1  'Right Justify
      Caption         =   "Buffer Size:"
      Height          =   255
      Left            =   2160
      TabIndex        =   37
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   8280
      Width           =   9255
   End
   Begin VB.Label lblNoDifferentValues 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Distinct Values:"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   645
      Width           =   2055
   End
   Begin VB.Label lblBufferSizeSQB 
      Alignment       =   1  'Right Justify
      Caption         =   "Buffer Size:"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label lblNoRepetitions 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Repetitions:"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblDataSetSize 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Set Size:"
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label lblSeed 
      Alignment       =   1  'Right Justify
      Caption         =   "Random Seed:"
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   285
      Width           =   1215
   End
End
Attribute VB_Name = "frmSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const IDNOSORT = 0
Const IDUNSTABLEQUICKSORT = 1
Const IDSTABLEBUFFEREDQUICKSORT = 2
Const IDSTABLEBINARYQUICKSORTCB = 3
Const IDSTABLEBINARYQUICKSORTTB = 4
Const IDINPLACEMERGESORT = 5
Const IDTRADITIONALMERGESORT = 6
Const IDINSERTIONSORT = 7
Const IDHEAPSORT = 8
Const IDSTABLEHEAPSORT = 9
Const IDSMOOTHSORT = 10
Const IDCRAIGSMOOTHSORT = 11

Dim lngWidthDiff As Long
Dim lngTopDiff As Long
Dim lngHeightDiff As Long

Dim theData() As DataElement
Dim freshData() As DataElement



Private Sub cmdClear_Click()
    Me.edResults.Text = ""
End Sub

Private Sub cmdGo_Click()

    Dim blOk As Boolean
    Dim strMessage As String
    
    Dim lngSeed As Long
    Dim lngDataSetSize As Long
    Dim lngNoRepetitions As Long
    Dim lngBufferSizeSQB As Long
    Dim lngBufferSizeSQBinCB As Long
    Dim lngBufferSizeSQBinTB As Long
    Dim lngBufferSizeIPMS As Long
    Dim lngNoDistinctValues As Long
    
    Dim dblSortTimes(11) As Double
    Dim strThisResult As String
    
    blOk = True
    strMessage = ""
    
    lngSeed = VerifyEdit(Me.edSeed, "random seed", blOk, strMessage, 1, 32767)
    lngNoDistinctValues = VerifyEdit(Me.edNoDistinctValues, "number of distinct values", blOk, strMessage, 0, 1073741823)
    lngDataSetSize = VerifyEdit(Me.edDataSetSize, "data set size", blOk, strMessage, 1, 1073741823)
    lngNoRepetitions = VerifyEdit(Me.edRepetitions, "number of repetitions", blOk, strMessage, 1, 100000)
    lngBufferSizeSQB = VerifyEdit(Me.edBufferSizeSQB, "(buffered) stable quick sort buffer size", blOk, strMessage, 4, 100000)
    lngBufferSizeSQBinCB = VerifyEdit(Me.edBufferSizeSQBinCB, "(CB) binary stable quick sort buffer size", blOk, strMessage, 0, 100000)
    lngBufferSizeSQBinTB = VerifyEdit(Me.edBufferSizeSQBinTB, "(TB) binary stable quick sort buffer size", blOk, strMessage, 0, 100000)
    lngBufferSizeIPMS = VerifyEdit(Me.edBufferSizeIPMS, "in place merge sort buffer size", blOk, strMessage, 0, 100000)
            
    Screen.MousePointer = vbHourglass
            
    If Me.cbxUnstableQuickSort.Value = False And Me.cbxDoInPlaceMerge.Value = False And Me.cbxDoInsertionSort.Value = False And Me.cbxDoStableQuickSort.Value = False And Me.cbxDoStableBinaryQuickSortCB.Value = False And Me.cbxDoStableBinaryQuickSortTB.Value = False And Me.cbxMergeSort.Value = False And Me.cbxDoHeapSort.Value = False And Me.cbxDoStableHeapSort.Value = False And Me.cbxDoSmoothSort.Value = False And Me.cbxDoCraigsSmoothSort.Value = False Then
        blOk = False
        strMessage = strMessage + Chr(13) + Chr(10) + "- No sorting methods have been chosen."
    End If
    
    If lngDataSetSize > 1000 And (Me.cbxDoInsertionSort.Value <> False Or Me.cbxDoStableHeapSort.Value <> False) Then
        If MsgBox("Warning, doing an O(N^2) sort on this many elements can take a long time.  Are you sure?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            blOk = False
            strMessage = strMessage + Chr(13) + Chr(10) + "- Operation cancelled because there are too many elements for an O(N^2) sort."
        End If
    End If
    
    If blOk Then
        
        'Get the memory required for the stable quick sort and in place merge sort
        PIVOTBUFFERSIZE = lngBufferSizeSQB
        SHUFFLENOBLOCKS = PIVOTBUFFERSIZE * 2 + 1
        RedimStableQuickSortArrays
        
        SMALLSEGMENTSIZECB = lngBufferSizeSQBinCB
        ReDim smallBufferCB(SMALLSEGMENTSIZECB) As DataElement
        
        SMALLSEGMENTSIZETB = lngBufferSizeSQBinTB
        ReDim smallBufferTB(SMALLSEGMENTSIZETB) As DataElement
        
        SMALLSEGMENTSIZEIPMS = lngBufferSizeIPMS
        ReDim smallBufferIPMS(SMALLSEGMENTSIZEIPMS) As DataElement
        
        GenerateData lngSeed, lngDataSetSize, lngNoDistinctValues
        
        If Me.cbxUnstableQuickSort.Value Then
            dblSortTimes(IDUNSTABLEQUICKSORT) = TestNow(IDUNSTABLEQUICKSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoStableQuickSort.Value Then
            dblSortTimes(IDSTABLEBUFFEREDQUICKSORT) = TestNow(IDSTABLEBUFFEREDQUICKSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoStableBinaryQuickSortCB.Value Then
            dblSortTimes(IDSTABLEBINARYQUICKSORTCB) = TestNow(IDSTABLEBINARYQUICKSORTCB, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoStableBinaryQuickSortTB.Value Then
            dblSortTimes(IDSTABLEBINARYQUICKSORTTB) = TestNow(IDSTABLEBINARYQUICKSORTTB, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoInPlaceMerge.Value Then
            dblSortTimes(IDINPLACEMERGESORT) = TestNow(IDINPLACEMERGESORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxMergeSort.Value Then
            dblSortTimes(IDTRADITIONALMERGESORT) = TestNow(IDTRADITIONALMERGESORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoInsertionSort.Value Then
            dblSortTimes(IDINSERTIONSORT) = TestNow(IDINSERTIONSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoHeapSort.Value Then
            dblSortTimes(IDHEAPSORT) = TestNow(IDHEAPSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoStableHeapSort.Value Then
            dblSortTimes(IDSTABLEHEAPSORT) = TestNow(IDSTABLEHEAPSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoSmoothSort.Value Then
            dblSortTimes(IDSMOOTHSORT) = TestNow(IDSMOOTHSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        If Me.cbxDoCraigsSmoothSort.Value Then
            dblSortTimes(IDCRAIGSMOOTHSORT) = TestNow(IDCRAIGSMOOTHSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB) - TestNow(IDNOSORT, lngDataSetSize, lngNoRepetitions, lngBufferSizeSQB)
        End If
        
        Me.lblStatus = ""
        Me.lblStatus.Refresh
                
        strThisResult = "Seed = " + CStr(lngSeed) + ", Size = " + CStr(lngDataSetSize) + ", Repetitions = " + CStr(lngNoRepetitions) + ", Distinct Values = " + CStr(lngNoDistinctValues)
        If Me.opInOrder.Value Then
            strThisResult = strThisResult + ", Data In Order"
        ElseIf Me.opReverse.Value Then
            strThisResult = strThisResult + ", Data In Reverse Order"
        ElseIf Me.opMostlyInOrder Then
            strThisResult = strThisResult + ", First 95% of Data In Order"
        Else
            strThisResult = strThisResult + ", Data in Random Order"
        End If
        
        If Me.cbxUnstableQuickSort.Value Then
            strThisResult = strThisResult + ", Unstable Quick Sort Time = " + Format(dblSortTimes(IDUNSTABLEQUICKSORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoStableQuickSort.Value Then
            strThisResult = strThisResult + ", Stable Buffered Quick Sort, Buffer = " + CStr(lngBufferSizeSQB) + ", Time = " + Format(dblSortTimes(IDSTABLEBUFFEREDQUICKSORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoStableBinaryQuickSortCB.Value Then
            strThisResult = strThisResult + ", Stable Binary Quick Sort (CB), Buffer = " + CStr(lngBufferSizeSQBinCB) + ", Time = " + Format(dblSortTimes(IDSTABLEBINARYQUICKSORTCB) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoStableBinaryQuickSortTB.Value Then
            strThisResult = strThisResult + ", Stable Binary Quick Sort (TB), Buffer = " + CStr(lngBufferSizeSQBinTB) + ", Time = " + Format(dblSortTimes(IDSTABLEBINARYQUICKSORTTB) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoInPlaceMerge.Value Then
            strThisResult = strThisResult + ", In Place Merge Sort, Buffer = " + CStr(lngBufferSizeIPMS) + ", Time = " + Format(dblSortTimes(IDINPLACEMERGESORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxMergeSort.Value Then
            strThisResult = strThisResult + ", Traditional Merge Sort Time = " + Format(dblSortTimes(IDTRADITIONALMERGESORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoInsertionSort.Value Then
            strThisResult = strThisResult + ", Insertion Sort Time = " + Format(dblSortTimes(IDINSERTIONSORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoHeapSort.Value Then
            strThisResult = strThisResult + ", Heap Sort Time = " + Format(dblSortTimes(IDHEAPSORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoStableHeapSort.Value Then
            strThisResult = strThisResult + ", Stable Heap Sort Time = " + Format(dblSortTimes(IDSTABLEHEAPSORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoSmoothSort.Value Then
            strThisResult = strThisResult + ", Smooth Sort Time = " + Format(dblSortTimes(IDSMOOTHSORT) / lngNoRepetitions, "0.0000")
        End If
        If Me.cbxDoCraigsSmoothSort.Value Then
            strThisResult = strThisResult + ", Craigs Smooth Sort Time = " + Format(dblSortTimes(IDCRAIGSMOOTHSORT) / lngNoRepetitions, "0.0000")
        End If
        
        If Me.edResults.Text = "" Then
            Me.edResults.Text = strThisResult
        Else
            Me.edResults.Text = Me.edResults.Text + Chr(13) + Chr(10) + strThisResult
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    If Not blOk Then
        MsgBox "Sorting not performed because:" + strMessage, vbExclamation, "Warning"
    End If
            
End Sub


Function VerifyEdit(ByRef edField As Control, ByVal strFieldName As String, ByRef blOk As Boolean, ByRef strMessage As String, ByVal lngMinVal As Long, ByVal lngMaxVal As Long) As Long

    Dim lngValue As Long
    Dim strText As String
    Dim blThisOk As Boolean
    
    blThisOk = True
    strText = edField.Text
    
    If Trim(strText) = "" Then
        strMessage = strMessage + Chr(13) + Chr(10) + "- A " + strFieldName + " must be entered."
        blOk = False
    Else
        On Error Resume Next
        lngValue = CLng(strText)
        If Err <> 0 Then
            strMessage = strMessage + Chr(13) + Chr(10) + "- The " + strFieldName + " must be a number between " + CStr(lngMinVal) + " and " + CStr(lngMaxVal) + "."
            blOk = False
            blThisOk = False
        End If
        On Error GoTo 0
        If blThisOk Then
            If lngValue < lngMinVal Or lngValue > lngMaxVal Then
                strMessage = strMessage + Chr(13) + Chr(10) + "- The " + strFieldName + " must be a number between " + CStr(lngMinVal) + " and " + CStr(lngMaxVal) + "."
                blOk = False
            Else
                edField.Text = CStr(lngValue)
            End If
        End If
    End If

    VerifyEdit = lngValue

End Function

Sub GenerateData(ByVal lngSeed As Long, ByVal lngDataSetSize As Long, ByVal lngNoDistinctValues As Long)

    Dim lngI As Long
    Dim lngJ As Long
    Dim lng95PCData As Long

    Me.lblStatus.Caption = "Generating Random Data"
    Me.lblStatus.Refresh

    ReDim theData(lngDataSetSize - 1) As DataElement
    ReDim freshData(lngDataSetSize - 1) As DataElement

    Rnd -1
    Randomize lngSeed
    
    For lngI = 0 To lngDataSetSize - 1
        freshData(lngI).theKey = Int(Rnd() * lngNoDistinctValues)
        freshData(lngI).originalOrder = lngI
    Next
    
    Me.lblStatus.Caption = "Sorting Random Data"
    Me.lblStatus.Refresh
    
    If Me.opInOrder.Value Or Me.opReverse.Value Then
        QuickSort freshData, 0, lngDataSetSize - 1
    End If
    If Me.opReverse.Value Then
        lngI = 0
        lngJ = lngDataSetSize - 1
        While lngI < lngJ
            swapElements freshData(lngI), freshData(lngJ)
            lngI = lngI + 1
            lngJ = lngJ - 1
        Wend
    End If
    If Me.opMostlyInOrder.Value Then
        lng95PCData = Int(lngDataSetSize * 0.95)
        QuickSort freshData, 0, lng95PCData
    End If

    For lngI = 0 To lngDataSetSize - 1
        freshData(lngI).originalOrder = lngI
    Next

    Me.lblStatus.Caption = "Data Generated"
    Me.lblStatus.Refresh

End Sub

Function TestNow(ByVal lngSortNo As Long, ByVal lngDataSetSize As Long, ByVal lngNoRepetitions As Long, ByVal lngBufferSize As Long) As Double

    Dim dblStartTime As Double
    Dim dblEndTime As Double
    Dim lngLoopNo As Long
    Dim lngElementNo As Long
    Dim blOk As Boolean
    
    Dim strSortName As String
    Dim lngDisplayTime As Long
    Dim lngNewDisplayTime As Long
    Dim strMessage As String
    
    blOk = True
    
    Select Case lngSortNo
    Case IDNOSORT
        strSortName = "No Sort"
    Case IDUNSTABLEQUICKSORT
        strSortName = "Unstable Quick Sort"
    Case IDSTABLEBUFFEREDQUICKSORT
        strSortName = "Stable Buffered Quick Sort"
    Case IDSTABLEBINARYQUICKSORTCB
        strSortName = "Stable Binary Quick Sort (CB)"
    Case IDSTABLEBINARYQUICKSORTTB
        strSortName = "Stable Binary Quick Sort (TB)"
    Case IDINPLACEMERGESORT
        strSortName = "In-Place Merge Sort"
    Case IDTRADITIONALMERGESORT
        strSortName = "Traditional Merge Sort"
    Case IDINSERTIONSORT
        strSortName = "Insertion Sort"
    Case IDHEAPSORT
        strSortName = "Traditional Heap Sort"
    Case IDSTABLEHEAPSORT
        strSortName = "Stable Heap Sort"
    Case IDSMOOTHSORT
        strSortName = "Smooth Sort"
    Case IDCRAIGSMOOTHSORT
        strSortName = "Craigs Smooth Sort"
    End Select
    
    lngDisplayTime = Int(Timer)
    Me.lblStatus = "Testing " + strSortName + " Rep 1"
    Me.lblStatus.Refresh
    
    dblStartTime = Timer
    
    For lngLoopNo = 1 To lngNoRepetitions
    
        lngNewDisplayTime = Int(Timer)
        If lngNewDisplayTime <> lngDisplayTime Then
            Me.lblStatus = "Testing " + strSortName + " Rep " + CStr(lngLoopNo)
            Me.lblStatus.Refresh
            lngDisplayTime = lngNewDisplayTime
        End If
            
    
        'Get a fresh copy of the data
        For lngElementNo = lngDataSetSize - 1 To 0 Step -1
            AssignElement theData(lngElementNo), freshData(lngElementNo)
        Next
        
        'Do the sort
        Select Case lngSortNo
        Case IDNOSORT
        Case IDUNSTABLEQUICKSORT
            QuickSort theData, 0, lngDataSetSize - 1
        Case IDSTABLEBUFFEREDQUICKSORT
            StableQuickSort theData, 0, lngDataSetSize - 1
        Case IDSTABLEBINARYQUICKSORTCB
            StableBinaryQuickSortCB theData, 0, lngDataSetSize - 1
        Case IDSTABLEBINARYQUICKSORTTB
            StableBinaryQuickSortTB theData, 0, lngDataSetSize - 1
        Case IDINPLACEMERGESORT
            InPlaceMergeSort 0, lngDataSetSize, theData
        Case IDTRADITIONALMERGESORT
            MergeSort 0, lngDataSetSize, theData
        Case IDINSERTIONSORT
            InsertSort 0, lngDataSetSize, theData
        Case IDHEAPSORT
            HeapSort theData, lngDataSetSize
        Case IDSTABLEHEAPSORT
            StableHeapSort theData, lngDataSetSize
        Case IDSMOOTHSORT
            SmoothSort theData
        Case IDCRAIGSMOOTHSORT
            CraigsSmoothSort theData, lngDataSetSize
        End Select
        
        'Test that it is sorted
        If lngSortNo <> IDNOSORT Then
            For lngElementNo = lngDataSetSize - 2 To 0 Step -1
                If theData(lngElementNo).theKey > theData(lngElementNo + 1).theKey Then
                    blOk = False
                    strMessage = strSortName + " failed, elements not sorted."
                Else
                    If lngSortNo <> IDUNSTABLEQUICKSORT And lngSortNo <> IDHEAPSORT And lngSortNo <> IDSMOOTHSORT And lngSortNo <> IDCRAIGSMOOTHSORT And theData(lngElementNo).theKey = theData(lngElementNo + 1).theKey And theData(lngElementNo).originalOrder > theData(lngElementNo + 1).originalOrder Then
                        blOk = False
                        strMessage = strSortName + " failed, elements not sorted."
                    End If
                End If
            Next
        End If
    Next
    
    dblEndTime = Timer
    
    If blOk Then
        Me.lblStatus = "Testing " + strSortName + " Complete"
    Else
        Me.lblStatus = strMessage
        MsgBox strMessage
    End If
    Me.lblStatus.Refresh
    
    
    TestNow = dblEndTime - dblStartTime

End Function

Private Sub cmdShowElements_Click()

    Dim lngFromElement As Long
    Dim lngToElement As Long
    Dim lngLowBounds As Long
    Dim lngHiBounds As Long
    Dim strMessage As String
    Dim blOk As Boolean
    
    Dim strRemainingResults As String
    Dim lngNoRows As Long
    Dim lngNextEnd As Long
    Dim lngCurrentPos As Long
    Dim strNewLine As String
    
    Dim lngI As Long
    
    blOk = True
    
    On Error Resume Next
    lngLowBounds = LBound(theData)
    If Err <> 0 Then
        strMessage = "You must do a sort before you can see the data."
        blOk = False
    End If
    On Error GoTo 0
    
    If blOk Then
        If LBound(freshData) < lngLowBounds Then
            lngLowBounds = LBound(freshData)
        End If
        lngHiBounds = UBound(theData)
        If UBound(freshData) > lngHiBounds Then
            lngHiBounds = UBound(freshData)
        End If
        lngHiBounds = lngHiBounds - 1
    End If
    
    If blOk Then
        lngFromElement = VerifyEdit(Me.edFromElement, "from element", blOk, strMessage, lngLowBounds, lngHiBounds)
        lngToElement = VerifyEdit(Me.edToElement, "to element", blOk, strMessage, lngFromElement, lngHiBounds)
    End If
    
    If blOk Then
        If lngToElement > lngFromElement + 200 Then
            blOk = False
            strMessage = "You may not display more than 200 elements at a time."
        Else
            If lngToElement < lngFromElement Then
                blOk = False
                strMessage = "The to element must be after the from element."
            End If
        End If
    End If
                
    If blOk Then
        'Count the number of rows already there
        strRemainingResults = Me.edResults.Text
        lngNoRows = 0
        lngCurrentPos = 1
        lngNextEnd = InStr(lngCurrentPos, strRemainingResults, Chr(13))
        While lngNextEnd <> 0
            lngNoRows = lngNoRows + 1
            lngCurrentPos = lngNextEnd + 1
            lngNextEnd = InStr(lngCurrentPos, strRemainingResults, Chr(13))
        Wend
        
        If lngNoRows > 210 Then
            lngCurrentPos = 1
            lngNextEnd = InStr(lngCurrentPos, strRemainingResults, Chr(13))
            While lngNoRows > 210
                lngNoRows = lngNoRows - 1
                lngCurrentPos = lngNextEnd + 1
                lngNextEnd = InStr(lngCurrentPos, strRemainingResults, Chr(13))
                Me.edResults.Text = strRemainingResults
            Wend
            strRemainingResults = Mid(strRemainingResults, lngCurrentPos + 1)
            Me.edResults.Text = strRemainingResults
            Me.edResults.Refresh
        End If
        
        If Me.edResults.Text <> "" Then
            Me.edResults.Text = Me.edResults.Text + Chr(13) + Chr(10) + Chr(13) + Chr(10)
        End If
        Me.edResults.Text = Me.edResults.Text + "                     Original       Original         Sorted         Sorted" + Chr(13) + Chr(10)
        Me.edResults.Text = Me.edResults.Text + "       Element            Key    Sequence No            Key    Sequence No" + Chr(13) + Chr(10)
        For lngI = lngFromElement To lngToElement
            strNewLine = Right("              " + CStr(lngI), 14) + " "
            If lngI >= LBound(freshData) And lngI <= UBound(freshData) Then
                strNewLine = strNewLine + Right("              " + CStr(freshData(lngI).theKey), 14) + " " + Right("              " + CStr(freshData(lngI).originalOrder), 14) + " "
            Else
                strNewLine = strNewLine + "                              "
            End If
            If lngI >= LBound(theData) And lngI <= UBound(theData) Then
                strNewLine = strNewLine + Right("              " + CStr(theData(lngI).theKey), 14) + " " + Right("              " + CStr(theData(lngI).originalOrder), 14) + " "
            Else
                strNewLine = strNewLine + "                              "
            End If
            strNewLine = strNewLine + Chr(13) + Chr(10)
            Me.edResults.Text = Me.edResults.Text + strNewLine
        Next
    Else
        MsgBox strMessage, vbExclamation, "Warning"
    End If

End Sub

Private Sub Form_Load()
    lngWidthDiff = Me.Width - Me.edResults.Width
    lngTopDiff = Me.Height - Me.lblStatus.Top
    lngHeightDiff = Me.Height - Me.edResults.Height
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.edResults.Width = Me.Width - lngWidthDiff
    Me.edResults.Height = Me.Height - lngHeightDiff
    Me.lblStatus.Top = Me.Height - lngTopDiff
    On Error GoTo 0
End Sub
