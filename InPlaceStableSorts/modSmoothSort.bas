Attribute VB_Name = "modSmoothSort"
Option Explicit


'Smooth sort was invented by Edsger Dijkstra and is based on a heap
'sort.  The difference being that this algorithm takes advantage of
'already sorted data and runs faster in that case.
'
'The following VB code was converted to VB by Ellis Dee from www.vbforums.com
'and the names of the variables and comments were changed by
'Craig Brown 22/5/08 to make the whole thing comprehendable.
'
'Smoothsort is based on HeapSort but is more complicated.  I
'recommend that you understand HeapSort before attempting to understand
'SmoothSort.
'
'The algorithm for HeapSort is not too complicated.  However Dijkstra
'optimised the algorithm and replace logical operations with faster
'subtle mathematical short-cuts.  In addition, the meanings of his
'variables b, c, p, r etc are not clear.  In short, Dijkstra's original
'pseudocode is terse and difficult to understand.  I hope that this version,
'retaining all of the original logic, is easier to understand.
'
'The weakness of HeapSort that SmoothSort addresses is that HeapSort
'builds a heap in reverse order.  It then continually pushes elements
'down the heap that have been taken from the bottom, in all likelihood,
'the elements chosen belong at the bottom and need to be pushed the
'entire depth of the heap.  SmoothSort addresses this weakness by
'building the heap in forward order, taking advantage of any pre-sorted
'data and then restructuring the heap at the top.
'
'Just like HeapSort, SmoothSort works in two stages.  The first stage is
'to build a heap.  The second stage is to continually remove the top
'element and restructure the heap.'
'
'
'SOME CONVENTIONS
'----------------
'In these notes, I assume that the sort is to be sorting in ascending order.
'I use the terms tree, heap and span interchangeably.
'I use the terms sub-tree, sub-heap and sub-span interchangeably.
'
'HOW I RENAMED VARIABLES
'-----------------------
'Dijkstra's pseudocode for this algorithm was terse.  To make it more
'readable, I renamed the variables:
'
'Dijkstra's Variable Name | My Variable Name
'----------------------------------------------------
'                       q | lngOneBasedIndex
'                       R | lngNodeIndex
'                       b | lngSubTreeSize
'                       c | lngLeftSubTreeSize
'                       p | lngLeftRightTreeAddress
'              r1, r2, r3 | lngLeftChildIndex, lngChildIndex etc
'
'HEAPS FOR BEGINNERS
'-------------------
'A heap is a binary tree that ensures that the maximum (or minimum)
'element is at the top.  It does this by ensuring that the value of each
'node is greater than the values of its two children.
'
'For Example:
'                             9
'                            / \
'                        ----   ----
'                       /           \
'                      8             4
'                     / \           / \
'                    /   \         /   \
'                   /     \       /     \
'                  2       6     2       3
'
'In this simple heap, it is guaranteed that the largest number is on top.
'This is done because 9 is guaranteed larger than 8 and 4, 8 is guaranteed
'to be larger than 2 and 6 and 4 is guaranteed to be larger than 2 and 3.
'
'A heap does not guarantee the location of the second, third, fourth, last
'or any other element.
'
'
'THE HEAP IN SMOOTHSORT
'----------------------
'The heap in SmoothSort is unbalanced.  The left sub-tree of any node will
'always be one level deeper than the right.  The sizes of fully complete
'trees/sub-trees are Leonardo numbers like 1, 1, 3, 5, 9, 15, 25, 41 and 67.
'
'These numbers are significant because each number is the sum of the
'previous two numbers plus one.  For example, 15 = 5 + 9 + 1.
'
'i.e.  The formula for Leonardo Numbers is:
'               L(n) = L(n-1) + L(n-2) + 1
'
'This is how it works.  Consider the tree:
'
'           9
'          / \
'        --   --
'       /       \
'      5         8
'     / \       / \
'    /   \     /   \
'   3     4   6     7
'  / \
' 1   2
'
'
'Once again, this tree is a heap because each node has a greater value then
'both of its two children.
'
'Also notice that there are 9 nodes in this tree.  Nine is a Leonardo number.
'Also notice that the size of the left sub-tree is 5 nodes and the size of the
'right sub-tree is three.  These are the two Leonardo numbers preceding nine
'and 3 + 5 + 1 = 9.
'
'The consequence of using these Leonardo number is that the left sub-tree has
'always got 1 more level than the right.  SmoothSort can be implemented with a
'balanced tree but the reason for choosing a skewed tree will become apparent
'later.
'
'HOW SMOOTHSORT STORES THE HEAP
'------------------------------
'The Heap is stored in memory in this order:
'- First the left sub-tree;
'- Then the right sub-tree;
'- Then the parent node.
'
'The SmoothSort heap example above would be stored as:
'
' Value   |1|2|3|4|5|6|7|8|9|
'         -------------------
' Address |0|1|2|3|4|5|6|7|8|
'
'In this case the numbers |1|2|3|4|5| is the left sub-tree and is stored first.
'The numbers |6|7|8| is the right sub-tree and is stored immediately afterwards.
'The number |9| is the parent and is stored immediately after that.
'
'If you look at the sub-tree that has 5 as a parent:
'
'The numbers |1|2|3| are the left sub-tree and are stored first.
'The number |4| is the right sub-tree and is stored immediately afterwards.
'The number |5| is the parent and is stored immediately after that.
'
'
'NAVIGATING THE SMOOTHSORT HEAP
'------------------------------
'When navigating the smoothsort heap/tree, we need to keep track of the following
'four numbers:
'- lngNodeIndex             This is the physical index into the array being sorted
'- lngSubTreeSize           The is the number of elements in this sub-tree which
'                           equals the size of the left sub-tree plus the size of
'                           the right sub-tree plus 1 for the node itself.
'- lngLeftSubTreeSize       This is the size of the left sub-tree of the current
'                           node.
'- lngLeftRightTreeAddress  This is complicated and is explained later.
'
'
'
'NAVIGATING DOWN TO THE LEFT CHILD
'---------------------------------
'
'To navigate down the the left child of a node, we need to skip back over the
'right hand child.  To find the sizes of the sub-heaps, we look at the sizes
'of the parent heap, notice that these are Leonardo numbers and that the new
'sizes are the preceding Leonardo numbers.
'
'This is how it works.  Recall our example heap:
'
'           9
'          / \
'        --   --
'       /       \
'      5         8
'     / \       / \
'    /   \     /   \
'   3     4   6     7
'  / \
' 1   2
'
'And how it is stored in memory:
'
' Value   |1|2|3|4|5|6|7|8|9|
'         -------------------
' Address |0|1|2|3|4|5|6|7|8|
'
'If we assume that the current node is at address 4 and has the value 5 above,
'this is what we know:
'
'lngNodeIndex = 4
'lngSubTreeSize = 5
'lngLeftSubTreeSize = 3
'
'At any point in time we can calculate the size of the right sub tree to be:
'   lngSubTreeSize - lngLeftSubTreeSize - 1 = 4
'
'lngNodeIndex(of the Left Child) = lngNodeIndex - size of the right sub tree - 1
'                                = lngNodeIndex - (lngSubTreeSize - lngLeftSubTreeSize - 1) - 1
'                                = lngNodeIndex - lngSubTreeSize + lngLeftSubTreeSize
'                                = 4 - 5 + 3
'                                = 2 (which has value 3)
'
'lngSubTreeSize(of the left Child) = lngLeftSubTreeSize
'                                  = 3
'
'lngLeftSubTreeSize(of the left Child) = size of the right sub tree
'                                      = lngSubTreeSize - lngLeftSubTreeSize - 1
'                                      = 5 - 3 - 1
'                                      = 1
'
'This gives us:
'
'lngNodeIndex = 2 (which has value 3)
'lngSubTreeSize = 3
'lngLeftSubTreeSize = 1
'
'And this correctly describes the left child of index 4 (value 5), at index 2 (value 3).
'
'
'
'NAVIGATING DOWN TO THE RIGHT CHILD
'----------------------------------
'This is how it works.  Recall our example heap:
'
'To navigate down to the right hand child, it is only necessary to step back one index
'in the data array.  To find the sub-tree sizes, we look at the sizes of the parent
'sub tree and notice that these sizes are Leonardo numbers.  We then use the Leonardo
'numbers two back from the parent's sizes as the new sizes.
'
'           9
'          / \
'        --   --
'       /       \
'      5         8
'     / \       / \
'    /   \     /   \
'   3     4   6     7
'  / \
' 1   2
'
'And how it is stored in memory:
'
' Value   |1|2|3|4|5|6|7|8|9|
'         -------------------
' Address |0|1|2|3|4|5|6|7|8|
'
'If we assume that the current node is at address 8 and has the value 9 above,
'this is what we know:
'
'lngNodeIndex = 8
'lngSubTreeSize = 9
'lngLeftSubTreeSize = 5
'
'At any point in time we can calculate the size of the right sub tree to be:
'   lngSubTreeSize - lngLeftSubTreeSize - 1 = 3
'
'lngNodeIndex(of the Right Child) = lngNodeIndex - 1
'                                 = 8 - 1
'                                 = 7 (which has value 8)
'
'lngSubTreeSize(of the Right Child) = size of the right sub tree
'                                   = lngSubTreeSize - lngLeftSubTreeSize - 1
'                                   = 9 - 5 - 1
'                                   = 3
'
'lngLeftSubTreeSize(of the Right Child) = Two Leonardo numbers down from lngLeftSubTreeSize (5)
'                                       = One Leonardo number down from the new lngSubTreeSize(of the Right Child)
'But recall that the formula for Leonardo numbers is:  L(n)   = L(n-1) + L(n-2) + 1
'                                                 So:  L(n-2) = L(n) - L(n-1) - 1
'So
'lngLeftSubTreeSize(of the Right Child) = lngLeftSubTreeSize - new lngSubTreeSize(of the Right Child) - 1
'                                       = 5 - 3 - 1
'                                       = 1
'This gives us:
'
'lngNodeIndex = 7 (which has value 8)
'lngSubTreeSize = 3
'lngLeftSubTreeSize = 1
'
'And this correctly describes the right child of index 8 (value 9), at index 7 value 8.
'
'
'NAVIGATING UP FROM A LEFT CHILD
'-------------------------------
'To navigate up from a left child, we need to step over the right hand sibling.  The
'size of the right hand sibling is the next Leonardo number up from the size of the
'original left hand sibling.
'Also the sub-heap sizes of the parent are also the next Leonardo number up from the
'corresponding sizes of the child.
'
'This is how it works.  Recall our example heap:
'
'           9
'          / \
'        --   --
'       /       \
'      5         8
'     / \       / \
'    /   \     /   \
'   3     4   6     7
'  / \
' 1   2
'
'And how it is stored in memory:
'
' Value   |1|2|3|4|5|6|7|8|9|
'         -------------------
' Address |0|1|2|3|4|5|6|7|8|
'
'If we assume that the current node is at address 4 and has the value 5 above,
'this is what we know:
'
'lngNodeIndex = 4
'lngSubTreeSize = 5
'lngLeftSubTreeSize = 3
'
'lngNodeIndex(of the Parent) = lngNodeIndex + size of the right sub tree + 1
'                            = lngNodeIndex + (next Leonardo size down from size of lngSubTreeSize) + 1
'(coincidentally)            = lngNodeIndex + lngLeftSubTreeSize + 1
'                            = 4 + 3 + 1
'                            = 8 (which has a value 9)
'
'lngSubTreeSize(of the Parent) = lngSubTreeSize + size of the right sub tree + 1
'                              = lngSubTreeSize + (next Leonardo size down from size of lngSubTreeSize) + 1
'(coincidentally)              = lngSubTreeSize + lngLeftSubTreeSize + 1
'                              = 5 + 3 + 1
'                              = 9
'
'lngLeftSubTreeSize(of the left Child) = lngSubTreeSize
'                                      = 5
'
'This gives us:
'
'lngNodeIndex = 8 (which has value 9)
'lngSubTreeSize = 9
'lngLeftSubTreeSize = 5
'
'And this correctly describes the parent of index 4 (value 5), at index 8 (value 9).
'
'
'
'
'NAVIGATING UP FROM A RIGHT CHILD
'--------------------------------
'To navigate up from a right child, we just need to step up by one as the parent is
'immediately after the right child.
'Also the sub-heap sizes of the parent are also two Leonardo numbers up from the
'corresponding sizes of the child.
'
'This is how it works.  Recall our example heap:
'
'           9
'          / \
'        --   --
'       /       \
'      5         8
'     / \       / \
'    /   \     /   \
'   3     4   6     7
'  / \
' 1   2
'
'And how it is stored in memory:
'
' Value   |1|2|3|4|5|6|7|8|9|
'         -------------------
' Address |0|1|2|3|4|5|6|7|8|
'
'If we assume that the current node is at address 7 and has the value 8 above,
'this is what we know:
'
'lngNodeIndex = 7
'lngSubTreeSize = 3
'lngLeftSubTreeSize = 1
'
'lngNodeIndex(of the Parent) = lngNodeIndex + 1
'                            = 8 (which has a value 9)
'
'lngSubTreeSize(of the Parent) = two Leonardo Numbers up from lngSubTreeSize
'                              = one Leonardo Number up from lngSubTreeSize + lngLeftSubTreeSize + 1
'                              = lngSubTreeSize + (lngSubTreeSize + lngLeftSubTreeSize + 1) + 1
'                              = 3 + 3 + 1 + 1 + 1
'                              = 9
'
'lngLeftSubTreeSize(of the Parent) = two Leonardo Numbers up from lngLeftSubTreeSize
'                                  = one Leonardo Number up from nlgSubTreeSize
'                                  = lngSubTreeSize + lngLeftSubTreeSize + 1
'                                  = 3 + 1 + 1
'                                  = 5
'
'This gives us:
'
'lngNodeIndex = 8 (which has value 9)
'lngSubTreeSize = 9
'lngLeftSubTreeSize = 5
'
'And this correctly describes the parent of index 7 (value 8), at index 8 (value 9).
'
'
'
'A SECOND TREE
'-------------
'
'Navigating the Heap using the methods above works pretty.  You can go up and down
'as needed and when you hit the bottom, the size of the left sub-tree will be < 1.
'
'There is a problem though:  It is that you need to know if the node is a left node
'or a right node in order to choose the correct algorithm for going up to its parent.
'
'To solve this problem Dijkstra introduced another virtual tree.  This virtual
'tree does not occupy any memory and consists of a single address that identifies
'the current position in the tree.  When SmoothSort navigates the heap, it
'simultaneously navigates the virtual tree so that the function points to the same
'relative place in both trees.
'
'The structure of this tree is defined as:
'- The address of the top node is 1;
'- The address of the left hand child p(left) = p(parent) * 2 - 1
'- The address of the right hand child p(right) = p(parent) * 4 - 1
'
'This produces the following tree:
'
'                                       1
'                                      / \
'                      ----------------   ----------------
'                     /                                   \
'                    1                                     3
'                   / \                                   / \
'            -------   -------                     -------   -------
'           /                 \                   /                 \
'          1                   3                 5                   11
'         / \                 / \               / \                 / \
'       --   --             --   --           --   --             --   --
'      /       \           /       \         /       \           /       \
'     1         3         5         11      9         19        21        43
'    / \       / \       / \       / \     / \       / \       / \       / \
'   1   3     5   11    9   19    21  43  17  35    37  75    41  83    85  171
'
'
'All of the addresses in this tree satisfy one of these two conditions:
'- p Mod 8 = 3 (if it is a right hand child node)
'- p Mod 4 = 1 (if it is a left hand child node)
'
'We can prove this inductively as:
'
'Assume p is a right hand node, it can be expressed as p = (8k+3) (k is an integer)
'The left hand child of p will be = (8k + 3) * 2 - 1
'                                 = 16k + 5
'Mod this by 8 and the answer is 5, therefore it is not a right hand child.
'Mod this by 4 and the answer is 1, therefore is is a left hand child.
'
'The right hand child of p will be = (8k + 3) * 4 - 1
'                                  = 32k + 11
'Mod this by 8 and the answer is 3, therefore it is a right hand child.
'Mod this by 4 and the answer is 3, therefore it is not a left hand child.
'
'If we then assume that it is a left hand node, it can be expressed as p = (4k+1).
'The left hand child of p will be = (4k+1) * 2 - 1
'                                 = 8k + 1
'Mod this by 8 and the answer = 1, therefore it is not a right hand child.
'Mod this by 4 and the answer = 1, therefore it is a left hand child.
'
'The right hand child of p will be = (4k+1) * 4 - 1
'                                  = 16k + 3
'
'Mod this by 8 and the anser is 3 therefore it is a right hand child.
'Mod this by 4 and the anser is 3 therefore it is not a left hand child.
'
'As the start of the tree is at p=1 which satisfies p Mod 4 = 1, then all
'values in the tree will satisfy the above rules and left hand nodes can
'be discerned from right hand nodes with certainty.
'
'Navigating up the tree from a left hand child is achieved by:
'   p(Parent) = (p(Left Child) + 1) / 2
'
'Navigating up the tree from a right hand child is achieved by:
'   p(Parent) = (p(Right Child) + 1) / 4
'
'Additionally, it can be seen that navigating from the left child
'to the right is done by:
'   p(Right) = p(Left) * 2 + 1
'
'NOT A PERFECT WORLD
'-------------------
'
'So far my examples have used data sets that were conveniently
'exactly the size of Leonardo numbers.  This won't often happen
'and will not be the case immediately once you start sorting.
'
'
'Consider this example (which is a whole Leonardo number):
'
'                     15 (Top of the heap)
'                    / \
'             -------   -------
'            /                 \
'           9                   14
'          / \                 / \
'        --   --              /   \
'       /       \            /     \
'      5         8          12      13
'     / \       / \        / \
'    /   \     /   \      /   \
'   3     4   6     7    10    11
'  / \
' 1   2
'
'
'If we had one less element of data, there would be two trees:
'
'           9                   14
'          / \                 / \
'        --   --              /   \
'       /       \            /     \
'      5         8          12      13
'     / \       / \        / \
'    /   \     /   \      /   \
'   3     4   6     7    10    11
'  / \
' 1   2
'
'Smoothsort, recognises that the parent of the #14 element above is
'beyond the limit of the data and it then gives the #14 element a
'third child like:
'
'            _________
'           9         \_________14  (Top of the Heap)
'          / \                 / \
'        --   --              /   \
'       /       \            /     \
'      5         8          12      13
'     / \       / \        / \
'    /   \     /   \      /   \
'   3     4   6     7    10    11
'  / \
' 1   2
'
'SmoothSort know how to do this because it knows that the size of the
'sub-heap topped by #14 (at index 13) is 5.  Therefore the top of the
'previous untopped heap is at 13 - 5 = index 8 (#9).
'
'If there was one less element of data again, there would be 3 trees:
'
'           9
'          / \
'        --   --
'       /       \
'      5         8          12      13
'     / \       / \        / \
'    /   \     /   \      /   \
'   3     4   6     7    10    11
'  / \
' 1   2
'
'Once again SmoothSort recognises that the parent of #13 is beyond the
'range of the data and gives it a third child at #12.  Similarly, it
'recognises that the parent of #12 is beyond the data range and gives
'it a third child #9.
'            _____
'           9     \
'          / \     \
'        --   --    \
'       /       \    \        ___
'      5         8    \_____12   \__13  (Top of the Heap)
'     / \       / \        / \
'    /   \     /   \      /   \
'   3     4   6     7    10    11
'  / \
' 1   2
'
'
'Removing one more element from heap produces:
'            _____
'           9     \
'          / \     \
'        --   --    \
'       /       \    \
'      5         8    \_____12  (Top of the Heap)
'     / \       / \        / \
'    /   \     /   \      /   \
'   3     4   6     7    10    11
'  / \
' 1   2
'
'And again:
'            ____
'           9    \
'          / \    \
'        --   --   \
'       /       \   \
'      5         8   \
'     / \       / \   \
'    /   \     /   \   \    ___
'   3     4   6     7   \_10   \___11  (Top of the Heap)
'  / \
' 1   2
'
'And again:
'            ____
'           9    \
'          / \    \
'        --   --   \
'       /       \   \
'      5         8   \
'     / \       / \   \
'    /   \     /   \   \
'   3     4   6     7   \_10  (Top of the Heap)
'  / \
' 1   2
'
'And again:
'
'           9   (Top of the Heap)
'          / \
'        --   --
'       /       \
'      5         8
'     / \       / \
'    /   \     /   \
'   3     4   6     7
'  / \
' 1   2
'
'
'There are two things that need to be noted and can be observed in this
'example:
'
'1.  Because of the way that the heaps are stored in memory, the
'    maximum number of nodes without their normal parents is log2(n).
'2.  Because the tree is skewed using Leonardo numbers, the number
'    of these orphan nodes is < log2(n).
'
'ENOUGH OF THE PREAMPLE, GET ON WITH THE CODE


Public Sub SmoothSort(plngArray() As DataElement)

    'SmoothSort is the main function that performs the smoothsort.
    'It has two main phases, the first is to build a heap.
    'The second phase is to remove the top element of the heap and
    'rebuild the heap.  This second phase is repeated until there
    'is nothing left in the heap and all of the data is sorted.
    
    Dim lngOneBasedIndex As Long
    Dim lngNodeIndex As Long
    Dim lngLeftRightTreeAddress As Long
    Dim lngSubTreeSize As Long
    Dim lngLeftSubTreeSize As Long
    
    'Initialise the variables.
    lngLeftRightTreeAddress = 1
    lngSubTreeSize = 1
    lngLeftSubTreeSize = 1
    lngOneBasedIndex = 1
    lngNodeIndex = 0
    
    'The first phase is to build the heap.  Loop through the data
    'one element at a time.
    'Each element is a node in the heap.
    Do While lngOneBasedIndex <> UBound(plngArray) + 1
    
        'This element is at the top of a sub-heap (which may just be itself only).
        
        'If the current node is the right child of its parent
        If lngLeftRightTreeAddress Mod 8 = 3 Then
        
            'Push this element down the sub-heap that it sits on - just like heap sort
            SmoothSift plngArray, lngNodeIndex, lngSubTreeSize, lngLeftSubTreeSize
            
            
            'The next element to be processed will be the parent of this sub-heap.
            
            '1. Move up a right leg of the virtual left/right tree
            lngLeftRightTreeAddress = (lngLeftRightTreeAddress + 1) \ 4
            
            '2. The SubTreeSizes of the parent of a right child is two steps back
            'down the sequence of Leonardo numbers, Move up the sequence of leonardo
            'numbers twice.
            SmoothUp lngSubTreeSize, lngLeftSubTreeSize
            SmoothUp lngSubTreeSize, lngLeftSubTreeSize
            
            'Dont worry about the parent of this element being off the scale.  It is
            'always the next element and we do a Trinkle on the last element anyway.
            
        ElseIf lngLeftRightTreeAddress Mod 4 = 1 Then 'This is always true if it gets here
        
            'If the current node is the left child of its parent
            
            'The parent of this node will be a distance away equal to the size of its
            'own left child.  See if this is within the bounds of the data.
            If lngOneBasedIndex + lngLeftSubTreeSize < UBound(plngArray) + 1 Then
                'If the parent of this node is within the data, just push this value down
                'its own sub heap.  Just like heap sort.
                SmoothSift plngArray, lngNodeIndex, lngSubTreeSize, lngLeftSubTreeSize
            Else
                'The parent of this node is beyond the end of the data.
                'Give this node a third child which is the last element prior to the start
                'of this sub-heap.  This element will be the sibling of this node or else
                'it will be the sibling of one of the parents of this node.
                SmoothTrinkle plngArray, lngNodeIndex, lngLeftRightTreeAddress, lngSubTreeSize, lngLeftSubTreeSize
            End If
            
            'The next element to be processed will be the far left-left-left child of the right sibling
            'of this node
            
            'To get to the right sibling of this node, the formula is p * 2 + 1.  (See the info above on
            'the virtual tree.
            'Then each time we go down the left leg we then apply the p * 2 - 1 formula.
            
            'As it turns out, the formula to go p * 2 + 1 and then to repeat p * 2 - 1 n times
            'is:     p^(n-1)+1
            '
            'Pretty tricky but it works
            Do
                SmoothDown lngSubTreeSize, lngLeftSubTreeSize
                lngLeftRightTreeAddress = lngLeftRightTreeAddress * 2
            Loop While lngSubTreeSize <> 1 'Continue until we reach the bottom of the tree
            lngLeftRightTreeAddress = lngLeftRightTreeAddress + 1
            
        End If
        lngOneBasedIndex = lngOneBasedIndex + 1
        lngNodeIndex = lngNodeIndex + 1
    Loop
    
    'SmoothTrinkle on the last element, will give the last element 3 children.  This will be needed
    'if the last element was a right child
    SmoothTrinkle plngArray, lngNodeIndex, lngLeftRightTreeAddress, lngSubTreeSize, lngLeftSubTreeSize
    
    
    'This loop is about reducing the size of the heap by 1 and reshuffling the
    'heap until the whole lot is sorted.  Just like heap sort except that the top
    'of the heap is where the heap gets smaller, meaning that we do not have to do
    'things in reverse and we do not have to pick an element from the bottom of the
    'heap and push it down from the top.
    Do While lngOneBasedIndex <> 1
        lngOneBasedIndex = lngOneBasedIndex - 1
        
        If lngSubTreeSize = 1 Then
            'This sub-tree only had one element
            
            'Prepare to look at the previous element in the next loop
            lngNodeIndex = lngNodeIndex - 1
            
            'Navigate in both trees to the next element.
            'First you navigate up, as long as it is a left leg (which it may not be,
            'it may already be a right leg, in which case you stay there).
            'Then you go across to the left leg.
            'Once again, another pretty tricky piece of maths but it works.
            lngLeftRightTreeAddress = lngLeftRightTreeAddress - 1
            Do While lngLeftRightTreeAddress Mod 2 = 0
                lngLeftRightTreeAddress = lngLeftRightTreeAddress / 2
                SmoothUp lngSubTreeSize, lngLeftSubTreeSize
            Loop
            
        ElseIf lngSubTreeSize >= 3 Then 'It must fall in here, sub trees are either size 1,1,3,5,9,15 etc
        
            'This makes the lngLeftRightTreeAddress even and will cause the Trinkle
            'function to navigate up to the right level.
            lngLeftRightTreeAddress = lngLeftRightTreeAddress - 1
            
            'This node has children, get the index of the left child
            lngNodeIndex = lngNodeIndex + lngLeftSubTreeSize - lngSubTreeSize
            'The right child, being immediately behind this node will now be the new top
            
            'If the node to be removed is the top top node then there are no nodes to the left of it.
            'We do not need to call SmoothSemiTrinkle to join the left child up with previous heaps.
            'If it is not the top top node then we need to call SmoothSemiTrinkle on the left child
            'to link it up with the heaps to the left.
            If lngLeftRightTreeAddress <> 0 Then
                SmoothSemiTrinkle plngArray, lngNodeIndex, lngLeftRightTreeAddress, lngSubTreeSize, lngLeftSubTreeSize
            End If
            
            'Navigate across from the left child to the right child.
            SmoothDown lngSubTreeSize, lngLeftSubTreeSize
            
            'Get the lngLeftRightTreeAddress of the left child
            lngLeftRightTreeAddress = lngLeftRightTreeAddress * 2 + 1
            
            'Finish navigating across from the left child to the right child
            lngNodeIndex = lngNodeIndex + lngLeftSubTreeSize
            
            'Call semi-smooth trinkle to make sure that it has three legs and that it links up with
            'the left child and all other previous heaps.
            SmoothSemiTrinkle plngArray, lngNodeIndex, lngLeftRightTreeAddress, lngSubTreeSize, lngLeftSubTreeSize
            
            
            SmoothDown lngSubTreeSize, lngLeftSubTreeSize
            lngLeftRightTreeAddress = lngLeftRightTreeAddress * 2 + 1
            
        End If
    Loop
End Sub

Private Sub SmoothUp(lngSubTreeSize As Long, lngLeftSubTreeSize As Long)

    'This function, passed two sequential Leonardo numbers like 15, 9
    'will step up the sequence of Leonardo numbers and return the next one.
    '
    'For example, if passed 15,9 - this function will return 25,15
    '             if passed 5,3  - this function will return 9, 5
    '
    'If called once, it will calculate the size of a parent heap for a left child.
    'If called once, it will calculate the size of the left sibling of a right child.
    'If called twice, it will calculate the size of a parent heap for a right child.

    Dim temp As Long
    
    temp = lngSubTreeSize + lngLeftSubTreeSize + 1
    lngLeftSubTreeSize = lngSubTreeSize
    lngSubTreeSize = temp
    
End Sub


Private Sub SmoothDown(lngSubTreeSize As Long, lngLeftSubTreeSize As Long)
    
    'This function, passed two sequential Leonardo numbers like 15, 9
    'will step down the sequence of Leonardo numbers and return the previous one.
    '
    'For example, if passed 15,9 - this function will return 9, 5
    '             if passed 5,3  - this function will return 3, 1
    '
    'If called once, it will calculate the size of a left child sub-heap.
    'If called once, it will calculate the size of the right sibling of a left child.
    'If called twice, it will calculate the size of a right child sub-heap.
    
    Dim temp As Long
    
    temp = lngSubTreeSize - lngLeftSubTreeSize - 1
    lngSubTreeSize = lngLeftSubTreeSize
    lngLeftSubTreeSize = temp
    
End Sub

Private Sub SmoothSift(plngArray() As DataElement, ByVal lngNodeIndex As Long, ByVal lngSubTreeSize As Long, ByVal lngLeftSubTreeSize As Long)
    
    'This function pushes the element on top of a binary heap down
    'until it reaches the correct place in the heap.
    'Just like heap sort.
    
    Dim lngChildIndex As Long
    
    'Do while the current tree has children
    Do While lngSubTreeSize >= 3
    
        'Get the index of the left child
        lngChildIndex = lngNodeIndex - lngSubTreeSize + lngLeftSubTreeSize
        
        'Compare the value of the left child with the right child to find
        'the child with the maximum value.
        If plngArray(lngChildIndex).theKey < plngArray(lngNodeIndex - 1).theKey Then
        
            'The right child has the greater value, this is the value that
            'will rise to the top of the heap if need be.
            lngChildIndex = lngNodeIndex - 1
            
            'Because we are going down the right child, we need to do an
            'extra SmoothDown operation because right children are two
            'steps down the Leonardo sequence.
            SmoothDown lngSubTreeSize, lngLeftSubTreeSize
            
            'We dont need to worry about the virtual left/right tree
            'because we wont be going back up the tree.
        End If
        
        'Compare the greater child with the parent
        If plngArray(lngNodeIndex).theKey >= plngArray(lngChildIndex).theKey Then
            'The parent was bigger, the job is done, no more to do.
            lngSubTreeSize = 1
        Else
        
            'The child is greater than the parent, swap them around
            Exchange plngArray, lngNodeIndex, lngChildIndex
            
            'Move down to the next level of the heap
            lngNodeIndex = lngChildIndex
            'Going down either leg only requires one step because an
            'extra step for the right has already been done.
            SmoothDown lngSubTreeSize, lngLeftSubTreeSize
        End If
    Loop
End Sub

Private Sub SmoothTrinkle(plngArray() As DataElement, ByVal lngNodeIndex As Long, ByVal lngLeftRightTreeAddress As Long, ByVal lngSubTreeSize As Long, ByVal lngLeftSubTreeSize As Long)

    'This function pushes the current node down into the heap
    'until it reaches the correct place.
    'It differs from SmoothSift though as the node passed
    'to this function is given three children:
    ' - The two normal children that all nodes have; plus
    ' - A third child which is the top of the previous complete sub-heap
    '
    'It assumes that the node is already the top of a properly constructed
    'heap with the normal 2 children.

    Dim lngChildIndex As Long
    Dim lngPreviousCompleteTreeIndex As Long
    
    'Consider the complete virtual tree:
    '
    '                                       1
    '                                      / \
    '                      ----------------   ----------------
    '                     /                                   \
    '                    1                                     3
    '                   / \                                   / \
    '            -------   -------                     -------   -------
    '           /                 \                   /                 \
    '          1                   3                 5                   11
    '         / \                 / \               / \                 / \
    '       --   --             --   --           --   --             --   --
    '      /       \           /       \         /       \           /       \
    '     1         3         5         11      9         19        21        43
    '    / \       / \       / \       / \     / \       / \       / \       / \
    '   1   3     5   11    9   19    21  43  17  35    37  75    41  83    85  171
    '
    '
    'If the number of elements in the heap is smaller than this by let say 3,
    'the virtual tree will look like this:
    '
    '
    '                    1
    '                   / \
    '            -------   -------
    '           /                 \
    '          1                   3                 5
    '         / \                 / \               / \
    '       --   --             --   --           --   --
    '      /       \           /       \         /       \
    '     1         3         5         11      9         19        21        43
    '    / \       / \       / \       / \     / \       / \       / \       / \
    '   1   3     5   11    9   19    21  43  17  35    37  75    41  83    85  171
    '
    '
    'This function may have been passed the details of the node at address #43 above.
    'It will join these heaps into a single heap and keep on processing until it reaches
    'the sub-heap with the address #1 at the top.
    
    'Keep looping until the virtual tree address is > 0 (ie all of the above heaps)
    Do While lngLeftRightTreeAddress > 0
    
        'Here is yet another tricky thing.  This bit is how the function navigates
        'from the top of one of the above complete heaps to the top of the previous.
        '
        'Firstly note that all of the addresses in the virtual tree are odd.
        '(Not surprisingly as the formulas for getting them are *2-1 and *4-1).
        '
        'Now consider the navigation above:
        '
        'Starting at address #43, Subtract 1 and divide by 2 gives 21.
        '#21 Is odd so it is a good address and since it is the left sibling of #43
        'we need to go up 1.
        '
        'Starting at address #21, subtract 1 and divide by 2 gives 10.
        '#10 is even and so it is not a good address, divide by 2 again to get #5.
        '#5 is odd so it is a good address and correct in the above example.
        'The SmoothUp function needs to be called once to get to the parent of #21,
        'it needs to be called twice to get to its parent (also the parent of #5).
        'This means that SmoothUp needs to be called three times and then SmoothDown
        'needs to be called once to navigate to #5.
        'The net is that SmoothUp needs to be called twice - the same number of times
        'that we need to divide by zero.
        '
        'Starting at address #5, we need to subtract 1 and then divide by 2 twice to
        'get address #1.  Also SmoothUp needs to be called twice.
        '
        'This loop achieves this navigation.  When SmoothTrinkle is called, the
        'lngLeftRightTreeAddress is valid (odd) and so this loop does not execute.
        'Each time the outer loop executes, it subtracts 1 from the lngLeftRigthTreeAddress
        '(later) and so this loop correctly does the / 2 and the SmoothUps.
        Do While lngLeftRightTreeAddress Mod 2 = 0
            lngLeftRightTreeAddress = lngLeftRightTreeAddress \ 2
            SmoothUp lngSubTreeSize, lngLeftSubTreeSize
        Loop
        
        'Get the index of the last full tree prior to this sub tree
        lngPreviousCompleteTreeIndex = lngNodeIndex - lngSubTreeSize
        'In the above example, for the node #43, get the address of #21 (ie its left sibling)
        'In the above example, for the node #21, get the address of #5 (ie left sibling of parent)
        
        'If this node is the top of the furthest left complete tree then stop processing
        If lngLeftRightTreeAddress = 1 Then
        
            'We are in this situation (The numbers here are values)
            '            _____
            '           9     \   <-lngNodeIndex points here
            '          / \     \
            '        --   --    \
            '       /       \    \
            '      5         8    \_____>9
            '     / \       / \
            '    /   \     /   \
            '   3     4   6     7
            '  / \
            ' 1   2
            '
            'There is nothing on the left of this heap.
        
            'Job is done, stop processing
            lngLeftRightTreeAddress = 0
            
        'Else if the value at the top of the previous complete heap is less than this node
        ElseIf plngArray(lngPreviousCompleteTreeIndex).theKey <= plngArray(lngNodeIndex).theKey Then
            
            'We are in this situation (The numbers here are values)
            '            _____
            '           9     \   <-lngPreviousCompleteTreeIndex points here
            '          / \     \
            '        --   --    \
            '       /       \    \
            '      5         8    \_____10  <-lngNodeIndex points here
            '     / \       / \        / \
            '    /   \     /   \      /   \
            '   3     4   6     7    7     6
            '  / \
            ' 1   2
            '
            '
            'Clearly the number 10 is larger than 9, the heap is good.
            
            'Do not push this value further down the tree, job done.
            lngLeftRightTreeAddress = 0
            
        Else
            
            'Make this even so that it gets calculated correctly on the next loop
            lngLeftRightTreeAddress = lngLeftRightTreeAddress - 1
            
            'If this complete heap has only one element
            If lngSubTreeSize = 1 Then
            
                'We are in this situation (The numbers here are values)
                '            _____
                '           9     \   <-lngPreviousCompleteTreeIndex points here
                '          / \     \
                '        --   --    \
                '       /       \    \
                '      5         8    \_____7  <-lngNodeIndex points here
                '     / \       / \
                '    /   \     /   \
                '   3     4   6     7
                '  / \
                ' 1   2
                '
                '
                'Clearly the number 9 is larger than 7
                
                'Just do a simple swap and move on to the next complete heap
                Exchange plngArray, lngNodeIndex, lngPreviousCompleteTreeIndex
                lngNodeIndex = lngPreviousCompleteTreeIndex
                
                'After doing this swap, the next complete heap will not
                'necessarily be valid, (ie in this example 7 would be on top of an 8).
                'If this is the last node processed in this function, the
                'SmoothSift call at the end will do this.
                'If this is not the last node processed, the code below compares
                'the previous top, with the children and deals with this.
                
            'Else if this complete heap has normal (2) children
            ElseIf lngSubTreeSize >= 3 Then
            
                'We are in one of these two situations:(The numbers here are values)
                'In both situations, the value at the top of the previuos complete heap
                'is greater than the top of this heap.
                'Situation A, has a greater child in this heap.
                'Situation B, has the parent greater in this heap.
                '
                'Situation A
                '            _____
                '           9     \   <-lngPreviousCompleteTreeIndex points here
                '          / \     \
                '        --   --    \
                '       /       \    \
                '      5         8    \_____7  <-lngNodeIndex points here
                '     / \       / \        / \
                '    /   \     /   \      /   \
                '   3     4   6     7    10    9
                '  / \
                ' 1   2
                '
                '
                'Situation B
                '            _____
                '           9     \   <-lngPreviousCompleteTreeIndex points here
                '          / \     \
                '        --   --    \
                '       /       \    \
                '      5         8    \_____7  <-lngNodeIndex points here
                '     / \       / \        / \
                '    /   \     /   \      /   \
                '   3     4   6     7    6     5
                '  / \
                ' 1   2
                '
                'Clearly the number 9 is larger than 7
                
                'Identify the maximum child of this heap
                'First get the top of the left child
                lngChildIndex = lngNodeIndex - lngSubTreeSize + lngLeftSubTreeSize
                'See whether the left or right child is greater.
                If plngArray(lngChildIndex).theKey < plngArray(lngNodeIndex - 1).theKey Then
                    'The right child is greater
                    'Use the right child
                    lngChildIndex = lngNodeIndex - 1
                    
                    'As we are using the right child, do an extra
                    'SmoothDown and an extra * 2
                    SmoothDown lngSubTreeSize, lngLeftSubTreeSize
                    lngLeftRightTreeAddress = lngLeftRightTreeAddress * 2
                End If
                
                'Now compare the value at the top of the previous complete tree
                'with the maximum child value.
                If plngArray(lngPreviousCompleteTreeIndex).theKey >= plngArray(lngChildIndex).theKey Then
                
                    'The top of the previous complete heap is greater than the top of this
                    'heap and greater than both of its children.
                    'Swap it into place
                    Exchange plngArray, lngNodeIndex, lngPreviousCompleteTreeIndex
                    
                    'Move on ready to do the next complete heap in the SmoothTrinkle
                    lngNodeIndex = lngPreviousCompleteTreeIndex
                Else
                
                    'The child is greater than the the top and greater than
                    'the top of the previous complete heap.
                    'Swap this up
                    Exchange plngArray, lngNodeIndex, lngChildIndex
                    
                    'We are going to stop SmoothTrinke, but now navigate to the
                    'child so that a final SmoothSift can make sure that the
                    'heap is valid.
                    lngNodeIndex = lngChildIndex
                    SmoothDown lngSubTreeSize, lngLeftSubTreeSize
                    
                    'Stop the SmoothTrinkle process.
                    lngLeftRightTreeAddress = 0
                End If
            End If
        End If
    Loop
    
    'Make sure that the top of the final heap is pushed down correctly.
    SmoothSift plngArray, lngNodeIndex, lngSubTreeSize, lngLeftSubTreeSize
    
End Sub

Private Sub SmoothSemiTrinkle(plngArray() As DataElement, ByVal lngNodeIndex As Long, ByVal lngLeftRightTreeAddress As Long, ByVal lngSubTreeSize As Long, ByVal lngLeftSubTreeSize As Long)
    
    'Function to call SmoothTrinkle but only if needed and from the context
    'of removing items from the heap.
    
    Dim lngIndexTopPreviousCompleteHeap As Long
    
    'Parameters to this function are:
    '  lngNodeIndex -            The index of the right child of a node being removed from the heap
    '  lngLeftRightTreeAddress - Tree address of the left node of a node being removed from the heap
    '  lngSubTreeSize -          SubTree size of the left sub-heap of a node being removed
    '  lngLeftSubTreeSize -      Left subtree size of the left sub-heap of a node being removed
    '                            As a Leonardo tree, the same size as the heap headed by lngNodeIndex.
    '
    'OR
    '  lngNodeIndex -            The index of the left child of a node being removed from the heap
    '  lngLeftRightTreeAddress - A number for example 20 that when divided by 2 a number of times will give
    '                            the virtual tree address of the previous complete heap prior to the heap where the top
    '                            is being removed.  The address is found when the number becomes odd.
    '
    'The values in lngSubTreeSize need to be SmoothUp'd the same number of times to give the corresponding values for
    'the complete heap prior to the heap where the top is being removed.
    '
    '  lngSubTreeSize -          SubTree size of the previous complete heap prior to the heap where the top is being removed
    '  lngLeftSubTreeSize -      Left subtree size of the previous complete heap prior to the heap where the top is being removed
    '                            As a Leonardo tree, the same size as the heap headed by lngNodeIndex before being SmoothUp'd.
    
    
    'Get the index of the previous complete heap.
    lngIndexTopPreviousCompleteHeap = lngNodeIndex - lngLeftSubTreeSize
    
    'If the top of the previous complete heap is larger then this one then swap it and Trinkle It.
    If plngArray(lngIndexTopPreviousCompleteHeap).theKey > plngArray(lngNodeIndex).theKey Then
        Exchange plngArray, lngNodeIndex, lngIndexTopPreviousCompleteHeap
        SmoothTrinkle plngArray, lngIndexTopPreviousCompleteHeap, lngLeftRightTreeAddress, lngSubTreeSize, lngLeftSubTreeSize
    End If
    
End Sub


Private Sub Exchange(mlngArray() As DataElement, plng1 As Long, plng2 As Long)

    Dim lngSwap As Long
    
    lngSwap = mlngArray(plng1).theKey
    mlngArray(plng1).theKey = mlngArray(plng2).theKey
    mlngArray(plng2).theKey = lngSwap
    
    lngSwap = mlngArray(plng1).originalOrder
    mlngArray(plng1).originalOrder = mlngArray(plng2).originalOrder
    mlngArray(plng2).originalOrder = lngSwap
    
    
End Sub


