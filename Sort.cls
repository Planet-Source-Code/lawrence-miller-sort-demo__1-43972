VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Swap(ByRef Temp1 As Integer, ByRef Temp2 As Integer)
  Dim Temp As Variant
  
    Temp = Temp1
    Temp1 = Temp2
    Temp2 = Temp
End Sub

Public Sub Bubble(ByRef SortArray() As Integer)
' ============================== BubbleSort ==================================
'    The BubbleSort algorithm cycles through SortArray, comparing adjacent
'    elements and swapping pairs that are out of order.  It continues to
'    do this until no pairs are swapped.
' ============================================================================
  Dim Row    As Integer
  Dim Limit  As Integer
  Dim MaxRow As Integer
  Dim Switch As Integer
  
    MaxRow = UBound(SortArray)
    Limit = MaxRow
    Do
      Switch = False
      For Row = 1 To Limit - 1
        If SortArray(Row) > SortArray(Row + 1) Then
          Swap SortArray(Row), SortArray(Row + 1)
          Switch = Row
        End If
      Next
      Limit = Switch
    Loop While Switch
End Sub

Public Sub Exchange(ByRef SortArray() As Integer)
' ============================= ExchangeSort =================================
'   The ExchangeSort compares each element in SortArray - starting with
'   the first element - with every following element.  If any of the
'   following elements is smaller than the current element, it is exchanged
'   with the current element and the process is repeated for the next
'   element in SortArray.
' ============================================================================
  Dim j      As Integer
  Dim Row    As Integer
  Dim MaxRow As Integer
  Dim MinRow As Integer
  
    MaxRow = UBound(SortArray)
    
    For Row = 1 To MaxRow
      MinRow = Row
      For j = Row + 1 To MaxRow
         If SortArray(j) < SortArray(MinRow) Then
            MinRow = j
         End If
      Next

      ' Found a row shorter than the current row, so swap those
      ' two array elements:
      If MinRow > Row Then
        Swap SortArray(Row), SortArray(MinRow)
      End If
    Next
End Sub

Public Sub Heap(ByRef SortArray() As Integer)
' =============================== HeapSort ===================================
'  The HeapSort procedure works by calling two other procedures - PercolateUp
'  and PercolateDown.  PercolateUp turns SortArray into a "heap," which has
'  the properties outlined in the diagram below:
'
'                               SortArray(1)
'                               /          \
'                    SortArray(2)           SortArray(3)
'                   /          \            /          \
'         SortArray(4)   SortArray(5)   SortArray(6)  SortArray(7)
'          /      \       /       \       /      \      /      \
'        ...      ...   ...       ...   ...      ...  ...      ...
'
'
'  where each "parent node" is greater than each of its "child nodes"; for
'  example, SortArray(1) is greater than SortArray(2) or SortArray(3),
'  SortArray(3) is greater than SortArray(6) or SortArray(7), and so forth.
'
'  Therefore, once the first FOR...NEXT loop in HeapSort is finished, the
'  largest element is in SortArray(1).
'
'  The second FOR...NEXT loop in HeapSort swaps the element in SortArray(1)
'  with the element in MaxRow, rebuilds the heap (with PercolateDown) for
'  MaxRow - 1, then swaps the element in SortArray(1) with the element in
'  MaxRow - 1, rebuilds the heap for MaxRow - 2, and continues in this way
'  until the array is sorted.
' ============================================================================
  Dim i      As Integer
  Dim MaxRow As Integer
  
    MaxRow = UBound(SortArray)
    For i = 2 To MaxRow
      PercolateUp SortArray, i
    Next
    
    For i = MaxRow To 2 Step -1
      Swap SortArray(1), SortArray(i)
      PercolateDown SortArray, i - 1
    Next
End Sub

Public Sub Insertion(ByRef SortArray() As Integer)
' ============================= InsertionSort ================================
'   The InsertionSort procedure compares the length of each successive
'   element in SortArray with the lengths of all the preceding elements.
'   When the procedure finds the appropriate place for the new element, it
'   inserts the element in its new place, and moves all the other elements
'   down one place.
' ============================================================================
  Dim j       As Integer
  Dim Row     As Integer
  Dim MaxRow  As Integer
  Dim TempVal As Integer
   
   MaxRow = UBound(SortArray)
   For Row = 2 To MaxRow
      TempVal = SortArray(Row)
      'TempLength = TempVal.Length
      For j = Row To 2 Step -1

         ' As long as the length of the J-1st element is greater than the
         ' length of the original element in SortArray(Row), keep shifting
         ' the array elements down:
        If SortArray(j - 1) > TempVal Then
          SortArray(j) = SortArray(j - 1)
         ' Otherwise, exit the FOR...NEXT loop:
        Else
          Exit For
        End If
      Next

      ' Insert the original value of SortArray(Row) in SortArray(J):
      SortArray(j) = TempVal
   Next 'Row
End Sub

Public Sub Quick(ByRef SortArray() As Integer, ByVal Low As Integer, ByVal High As Integer)
' ============================== QuickSort ===================================
'   QuickSort works by picking a random "pivot" element in SortArray, then
'   moving every element that is bigger to one side of the pivot, and every
'   element that is smaller to the other side.  QuickSort is then called
'   recursively with the two subdivisions created by the pivot.  Once the
'   number of elements in a subdivision reaches two, the recursive calls end
'   and the array is sorted.
' ============================================================================
  Dim i         As Integer
  Dim j         As Integer
  Dim RandIndex As Integer
  Dim Partition As Integer
  
    If Low < High Then
      ' Only two elements in this subdivision; swap them if they are out of
      ' order, then end recursive calls:
      If High - Low = 1 Then
        If SortArray(Low) > SortArray(High) Then Swap SortArray(Low), SortArray(High)
      Else
        ' Pick a pivot element at random, then move it to the end:
        RandIndex = RandInt%(Low, High)
        If RandIndex <> High Then Swap SortArray(High), SortArray(RandIndex)
        Partition = SortArray(High)
        i = Low: j = High
        Do
          ' Move in from both sides towards the pivot element:
          'i = Low: j = High
          Do While (i < j) And (SortArray(i) <= Partition)
            i = i + 1
          Loop
          Do While (j > i) And (SortArray(j) >= Partition)
            j = j - 1
          Loop
          ' If we haven't reached the pivot element, it means that two
          ' elements on either side are out of order, so swap them:
          If i < j Then Swap SortArray(i), SortArray(j)
        Loop While i < j

        ' Move the pivot element back to its proper place in the array:
        If i <> High Then Swap SortArray(i), SortArray(High)

        ' Recursively call the QuickSort procedure (pass the smaller
        ' subdivision first to use less stack space):
        If (i - Low) < (High - i) Then
          Quick SortArray(), Low, i - 1
          Quick SortArray(), i + 1, High
        Else
          Quick SortArray(), i + 1, High
          Quick SortArray(), Low, i - 1
        End If
      End If
    End If
End Sub

Public Sub Shell(ByRef SortArray() As Integer)
' =============================== ShellSort ==================================
'  The ShellSort procedure is similar to the BubbleSort procedure.  However,
'  ShellSort begins by comparing elements that are far apart (separated by
'  the value of the Offset variable, which is initially half the distance
'  between the first and last element), then comparing elements that are
'  closer together (when Offset is one, the last iteration of this procedure
'  is merely a bubble sort).
' ============================================================================
  Dim Row    As Integer
  Dim Offset As Integer
  Dim MaxRow As Integer
  Dim Limit  As Integer
  Dim Switch As Integer
  
    ' Set comparison offset to half the number of records in SortArray:
    MaxRow = UBound(SortArray)
    Offset = MaxRow \ 2

    Do While Offset > 0          ' Loop until offset gets to zero.
      Limit = MaxRow - Offset
      Do
        Switch = False         ' Assume no switches at this offset.
        ' Compare elements and switch ones out of order:
        For Row = 1 To Limit
          If SortArray(Row) > SortArray(Row + Offset) Then
            Swap SortArray(Row), SortArray(Row + Offset)
            Switch = Row
          End If
        Next

        ' Sort on next pass only to where last switch was made:
        Limit = Switch - Offset
      Loop While Switch
      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
    Loop
End Sub

Private Sub PercolateDown(ByRef SortArray() As Integer, ByVal MaxLevel As Integer)
' ============================ PercolateDown =================================
'   The PercolateDown procedure restores the elements of SortArray from 1 to
'   MaxLevel to a "heap" (see the diagram with the HeapSort procedure).
' ============================================================================
  Dim i     As Integer
  Dim Child As Integer
  
    i = 1

   ' Move the value in SortArray(1) down the heap until it has
   ' reached its proper node (that is, until it is less than its parent
   ' node or until it has reached MaxLevel, the bottom of the current heap):
    Do
      Child = 2 * i             ' Get the subscript for the child node.

      ' Reached the bottom of the heap, so exit this procedure:
      If Child > MaxLevel Then Exit Do

      ' If there are two child nodes, find out which one is bigger:
      If Child + 1 <= MaxLevel Then
        If SortArray(Child + 1) > SortArray(Child) Then
          Child = Child + 1
        End If
      End If

      ' Move the value down if it is still not bigger than either one of
      ' its children:
      If SortArray(i) < SortArray(Child) Then
        Swap SortArray(i), SortArray(Child)
        i = Child
      ' Otherwise, SortArray has been restored to a heap from 1 to MaxLevel,
      ' so exit:
      Else
        Exit Do
      End If
   Loop
End Sub

Private Sub PercolateUp(ByRef SortArray() As Integer, ByVal MaxLevel As Integer)
' ============================== PercolateUp =================================
'   The PercolateUp procedure converts the elements from 1 to MaxLevel in
'   SortArray into a "heap" (see the diagram with the HeapSort procedure).
' ============================================================================
  Dim i As Integer
  Dim Parent As Integer
  
    i = MaxLevel

    ' Move the value in SortArray(MaxLevel) up the heap until it has
    ' reached its proper node (that is, until it is greater than either
    ' of its child nodes, or until it has reached 1, the top of the heap):
    Do Until i = 1
      Parent = i \ 2            ' Get the subscript for the parent node.

      ' The value at the current node is still bigger than the value at
      ' its parent node, so swap these two array elements:
      If SortArray(i) > SortArray(Parent) Then
        Swap SortArray(Parent), SortArray(i)
        i = Parent
      ' Otherwise, the element has reached its proper place in the heap,
      ' so exit this procedure:
      Else
        Exit Do
      End If
   Loop
End Sub

Private Function RandInt%(ByVal lower As Integer, ByVal Upper As Integer)
   RandInt% = Int(Rnd * (Upper - lower + 1)) + lower
End Function

