Attribute VB_Name = "modSort"
Public Sub QuickSort(vArray As Variant, inLow As Double, inHigh As Double)
'/  Created On:     03/11/2019                      Last Modified: 03/11/2019
'/  Description:    Sorts Single Dimension Array. Quicksort is One of Fastest Methods
'/                  For Sorting Data in Arrays. Average Run-Time for the Algorithm
'/                  is O(n log n) w/ the Worst-Case Sort Time  Being O(n^2)
'/
'/  Ref Link: [https://stackoverflow.com/questions/152319/vba-array-sort-function]
'/==================================================================================
Dim pivot   As Variant
Dim tmpSwap As Variant
Dim tmpLow  As Double
Dim tmpHigh As Double
    
    tmpLow = inLow
    tmpHigh = inHigh
    
    pivot = vArray((inLow + inHigh) \ 2)
    While (tmpLow <= tmpHigh)
        While (vArray(tmpLow) < pivot And tmpLow < inHigh)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHigh) And tmpHigh > inLow)
            tmpHigh = tmpHigh - 1
        Wend
        
        If (tmpLow <= tmpHigh) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHigh)
            vArray(tmpHigh) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHigh = tmpHigh - 1
        End If
    Wend
    
    If (inLow < tmpHigh) Then QuickSort vArray, inLow, tmpHigh
    If (tmpLow < inHigh) Then QuickSort vArray, tmpLow, inHigh
End Sub
