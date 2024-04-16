Attribute VB_Name = "modSearch"
Public Function BinarySearch( _
    SearchFor As String, _
    vArray As Variant, _
    Optional verbose As Boolean = True _
    ) As Long:
'/  Created On:  04/17/2019                  Last Modified: 04/17/2019
'/  Description: Efficient & Quick Search Algorithm For Locating and Identifying
'/               An Item Within a Sorted Single Dimension Array
'/==================================================================================
Dim MinIndex    As Long
Dim MaxIndex    As Long
Dim MidPoint    As Long
Dim IsMultiArr  As Long
Dim i           As Long
    
    '# [Validation] Check if Array is Multi-Dim
    On Error Resume Next
    IsMultiArr = (UBound(vArray, 2) > 1)
    If IsMultiArr Then
        If pLog Then Debug.Print Now(); " [Error]<BinarySearch> Multi Dim Array"
        BinarySearch = -1
        Exit Function
    End If
    On Error GoTo 0
    
    MinIndex = LBound(vArray, 1)
    MaxIndex = UBound(vArray, 1)
    MidPoint = Round((MinIndex + MaxIndex / 2), 0)
    
    If SearchFor = vArray(MidPoint) Then
        BinarySearch = MidPoint
        Exit Function
    End If
    
    If verbose Then Debug.Print vbTab, "Loop#", "MinIndex", "MidPoint", "MaxIndex", "Value"
    i = 0
    While MinIndex <> MaxIndex
        i = i + 1
        If pLog Then Debug.Print vbTab, i, MinIndex, MidPoint, MaxIndex, "v:="; vArray(MidPoint)
        If (SearchFor < vArray(MidPoint)) Then
            '# Less Than Midpoint Value
            MaxIndex = MidPoint - 1
            MidPoint = Round((MinIndex + MaxIndex + 1) / 2, 0)
        
        ElseIf (SearchFor > vArray(MidPoint)) Then
            '# Greater Than Midpoint Value
            MinIndex = MidPoint + 1
            MidPoint = Round((MinIndex + MaxIndex + 1) / 2, 0) - 1
        
        Else
            '# Match Condition
            If pLog Then Debug.Print vbCrLf; Now(); "<BinarySearch> Item Found"
            BinarySearch = MidPoint
            Exit Function
        End If
    Wend
    BinarySearch = -1
    If pLog Then Debug.Print vbCrLf; Now(); "<BinarySearch> Item Not Found"
End Function

