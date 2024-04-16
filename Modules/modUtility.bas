Attribute VB_Name = "modUtility"
'/  [Module]        Excel General Utility Functions & Procedures
'/  Created By:     Ryan Tamashiro
'/  Contact Info:   Tamashiroryan@gmail.com
'/**********************************************************************************
'/  **READ ME**
'/
'/  Functions:
'/  01. LastColumn                      / Last populated column of specified row
'/  02. LastRow                         / Last populated row of specified column
'/
'/**********************************************************************************

Option Explicit

Public Enum xlToggleSettings:
    xlEnable = -1
    xlDisable = 0
End Enum

Public Enum xlCellTextAlignment:
    txtAlignLeft = 1
    txtAlignRight = 2
    txtAlignCenter = 3
End Enum

Public Function LastColumn(wsRef As Worksheet, Optional rowRef As Long = 1) As Integer
'/  Created On:     01/07/2021          Last Modified:  01/07/2021
'/----------------------------------------------------------------------------------
'/  Description:    Locates and returns the column index associated with the last
'/                  populated column of a user specified row and worksheet.
'/
'/                  INPUT
'/                      wsRef:      Worksheet Reference
'/                      [rowRef]:   Optional user specified row index reference
'/                                  (Default Row: 1)
'/                  ----------------------------------------------------------------
'/                  OUTPUT
'/                      LastColumn: <Integer> (Small Number Data Type | Max: 32000)
'/                                  Column index assoicated with last populated column
'/                                  in user specified row.
'/                  ----------------------------------------------------------------
'/                  ERROR           >> Return: -1
'/
'/----------------------------------------------------------------------------------
    LastColumn = -1
    If rowRef < 1 Then Exit Function
    LastColumn = wsRef.Cells(rowRef, Columns.Count).End(xlToLeft).Column
End Function


Public Function LastRow(wsRef As Worksheet, Optional colRef As Long = 1) As Long
'/  Created On:     01/07/2021          Last Modified:  01/07/2021
'/----------------------------------------------------------------------------------
'/  Description:    Locates and returns the column index associated with the last
'/                  populated column of a user specified row and worksheet.
'/
'/                  INPUT
'/                      wsRef:      Hiearchichal worksheet reference
'/                      [rowRef]:   Optional user specified row for which the last
'/                                  populated column is to be found.
'/                                  (Default Column: 1)
'/                  ----------------------------------------------------------------
'/                  OUTPUT
'/                      LastColumn: <Integer> (Small Number Data Type | Max: 32000)
'/                                  Column index assoicated with last populated column
'/                                  in user specified row.
'/                  ----------------------------------------------------------------
'/                  ERROR           >> Return: -1
'/
'/----------------------------------------------------------------------------------
    LastRow = -1
    If colRef < 1 Then Exit Function
    LastRow = wsRef.Cells(Rows.Count, colRef).End(xlUp).Row
End Function
