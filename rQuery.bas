Attribute VB_Name = "rQ"
Option Explicit


Function rngEmptyRight(rngStart As Range) As Range
    
    Set rngEmptyRight = Range(rngStart, rngStart.End(xlToRight)(1, 0))

End Function


'TODO: coder un rngRight, left, etc, avec un arrêt par valeur Range(cell, cell.find)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngUp: Return a vertical range containing all values above a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function columnUp(rngStartCell As Range, Optional nCell As Integer = 0) As Range

    Dim rngUp As Range
    If nCell = 0 Then
        If isEmpty(rngStartCell.Offset(-1, 0)) Then
        
            Set rngUp = rngStartCell
        Else
            Set rngUp = Range(rngStartCell, rngStartCell.End(xlUp))
        End If
    Else
    
        Set rngUp = Range(rngStartCell, rngStartCell.Offset(-nCell + 1, 0))
    End If
    
    Set columnUp = rngUp
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngDown: Return a vertical range containing all values below a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function columnDown(rngStartCell As Range, Optional nCell As Integer = 0) As Range
        
    Dim rng As Range
    If nCell = 0 Then
        If isEmpty(rngStartCell.Offset(1, 0)) Then
        
            Set rng = rngStartCell
        Else
            Set rng = Range(rngStartCell, rngStartCell.End(xlDown))
        End If
    Else
        
        Set rng = Range(rngStartCell, rngStartCell.Offset(nCell - 1, 0))
    End If
    
    Set columnDown = rng
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngRight: Return an horizontal range containing all cells after a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rowRight(rngStartCell As Range, Optional nCell As Integer = 0) As Range

    Dim rngRight As Range
    If nCell = 0 Then
        If isEmpty(rngStartCell.Offset(0, 1)) Then
        
            Set rngRight = rngStartCell
        Else
            Set rngRight = Range(rngStartCell, rngStartCell.End(xlToRight))
        End If
    Else
        
        Set rngRight = Range(rngStartCell, rngStartCell.Offset(0, nCell - 1))
    End If
    Set rowRight = rngRight
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngRight: Return an horizontal range containing all cells before a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rowLeft(rngStartCell As Range, Optional nCell As Integer = 0) As Range
Dim rngLeft As Range
    If nCell = 0 Then
        If isEmpty(rngStartCell.Offset(0, -1)) Then
        
            Set rngLeft = rngStartCell
        Else
            Set rngLeft = Range(rngStartCell, rngStartCell.End(xlToLeft))
        End If
    Else
        
        Set rngLeft = Range(rngStartCell, rngStartCell.Offset(0, -nCell + 1))
    End If
    
    Set rowLeft = rngLeft
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngArray(rngStartCell As Range, Optional nRow As Integer = 0, Optional nCol As Integer = 0) As Range
    
    Dim rngEndCell As Range
    
    If isEmpty(rngStartCell.Offset(1, 0)) And isEmpty(rngStartCell.Offset(0, 1)) Then
            
        Set rngEndCell = rngStartCell
    
    ElseIf isEmpty(rngStartCell.Offset(1, 0)) Then
        
        Set rngEndCell = rngStartCell.End(xlToRight)
    
    ElseIf isEmpty(rngStartCell.Offset(0, 1)) Then
        Set rngEndCell = rngStartCell.End(xlDown)
    
    Else
        Set rngEndCell = rngStartCell.End(xlDown).End(xlToRight)
    End If
    
    If nRow = 0 And nCol = 0 Then
        
        Set rngArray = Range(rngStartCell, rngEndCell)
    
    ElseIf nRow = 0 And nCol > 0 Then
        
        Set rngArray = Range(rngStartCell, rngStartCell.End(xlDown).Offset(0, nCol - 1))
        
    ElseIf nRow > 0 And nCol = 0 Then
    
        Set rngArray = Range(rngStartCell, rngStartCell.End(xlToRight).Offset(nRow - 1, 0))
        
    ElseIf nRow > 0 And nCol > 0 Then
    
        Set rngArray = Range(rngStartCell, rngStartCell.Offset(nRow - 1, nCol - 1))
        
    Else
    
        Err.Raise 513, , "Wrong inputs for nRow and nCol : nRow and nCol are number >= 0"
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngArrayFromEnd(rngLastCell As Range, Optional nRow As Integer = 0, Optional nCol As Integer = 0) As Range

    If nRow = 0 And nCol = 0 Then
    
        Set rngArrayFromEnd = Range(rngLastCell, rngLastCell.End(xlUp).End(xlToLeft))
    
    ElseIf nRow = 0 And nCol > 0 Then
        
        Set rngArrayFromEnd = Range(rngLastCell, rngLastCell.End(xlUp).Offset(0, -nCol + 1))
        
    ElseIf nRow > 0 And nCol = 0 Then
    
        Set rngArrayFromEnd = Range(rngLastCell, rngLastCell.End(xlToLeft).Offset(-nRow + 1, 0))
        
    ElseIf nRow > 0 And nCol > 0 Then
    
        Set rngArrayFromEnd = Range(rngLastCell, rngLastCell.Offset(-nRow + 1, -nCol + 1))
        
    Else
    
        MsgBox "Wrong inputs for nRow and nCol : nRow and nCol are number >= 0"
    End If
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function indexOfFirtNonEmptyCell(rngSource As Range) As Integer
    
    Dim rngIt As Range
    Dim iValue As Integer
    iValue = 1
    For Each rngIt In rngSource
        
        If Not isEmpty(rngIt) Then
            GetFirstNonBlankValueIndex = iValue
            Exit Function
        End If
        Inc iValue
    Next


End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function indexOfValueDifferent(rngSource As Range, valueToLook) As Integer
    
    Dim rngIt As Range
    Dim iValue As Integer
    iValue = 1
    For Each rngIt In rngSource.Cells
        
        If rngIt.value <> valueToLook Then
            GetFirstIndexOfValueDifferent = iValue
            Exit Function
        End If
        Inc iValue
    Next


End Function

Function findCellDifferentOf(valueToSkip, rngSource) As Range
        
    Dim rngIt As Range
        
    For Each rngIt In rngSource.Cells
        
        If rngIt.value <> valueToSkip Then
            Set FindRangeDifferentThan = rngIt
            Exit Function
        End If
        
    Next

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function countIdenticalValues(rngToLookFC As Range, Optional direction As XlDirection = xlDown) As Integer

    Dim ValSource
    Dim nValue As Integer
    Dim rngIt As Range
    
    ValSource = rngToLookFC.value
    nValue = 0
    Set rngIt = rngToLookFC
    
    If direction = xlDown Then
    
        Do While rngIt.value = ValSource
    
            Inc nValue
            Set rngIt = rngIt.Offset(1, 0)
            
        Loop
        
    ElseIf direction = xlToRight Then
    
        Do While rngIt.value = ValSource
    
            Inc nValue
            Set rngIt = rngIt.Offset(0, 1)
            
        Loop
    
    End If
            
    GetNumberOfSameValues = nValue
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function selectArrayColumns(rngSourceArray As Range, rngIncludeColumn As Range) As Range
    
    Dim rngColumn As Range, rngEffectiveSource As Range
    Dim iCol As Integer
    
    iCol = 1
    
    For iCol = 1 To rngSourceArray.Columns.Count
        
        Set rngColumn = rngSourceArray.Columns(iCol).Cells
        If rngIncludeColumn(1, iCol) = True Then
            If rngEffectiveSource Is Nothing Then
            
                Set rngEffectiveSource = rngColumn
            Else
                Set rngEffectiveSource = Union(rngEffectiveSource, rngColumn)
            End If
        End If
        
    Next
    Set SelectColumnsFromArray = rngEffectiveSource
End Function

Function rngSelectArrayRows(rngSourceArray As Range, rngIncludeRows As Range) As Range
    
    Dim rngRow As Range, rngEffectiveSource As Range
    Dim iRow As Integer
    
    For iRow = 1 To rngSourceArray.Rows.Count
        
        Set rngRow = rngSourceArray.Rows(iRow).Cells
        If rngIncludeRows(iRow, 1) = True Then
            If rngEffectiveSource Is Nothing Then
            
                Set rngEffectiveSource = rngRow
            Else
                Set rngEffectiveSource = Union(rngEffectiveSource, rngRow)
            End If
        End If
        
    Next
    Set rngSelectRows = rngEffectiveSource

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyColAndInsertValue(rngColRef As Range, rngColDest As Range, rngDest As Range, valToInsert, valWhereToInsert) As Range

    Dim rngIt As Range
    Dim added As Boolean
    Dim iVal As Integer
    
    added = False
    iVal = 1
    
    For iVal = 1 To rngColSource.Cells.Count + 1
        Set rngIt = rngColSource(iVal)
        rngDest(iVal).value = rngIt
        If rngIt.Offset(1, 0) > valWhereToInsert And added = False And IsNumeric(rngIt.Offset(1, 0)) Then
            
            Inc iVal
            rngDest(iVal) = valToInsert
            Set InsertValueInRange = rngIt
            added = True
                        
        End If
    Next
    
    If added = False Then
        rngDest(iVal - 1) = valToInsert
        Set InsertValueInRange = rngColSource.End(xlDown)
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FindIndexCell(rngCol As Range, ValSource As Double) As Range

    Dim rngIt As Range
        
    For Each rngIt In rngCol
        If rngIt.Offset(1, 0).value > ValSource Then
            Set FindIndexCell = rngIt
            Exit Function
        End If
    Next
    If rngIt Is Nothing Then
        Set FindIndexCell = rngCol.End(xlDown)
    End If
    
End Function
Function rngAddToUnion(rangeToAdd As Range, unionToAppend As Range)

    Set rngAddToUnion = Union2(unionToAppend, rangeToAdd)
    
End Function

Function rngBuildUnion(rangesToAddToUnion As Collection, Optional unionSource As Range = Nothing)
        
    Dim myUnion As Range
    Set myUnion = Nothing
    For Each rng In rangesToAddToUntion
        Set myUnion = Union2(myUnion, rng)
    Next
    Set rngBuildUnion = myUnion
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildUnionOfTablesFCS(rngFirstTableFC As Range, nBlankCellsBetweenTables As Integer) As Range
    
    Dim rngIt As Range
            
    Set rngIt = rngFirstTableFC
    
    Do While Not isEmpty(rngIt)
        
        If BuildUnionOfTablesFCS Is Nothing Then
        
            Set BuildUnionOfTablesFCS = rngFirstTableFC
        Else
        
            Set BuildUnionOfTablesFCS = Union(BuildUnionOfTablesFCS, rngIt)
            Set rngIt = rngIt.End(xlDown).Offset(nBlankCellsBetweenTables + 1, 0)
        End If
    Loop

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function insertRowInRange(rngRowCellWhereToInsert As Range, rngValuesToInsert As Range)
        
    Dim rngRow As Range
    Set rngRow = rngRight(rngRowCellWhereToInsert)
    
    rngValuesToInsert.Copy
    rngRow.Insert xlShiftDown
              
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyAndPasteRange: Copy a source range to any destination range
'@input rngRangeToCopy range: The range to copy'
'@input rngDestFC range: The first cell of the destination range
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub copyAndPaste(rngRangeToCopy As Range, rngDestFC As Range, Optional blTranspose As Boolean = False)

    rngRangeToCopy.Copy
    
    rngDestFC.PasteSpecial xlPasteValues, , , blTranspose
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NewColumnRange: Return a cell range to a new column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function NewColumnRange(rngSourceFC As Range) As Range
    
    Set NewColumnRange = rngSourceFC.End(xlToRight).Offset(0, 1)
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NewRowRange: Return the cell range corresponding to the next row of the last array row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function NewRowRange(rngSourceFC As Range) As Range
    
    If isEmpty(rngSourceFC) Then
        Set NewRowRange = rngSourceFC
    ElseIf isEmpty(rngSourceFC.Offset(1, 0)) Then
        Set NewRowRange = rngSourceFC.Offset(1, 0)
    Else
        Set NewRowRange = rngSourceFC.End(xlDown).Offset(1, 0)
    End If
End Function

Function rngDownValue(rngFC As Range, value As Variant) As Range


    Dim rngIt As Range
    Set rngIt = rngFC
    Do While rngIt.Offset(1, 0) = value
    
        Set rngIt = rngIt.Offset(1, 0)
    Loop
    Set rngDownValue = Range(rngFC, rngIt)
End Function


Function tail(rngSourceFC As Range, Optional stopCond As Variant = Nothing, Optional direction As XlDirection = xlDown) As Range
    
    Dim rngIt As Range
       
    If Not IsObject(stopCond) Then
        
        If stopCond <> "OnValChange" Then
            Set rngIt = rngSourceFC
            If direction = xlDown Then
        
                Do While rngIt.Offset(1, 0) = stopCond
    
                    Set rngIt = rngIt.Offset(1, 0)
                Loop
                
            ElseIf direction = xlToRight Then
                    
                Do While rngIt.Offset(0, 1) = stopCond
    
                    Set rngIt = rngIt.Offset(0, 1)
                Loop
            End If
            
       ElseIf stopCond = "OnValChange" Then
        
            Set rngIt = rngSourceFC
            If direction = xlDown Then
        
                Do While rngIt.Offset(1, 0) = rngSourceFC.value
    
                    Set rngIt = rngIt.Offset(1, 0)
                Loop
                
            ElseIf direction = xlToRight Then
                    
                Do While rngIt.Offset(0, 1) = rngSourceFC.value
    
                    Set rngIt = rngIt.Offset(0, 1)
                Loop
            End If
        End If
    Else
    
        If direction = xlToRight Then
            
            If Not isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(0, 1)) Then
                Set rngIt = rngSourceFC
                
            ElseIf isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(0, 1)) Then
            
                Set rngIt = rngSourceFC.Offset(0, 1)
                
            Else
                Set rngIt = rngSourceFC.End(direction)
            End If
                    
        ElseIf direction = xlToLeft Then
        
            If Not isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(0, -1)) Then
                Set rngIt = rngSourceFC
                
            ElseIf isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(0, -1)) Then
            
                Set rngIt = rngSourceFC.Offset(0, -1)
                
            Else
                Set rngIt = rngSourceFC.End(direction)
            End If
                    
        
        ElseIf direction = xlDown Then
    
             If Not isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(1, 0)) Then
                Set rngIt = rngSourceFC
            ElseIf isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(2, 0)) Then
                Set rngIt = rngSourceFC.Offset(1, 0)
            ElseIf isEmpty(rngSourceFC) Then
                Set rngIt = rngSourceFC.Offset(1, 0).End(direction)
            Else
                Set rngIt = rngSourceFC.End(direction)
            End If
            
        ElseIf direction = xlUp Then
        
             If Not isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(-1, 0)) Then
                Set rngIt = rngSourceFC
            ElseIf isEmpty(rngSourceFC) And isEmpty(rngSourceFC.Offset(-2, 0)) Then
                Set rngIt = rngSourceFC.Offset(-1, 0)
            ElseIf isEmpty(rngSourceFC) Then
                Set rngIt = rngSourceFC.Offset(-1, 0).End(direction)
            Else
                Set rngIt = rngSourceFC.End(direction)
            End If
            
        End If
        
    End If
                   
    Set tail = rngIt
End Function

Function subArrayRange(rngArray, rowIndexes, colIndexes) As Range

    Set rngSubArray = Range(rngArray(rowIndexes(1), colIndexes(1)), rngArray(rowIndexes(2), colIndexes(2)))
    
End Function


Function Union2(ParamArray Ranges() As Variant) As Range
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Union2
   ' A Union operation that accepts parameters that are Nothing.
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       Dim n As Long
       Dim RR As Range
       For n = LBound(Ranges) To UBound(Ranges)
           If IsObject(Ranges(n)) Then
               If Not Ranges(n) Is Nothing Then
                   If TypeOf Ranges(n) Is Excel.Range Then
                       If Not RR Is Nothing Then
                           Set RR = Application.Union(RR, Ranges(n))
                       Else
                           Set RR = Ranges(n)
                       End If
                   End If
               End If
           End If
       Next n
       Set Union2 = RR
   End Function
    

Function ProperUnion(ParamArray Ranges() As Variant) As Range
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProperUnion
' This provides Union functionality without duplicating
' cells when ranges overlap. Requires the Union2 function.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim ResR As Range
    Dim n As Long
    Dim R As Range
    
    If Not Ranges(LBound(Ranges)) Is Nothing Then
        Set ResR = Ranges(LBound(Ranges))
    End If
    For n = LBound(Ranges) + 1 To UBound(Ranges)
        If Not Ranges(n) Is Nothing Then
            For Each R In Ranges(n).Cells
                If Application.Intersect(ResR, R) Is Nothing Then
                    Set ResR = Union2(ResR, R)
                End If
            Next R
        End If
    Next n
    Set ProperUnion = ResR
End Function

Function isRangeEmpty(rngSource As Range) As Boolean

    If WorksheetFunction.Count(rngSource) = 0 Then
    
        rngIsEmpty = False
        Exit Function
        
    End If
    
    rngIsEmpty = True
    
End Function


Public Function sortRange(rngToSort As Range, colSort As Integer, Optional order As XlSortOrder = xlAscending)

    rngToSort.Sort rngToSort.Columns(colSort), order
    
End Function




Function getWorksheetNamedRanges(wsSource As Worksheet, Optional nameFilter As String = "") As Collection
    
    Dim myName As name
    Dim nameCollection As New Collection
    
    For Each myName In wsSource.names
        If nameFilter <> "" Then
        
            If InStr(1, myName.name, nameFilter) Then
                nameCollection.Add Range(myName), Split(myName.name, "!")(1)
            End If
        Else
            nameCollection.Add Range(myName), Split(myName.name, "!")(1)
        End If
    Next
    
    Set getWorksheetNamedRanges = nameCollection
End Function


Public Function columnToRow(rngVertical As Range, rngRes_FC As Range, Optional delete As Boolean = False)

    Dim rngIter As Range
    Dim iCell As Integer
    iCell = 1
    For Each rngIter In rngVertical
        
        rngRes_FC.Offset(0, iCell - 1).value = rngIter.value
        iCell = iCell + 1
    
    Next
    
    If delete = True Then
        rngVertical.ClearContents
    End If
    
End Function

Function getWorkbookNamedRanges() As String()

    Dim myName As name
    Dim nName As Integer: nName = ThisWorkbook.names.Count
    Dim names() As String
    ReDim names(nName)
    Dim iName
    For iName = 1 To nName
        
        names(iName - 1) = ThisWorkbook.names(iName).name
        
    Next
    
    WorkbookNamedRanges = names
End Function

