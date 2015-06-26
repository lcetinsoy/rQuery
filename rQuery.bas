Attribute VB_Name = "libRange"
Option Explicit


Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function WorkbookNamedRanges() As String()

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

Function rngEmptyRight(rngStart As Range) As Range
    
    Set rngEmptyRight = Range(rngStart, rngStart.End(xlToRight)(1, 0))

End Function

Function GetWorksheetNamedRanges(wsSource As Worksheet, Optional nameFilter As String = "") As Collection
    
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
    
    Set GetWorksheetNamedRanges = nameCollection
End Function

Function rngConvertToDoubleArray(rngSource As Range, Optional blTranspose As Boolean = False) As Double()
    Dim nCol As Integer, nRow As Integer
    Dim iCol As Integer, iRow As Integer
    Dim myArray() As Double
    
    nRow = rngSource.Rows.Count
    nCol = rngSource.Columns.Count
        
        
    If nCol = 1 Or nRow = 1 Then
        Dim nCells As Integer
        nCells = WorksheetFunction.max(nRow, nCol)
        ReDim myArray(nCells)
                    
        For iRow = 1 To nCells
            
            myArray(iRow) = rngSource.Cells(iRow)
            
        Next

    Else
        If blTranspose Then

            For iRow = 1 To nRow
                
                For iCol = 1 To nCol
                    myArray(iRow, iCol) = rngSource(iRow, iCol)
                Next
                
            Next
        Else
            
            ReDim myArray(nCol, nRow)
                    
            For iRow = 1 To nRow
                For iCol = 1 To nCol
                    myArray(iCol, iRow) = rngSource(iRow, iCol)
                Next
            Next
        
        End If
        
    End If
    
    
    rngConvertToDoubleArray = myArray
End Function

Public Sub displayLine(rngFirstCell As Range, arrayValue)
    
    Dim cellValue
    Dim rngCell As Range
    Set rngCell = rngFirstCell
    
    For Each cellValue In arrayValue
        
        rngCell.value = cellValue
        Set rngCell = rngCell.Offset(0, 1)
            
    Next

End Sub

Sub DisplayElements(elements As Variant, rngFirstCell As Range, Optional direction As XlDirection = xlDown)
    
    Dim elt
    Dim iElt As Integer
    
    iElt = 1
    
    If direction = xlDown Then
    
        For Each elt In elements
            rngFirstCell(iElt) = elt
            iElt = iElt + 1
        Next
        
    ElseIf direction = xlToRight Then
        
        For Each elt In elements
            rngFirstCell(1, iElt) = elt
            iElt = iElt + 1
        Next
        
    End If
    
End Sub


Public Function rngTransposeToH(rngVertical As Range, rngRes_FC As Range, Optional delete As Boolean = False)

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


Option Explicit
Option Base 1

'TODO: coder un rngRight, left, etc, avec un arrêt par valeur Range(cell, cell.find)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngUp: Return a vertical range containing all values above a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngUp(rngStartCell As Range, Optional nCell As Integer = 0) As Range

    
    If nCell = 0 Then
        If IsEmpty(rngStartCell.Offset(-1, 0)) Then
        
            Set rngUp = rngStartCell
        Else
            Set rngUp = Range(rngStartCell, rngStartCell.End(xlUp))
        End If
    Else
    
        Set rngUp = Range(rngStartCell, rngStartCell.Offset(-nCell + 1, 0))
    End If
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngDown: Return a vertical range containing all values below a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngDown(rngStartCell As Range, Optional nCell As Integer = 0) As Range
        
    If nCell = 0 Then
        If IsEmpty(rngStartCell.Offset(1, 0)) Then
        
            Set rngDown = rngStartCell
        Else
            Set rngDown = Range(rngStartCell, rngStartCell.End(xlDown))
        End If
    Else
        
        Set rngDown = Range(rngStartCell, rngStartCell.Offset(nCell - 1, 0))
    End If
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngRight: Return an horizontal range containing all cells after a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngRight(rngStartCell As Range, Optional nCell As Integer = 0) As Range

    
    If nCell = 0 Then
        If IsEmpty(rngStartCell.Offset(0, 1)) Then
        
            Set rngRight = rngStartCell
        Else
            Set rngRight = Range(rngStartCell, rngStartCell.End(xlToRight))
        End If
    Else
        
        Set rngRight = Range(rngStartCell, rngStartCell.Offset(0, nCell - 1))
    End If
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'rngRight: Return an horizontal range containing all cells before a given cell (included)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngLeft(rngStartCell As Range, Optional nCell As Integer = 0) As Range

    If nCell = 0 Then
        If IsEmpty(rngStartCell.Offset(0, -1)) Then
        
            Set rngLeft = rngStartCell
        Else
            Set rngLeft = Range(rngStartCell, rngStartCell.End(xlToLeft))
        End If
    Else
        
        Set rngLeft = Range(rngStartCell, rngStartCell.Offset(0, -nCell + 1))
    End If
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngArray(rngStartCell As Range, Optional nRow As Integer = 0, Optional nCol As Integer = 0) As Range
    
    Dim rngEndCell As Range
    
    If IsEmpty(rngStartCell.Offset(1, 0)) And IsEmpty(rngStartCell.Offset(0, 1)) Then
            
        Set rngEndCell = rngStartCell
    
    ElseIf IsEmpty(rngStartCell.Offset(1, 0)) Then
        
        Set rngEndCell = rngStartCell.End(xlToRight)
    
    ElseIf IsEmpty(rngStartCell.Offset(0, 1)) Then
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
Public Function dispArray(array1D, rngFirstCell As Range)
    
    Dim elt
    Dim iElt As Integer
    
    iElt = 1
    For Each elt In array1D
        rngFirstCell(iElt) = elt
        Inc iElt
    Next
         
End Function

Sub DispElements(elements As Variant, rngFirstCell As Range, Optional direction As XlDirection = xlDown)
    
    Dim elt
    Dim iElt As Integer
    
    iElt = 1
    
    If direction = xlDown Then
    
        For Each elt In elements
            rngFirstCell(iElt) = elt
            Inc iElt
        Next
        
    ElseIf direction = xlToRight Then
        
        For Each elt In elements
            rngFirstCell(, iElt) = elt
            Inc iElt
        Next
        
    End If
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function display2DArray(array2D, rngFirstCell)

    Range(rngFirstCell, rngFirstCell.Offset(UBound(array2D, 1) - 1, UBound(array2D, 2) - 1)) = array2D

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function display3DArray(array3D, rngFirstCells)

    Dim iCell As Integer
    Dim iRow As Integer, iCol As Integer
    
    Dim rng As Range
    iCell = 1
    For Each rng In rngFirstCells
    
        For iRow = 1 To UBound(array3D, 1)
        
            For iCol = 1 To UBound(array3D, 2)
            
                rng.Offset(iRow - 1, iCol - 1).value = array3D(iRow, iCol, iCell)
            
            Next
        
        Next
        iCell = iCell + 1
    Next

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFirstNonBlankValueIndex(rngSource As Range) As Integer
    
    Dim rngIt As Range
    Dim iValue As Integer
    iValue = 1
    For Each rngIt In rngSource
        
        If Not IsEmpty(rngIt) Then
            GetFirstNonBlankValueIndex = iValue
            Exit Function
        End If
        Inc iValue
    Next


End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFirstIndexOfValueDifferent(rngSource As Range, valueToLook) As Integer
    
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

Function FindRangeDifferentThan(valueToSkip, rngSource) As Range
        
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
Public Function GetNumberOfSameValues(rngToLookFC As Range, Optional direction As XlDirection = xlDown) As Integer

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
Function SelectColumnsFromArray(rngSourceArray As Range, rngIncludeColumn As Range) As Range
    
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

Function rngSelectRows(rngSourceArray As Range, rngIncludeRows As Range) As Range
    
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
    
    Do While Not IsEmpty(rngIt)
        
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
Public Function InsertRowInRange(rngRowCellWhereToInsert As Range, rngValuesToInsert As Range)
        
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
Public Sub CopyAndPasteRange(rngRangeToCopy As Range, rngDestFC As Range, Optional blTranspose As Boolean = False)

    rngRangeToCopy.Copy
    
    rngDestFC.PasteSpecial xlPasteValues, , , blTranspose
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CountDimensions: Count the dimension number of an Array
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CountDimensions(arraySource) As Integer
    Dim res
    Dim DimNum As Integer
      'Sets up the error handler.
    On Error GoTo FinalDimension
    For DimNum = 1 To 100

         'It is necessary to do something with the LBound to force it
         'to generate an error.
        res = LBound(arraySource, DimNum)

    Next DimNum

    Exit Function

      ' The error routine.
FinalDimension:

    CountDimensions = DimNum - 1

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DisplayObjectFamily(familyToDisplay As Variant, rngDestFC As Range)
    
    Dim iterator
    Dim iValue As Integer
    
    iValue = 1
    For Each iterator In familyToDisplay
        rngDestFC(iValue).value = iterator
        Inc iValue
    Next

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
    
    If IsEmpty(rngSourceFC) Then
        Set NewRowRange = rngSourceFC
    ElseIf IsEmpty(rngSourceFC.Offset(1, 0)) Then
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

Function rngDownValueTst()

    With wsTests
        .UsedRange.ClearContents
        dispArray Array(1, 1, 1, 3), .Range("A1")
        .Range("C1") = rngDownValue(.Range("A1"), 1).Address
    End With
    
End Function

Function rngEnd(rngSourceFC As Range, Optional stopCond As Variant = Nothing, Optional direction As XlDirection = xlDown) As Range
    
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
         If Not IsEmpty(rngSourceFC) And IsEmpty(rngSourceFC.Offset(1, 0)) Then
            Set rngIt = rngSourceFC
        ElseIf IsEmpty(rngSourceFC) And IsEmpty(rngSourceFC.Offset(2, 0)) Then
            Set rngIt = rngSourceFC.Offset(1, 0)
        ElseIf IsEmpty(rngSourceFC) Then
            Set rngIt = rngSourceFC.Offset(1, 0).End(direction)
        Else
            Set rngIt = rngSourceFC.End(direction)
        End If
    End If
               
    
    Set rngEnd = rngIt
End Function

Sub rngEndTst()

    With wsTests
        .UsedRange.ClearContents
        dispArray Array(1, 1, 1, 3), .Range("A1")
        .Range("C1").value = rngEnd(.Range("A1"), "OnValChange").Address
    End With
End Sub


Function rngSubArray(rngArray, rowIndexes, colIndexes) As Range

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

Function rngIsEmpty(rngSource As Range) As Boolean

    If WorksheetFunction.Count(rngSource) = 0 Then
    
        rngIsEmpty = False
        Exit Function
        
    End If
    
    rngIsEmpty = True
    
End Function



Public Sub NameAdd(rngCellsToName As Range, strBasename As String)

    Dim rngIter As Range
    Dim iCell As Integer
    
    iCell = 1
    For Each rngIter In rngCellsToName
    
        rngIter.name = strBasename & CStr(iCell)
        iCell = iCell + 1
    Next
    

End Sub
Public Function NameDelete(rngCells As Range)

    Dim rngIter As Range
    Dim iCell As Integer
    On Error Resume Next
    For Each rngIter In rngCells
    
        names(rngIter.name).delete
    
    Next
    On Error GoTo 0
End Function

Public Function rngSort(rngToSort As Range, colSort As Integer, Optional order As XlSortOrder = xlAscending)

    rngToSort.Sort rngToSort.Columns(colSort), order
    
End Function



