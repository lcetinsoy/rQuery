rQuery a VBA library for range manipulation
===========================================


Why
---

VBA synthax for selecting ranges can be quite verbose and error prone.

rQuery provides convinient ways to select data ranges.


Installation
------------
Download the file and put it in your vba project.

Usage
-----
set myRowRange = rngRight(mySheet.range("A1")) 'Returns a row containing all cells starting from cell A1 to the last non empty cell

set myRowRange = rngRight(mySheet.range("A1"), 10) ' Returns the range A1: A10

set myColumnRange = rngDown(mySheet.range("B1")) 'Returns a column containing all cells starting from cells B1 to the last non empty cell

set myArrayRange = rngArray(mySheet.range("C1")) 'Return a table range. Selection end to the first empty column or row'

set myArrayRange = rngArray(mySheet.range("C1"), nRow = 10) 'Return a table range with 10 row. Number or column is set automatically


Contributing
------------

Automatic testing are comming
right now juste clone and pull request


Documentation
-------------
```vba
rngRight(rngStartCell As Range, Optional nCell As Integer = 0) as Range
```
Returns a row range starting from rngStartCell going to the right

If nCell is set to 0, rngRight automatically ends the row before the next empty cell

rngLeft(rngStartCell As Range, Optional nCell As Integer = 0) As Range

Returns a row range starting from rngStartCell going to the right

If nCell is set to 0, rngLeft automatically ends the row before the next empty cell

rngDown(rngStartCell As Range, Optional nCell As Integer = 0) As Range

Returns a column range starting from rngStartCell going down

If nCell is set to 0, rngDown automatically ends the column before the next empty cell

rngUp(rngStartCell As Range, Optional nCell As Integer = 0) As Range

Returns a column range starting from rngStartCell going up

If nCell is set to 0, rngDown automatically ends the column before the next empty cell

rngArray(rngStartCell As Range, Optional nRow As Integer = 0, Optional nCol As Integer = 0) As Range

Returns a two dimensional range starting at rngStartCell

Leaving nCol or nRow to 0 tells rngArray to add data to a row (or a column) until the next empty cell (not included)

rngArrayFromEnd(rngLastCell As Range, Optional nRow As Integer = 0, Optional nCol As Integer = 0) As Range

Returns a two dimensional range but the selection is made for the bottom right end corner of the table you want to select

rngEnd(rngSourceFC As Range, Optional stopCond As Variant = Nothing, Optional direction As XlDirection = xlDown) As Range

Returns the last cell of a data set in a given direction
