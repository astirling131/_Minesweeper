Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

'' DEBUG FLAG
Dim debug_flag As String
debug_flag = ThisWorkbook.Sheets("settings").Range("$D$2").Value
If debug_flag = "On" Then
    Exit Sub
End If

'-----------------------------------------------------------------------------
Dim wb As Workbook, ws As Worksheet
Dim rng As Range, lookupRange As Range
Dim cellVal As String, addr As String, namedRangeName As String, nmRange As String
Dim xCount As Long
Dim shp As Shape
  
Set wb = ThisWorkbook
Set ws = wb.Sheets("new_game")

ws.Unprotect

Application.ScreenUpdating = False

addr = Target.Address

Set lookupRange = wb.Sheets("settings").Range("A:B")
On Error Resume Next
namedRangeName = Application.WorksheetFunction.VLookup(addr, lookupRange, 2, False)
On Error GoTo 0

nmRange = "group" & namedRangeName

If nmRange = "group" Then
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub
Else
    Set rng = ws.Range(namedRangeName)
End If

'Check the value of the cell
cellVal = rng.Value

If cellVal = "X" Then
    ' Game ends if the user clicks a bomb
    Call minesweeperBomb
Else
    'Else check the number of surrounding X's
    xCount = CountAdjacentXs(rng)
    'Update the cell to that value
    rng = xCount
    ws.Range(addr).Select

    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    
    'Blank it if there were no X's
    If xCount = 0 Then
        SendKeys "{ENTER}"
        rng = ""
    End If
    'Delete the shapes to show the cell has been uncovered
    For Each shp In ActiveSheet.Shapes
        If shp.Name = nmRange Then
            shp.Delete
            Exit For
        End If
    Next shp
End If

ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

Application.ScreenUpdating = True

SendKeys "{ENTER}"

End Sub
