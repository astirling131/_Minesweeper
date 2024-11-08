Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

'' DEBUG FLAG
Dim debug_flag As String
'Check the current debug setting
debug_flag = ThisWorkbook.Sheets("settings").Range("$D$2").Value
If debug_flag = "On" Then
    'Do not run the code if debug mode is set to on.
    Exit Sub
End If

'-----------------------------------------------------------------------------

'Optimise code
Application.ScreenUpdating = False

'Declare variables'
Dim wb As Workbook, ws As Worksheet
Dim rangeStart As Range, randomCell As Range
Dim i As Long, j As Long, xCount As Long
Dim addr As String

'Set the workbook and worksheet objects
Set wb = ThisWorkbook
Set ws = wb.Sheets("new_game")

'Unprotect the game worksheet
ws.Unprotect

'Get the address of the cell that was right clicked
addr = Target.Address

'Add the unicode character for the flag symbol in the cell that was right clicked
ws.Range(addr).Value = ChrW(9873)

'Update the font in the cell so the flag shows
ws.Range(addr).Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

'Re-protect the worksheet
ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

'Cancel the right click action
Cancel = True

'Optimise code
Application.ScreenUpdating = True

End Sub