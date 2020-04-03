Option Explicit

'------------------------------
'Select cell A1 in all sheets 
'------------------------------
Sub moveToA1()

    Application.ScreenUpdating = False
    
    Dim sheetcnt As Integer
    Dim intCnt As Integer

    On Error Resume Next
    sheetcnt = ActiveWorkbook.Sheets.Count

    For intCnt = sheetcnt To 1 Step -1
        If Sheets(intCnt).Visible = True Then
            Sheets(intCnt).Select
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Set the borders of the selection area.
'------------------------------
Sub drawSelectionBorders()

    Application.ScreenUpdating = False

    Dim rngLoop As Range
    Dim rngLastCell As Range
    Dim styleSetArea As Range

    On Error Resume Next

    'Initial 
    Selection.Borders.LineStyle = xlNone
    'The last cell
    Set rngLastCell = Cells(Selection.Rows.Count + Selection.Row - 1, Selection.Columns.Count + Selection.Column - 1)

    For Each rngLoop In Selection.Cells
        If Len(rngLoop) Then
            Set styleSetArea = Range(rngLoop, rngLastCell)
            With styleSetArea
                .Borders.LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        End If
    Next

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Get hyperlink URL
'------------------------------
Sub getHyperlinks()

    Application.ScreenUpdating = False

    Dim rngLoop As Range
    Dim strRightOffset As String
    Dim intRightOffset As Integer

    On Error Resume Next

    'Set right offset number
    strRightOffset = InputBox("Enter the number of columns to be offset to the right.", "Number of offset", "1")
    intRightOffset = IIf(IsNumeric(strRightOffset), strRightOffset, 1)
    
    'Get hyperlink URL
    For Each rngLoop In Selection.Cells
        rngLoop.Offset(0, intRightOffset) = rngLoop.Hyperlinks(1).Address
    Next

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Sort by sheet name
'------------------------------
Sub sortSheet()
    Dim n As Integer
    Dim i As Integer
    Dim R As Variant

    Application.ScreenUpdating = False

    n = Sheets.Count

    'Sheet name
    For i = 1 To n
        Cells(i, Columns.Count).NumberFormatLocal = "@"
        Cells(i, Columns.Count) = Sheets(i).Name
    Next

    'Ascending order
    Cells(1, Columns.Count).Resize(n).Sort Cells(1, Columns.Count)
    R = Cells(1, Columns.Count).Resize(n)
    Cells(1, Columns.Count).Resize(n) = ""

    'Move
    For i = 1 To n
        Sheets(R(n - i + 1, 1)).Move Sheets(1)
    Next

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Add hyperlinks
'------------------------------
Sub AddHyperLink()

    Application.ScreenUpdating = False

    Dim str As String
    str = "Reference"
    If ActiveCell.Text <> "" Then
        str = ActiveCell.Text
    End If
    
    ActiveCell.FormulaR1C1 = "=HYPERLINK(""#"" & CELL(""address"",R[0]C[0]),""" & str & """)"

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Cross selection
'------------------------------
Sub showRowAndCol()

    Application.ScreenUpdating = False

    Dim strAddress As String
    Dim arrStartCellAddress() As String
    Dim arrLastCellAddress() As String
    Dim rngLastCell As Range
    'The last cell
    Set rngLastCell = Cells(Selection.Rows.Count + Selection.Row - 1, Selection.Columns.Count + Selection.Column - 1)
    arrStartCellAddress = Split(ActiveCell.Address, "$")
    arrLastCellAddress = Split(rngLastCell.Address, "$")
    strAddress = ActiveCell.Address & "," & arrStartCellAddress(1) & ":" & arrLastCellAddress(1) & "," & arrStartCellAddress(2) & ":" & arrLastCellAddress(2)

    Range(strAddress).Select

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Auto adjust the height of merged cells
'------------------------------
Sub AutoFitMergedCells()
    Dim oRange As Range
    Dim intLoop As Integer
    Dim oldWidth As Single
    Dim oldZZWidth As Single
    Dim newHeight As Single
    Dim strLen As Single
    Set oRange = Selection

    Application.ScreenUpdating = False

    With ActiveSheet
        oldWidth = 0
        For intLoop = 1 To oRange.Columns.Count
            oldWidth = oldWidth + .Cells(1, oRange.Column + intLoop - 1).ColumnWidth
        Next intLoop
        oRange.MergeCells = False
        strLen = Len(.Cells(oRange.Row, oRange.Column).Value)
        oldZZWidth = .Range("XFC1048575").ColumnWidth
        .Range("XFC1048575") = Left(.Cells(oRange.Row, oRange.Column).Value, strLen)
        .Range("XFC1048575").WrapText = True
        .Columns("XFC").ColumnWidth = oldWidth
        .Rows("1048575").EntireRow.AutoFit
        newHeight = .Rows("1048575").RowHeight / oRange.Rows.Count
        .Rows(CStr(oRange.Row) & ":" & CStr(oRange.Row + oRange.Rows.Count - 1)).RowHeight = newHeight
        oRange.MergeCells = True
        oRange.WrapText = True
        .Rows("1048575").Delete
    End With

    Application.ScreenUpdating = True

End Sub

'------------------------------
'Remove VBA Password
'------------------------------
Sub removeVBAPassword()
On Error GoTo Err_Pro
    Dim Filename
    Dim i As Long

    Filename = Application.GetOpenFilename("Excel file(*.xls & *.xla & *.xlt), *.xls;*.xla;*.xlt", , "VBA")

    If Dir(Filename) = "" Then
        MsgBox "Not Find the file"
        Exit Sub
    End If

    Dim GetData As String * 5
    Open Filename For Binary As #1
    Dim CMGs As Long
    Dim DPBo As Long

    For i = 1 To LOF(1)
        Get #1, i, GetData
        If GetData = "CMG=""" Then CMGs = i
        If GetData = "[Host" Then DPBo = i - 2: Exit For
    Next

    If CMGs = 0 Then
        MsgBox "No password.", 32, "msg"
        Exit Sub
    End If

    Dim St As String * 2
    Dim s20 As String * 1

    '0D0A string
    Get #1, CMGs - 2, St

    'Hex string
    Get #1, DPBo + 16, s20

    For i = CMGs To DPBo Step 2
        Put #1, i, St
    Next

    If (DPBo - CMGs) Mod 2 <> 0 Then
        Put #1, DPBo + 1, s20
    End If
    MsgBox "OK!", 32, "msg"

    Close #1
    Exit Sub
Err_Pro:

End Sub

