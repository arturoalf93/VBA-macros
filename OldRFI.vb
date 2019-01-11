Attribute VB_Name = "OldRFI"

Function Col_Letter(lngCol As Integer) As String
'NOT USED. DIRECTLY USED Split(Cells(1, lngCol).Address(True, False), "$")(0) in the code
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
Sub UnprotectAll()
Attribute UnprotectAll.VB_ProcData.VB_Invoke_Func = "u\n14"

Dim pw As String
Dim ws As Worksheet

pw = "****"

For Each ws In Worksheets
    If ws.Name = "P2P" Or ws.Name = "Sourcing" Or ws.Name = "Spend Analytics" Or ws.Name = "SXM" Or _
    ws.Name = "CLM" Or ws.Name = "Common Requirements" Or ws.Name = "Temp Staffing" Or ws.Name = "Contracted Services SOW" Or ws.Name = "Independent Contract Workers" Then
        ws.Unprotect Password:=pw
    End If
Next

End Sub


Sub ProtectAll()
Attribute ProtectAll.VB_ProcData.VB_Invoke_Func = "p\n14"

Dim pw As String
Dim ws As Worksheet

pw = "****"

For Each ws In Worksheets
    If ws.Name = "P2P" Or ws.Name = "Sourcing" Or ws.Name = "Spend Analytics" Or ws.Name = "SXM" Or _
    ws.Name = "CLM" Or ws.Name = "Common Requirements" Or ws.Name = "Temp Staffing" Or ws.Name = "Contracted Services SOW" Or ws.Name = "Independent Contract Workers" Then
        ws.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    End If
Next
 
End Sub
Function CurrentScoreRow() As Integer
Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 3 To 50      'number of rows downwards
    For j = 1 To 500        'number of columns to the right
        If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Current score" Or Cells(i, j).Value = "Current Score" Then
                CurrentScoreRow = i
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 501 Then
    msgbox "Current Score not found"
End If
End Function
Function CurrentScoreColumn() As Integer
Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 1 To 50      'number of rows downwards
    For j = 1 To 500        'number of columns to the right
        If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Current score" Then
                CurrentScoreColumn = j
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 501 Then
    msgbox "Current Score not found"
End If
End Function


Sub CreateSelfScoreColumn()
Attribute CreateSelfScoreColumn.VB_ProcData.VB_Invoke_Func = "s\n14"
'Creates the Self-Score column in the left of the Current Score one

Application.ScreenUpdating = False

Dim i As Integer, j As Integer, k As Integer, flag As Integer, HeaderRow As Integer, CsColumn As Integer, NumberOfSS As Integer
Dim SScolumnsNumbers(99) As Integer
Dim SScolumnsLetters(99) As String
Dim AuxArr
Dim Formula As String, CsColumnLetter As String

HeaderRow = CurrentScoreRow()
CsColumn = CurrentScoreColumn()
CsColumnLetter = Split(Cells(1, CsColumn).Address(True, False), "$")(0)

'Insert column in the left of Current score
Columns(CsColumnLetter & ":" & CsColumnLetter).Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.Clear

'Update Currentscore location
CsColumn = CsColumn + 1
CsColumnLetter = Split(Cells(1, CsColumn).Address(True, False), "$")(0)

'Count how many Self-Scores there are and where they are
k = 0
For j = 1 To CsColumn
    If Cells(HeaderRow, j) = "Self-score" Or Cells(HeaderRow, j) = "Self-Score" Or Cells(HeaderRow, j) = "Self-score (2)" Or Cells(HeaderRow, j) = "Self-Score (2)" Then
        SScolumnsNumbers(k) = j
        SScolumnsLetters(k) = Split(Cells(1, j).Address(True, False), "$")(0) ' Convert column number to column
        k = k + 1
    End If
Next j
NumberOfSS = k
If NumberOfSS = 0 Then
    msgbox "No Self-Scores found"
    Exit Sub
End If

Cells(HeaderRow, CsColumn - 1).Value = "Current Self-Score"
Cells(HeaderRow, CsColumn - 1).Select
Selection.Font.Bold = True
    With Selection.Font
        .Size = 14
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

For i = 1 To 1000
    If Cells(HeaderRow + i, 1).Value <> "" Then
        Formula = GetFormula(HeaderRow + i, NumberOfSS, SScolumnsLetters)
        Cells(HeaderRow + i, CsColumn - 1).Value = Formula
        Cells(HeaderRow + i, CsColumn - 1).Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        
    End If
Next i

Application.ScreenUpdating = True

End Sub


Function GetFormula(ByVal row As Integer, ByVal NumberOfColumns As Integer, ByRef arr() As String) As String
Dim j As Integer

GetFormula = "="
For j = NumberOfColumns - 1 To 0 Step -1
    GetFormula = GetFormula & "IF(" & arr(j) & row & "<>" & Chr(34) & Chr(34) & "," & arr(j) & row & ","
Next j
GetFormula = GetFormula & Chr(34) & Chr(34)
For j = 1 To NumberOfColumns
    GetFormula = GetFormula & ")"
Next j

End Function

Sub CheckandCorrectSS0to5()
Attribute CheckandCorrectSS0to5.VB_ProcData.VB_Invoke_Func = "k\n14"

Dim i As Integer, j As Integer, k As Integer, flag As Integer, HeaderRow As Integer, CsColumn As Integer, NumberOfSS As Integer
Dim SScolumnsNumbers(99) As Integer
Dim SScolumnsLetters(99) As String
Dim AuxArr
Dim Formula As String, CsColumnLetter As String

HeaderRow = CurrentScoreRow()
CsColumn = CurrentScoreColumn()

'Count how many Self-Scores there are
k = 0
For j = 1 To CsColumn
    If Cells(HeaderRow, j) = "Self-score" Or Cells(HeaderRow, j) = "Self-Score" Or Cells(HeaderRow, j) = "Self-score (2)" Or Cells(HeaderRow, j) = "Self-Score (2)" Then
        SScolumnsNumbers(k) = j
        SScolumnsLetters(k) = Split(Cells(1, j).Address(True, False), "$")(0) ' Convert column number to letter
        k = k + 1
    End If
Next j
NumberOfSS = k
If NumberOfSS = 0 Then
    msgbox "No Self-Scores found"
    Exit Sub
End If

'Get the first number of the self-score, or warn otherwise
For i = 1 To 1000
    If Cells(HeaderRow + i, 1).Value <> "" Then
        For k = 0 To NumberOfSS - 1
            If Cells(HeaderRow + i, SScolumnsNumbers(k)).Value <> "" Then
                If Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = 0 Or Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = 1 Or Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = 2 Or _
                Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = 3 Or Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = 4 Or Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = 5 Then
                    'do nothing. There is no need to write anything
                ElseIf Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1) = 0 Or Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1) = 1 _
                Or Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1) = 2 Or Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1) = 3 _
                Or Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1) = 4 Or Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1) = 5 Then
                        Cells(HeaderRow + i, SScolumnsNumbers(k)).Value = Left(Cells(HeaderRow + i, SScolumnsNumbers(k)).Value, 1)
                Else
                        msgbox "First letter of " & ActiveSheet.Name & ", row " & HeaderRow + i & ", column " & SScolumnsNumbers(k) & " different than 0, 1, 2, 3, 4 or 5"
                        Exit Sub
                End If
                
            End If
        Next k
    End If
Next i


End Sub

Sub AddColumns()
Attribute AddColumns.VB_ProcData.VB_Invoke_Func = "t\n14"

Dim i As Integer, j As Integer, k As Integer, HeaderRow As Integer, CsColumn As Integer, LeftBound As Integer, RightBound As Integer
Dim CSSColumnLetter As String, LeftBoundLetter As String, RightBoundLetter As String, aux As String
Dim ColumnWidths(9) As Integer
Dim ColumnLetters(9) As String

ColumnWidths(0) = 6
ColumnWidths(1) = 50
ColumnWidths(2) = 10
ColumnWidths(3) = 6
ColumnWidths(4) = 10
ColumnWidths(5) = 6
ColumnWidths(6) = 25
ColumnWidths(7) = 10
ColumnWidths(8) = 6
ColumnWidths(9) = 10
HeaderRow = CurrentScoreRow()
CsColumn = CurrentScoreColumn()
CSSColumnLetter = Split(Cells(1, CsColumn - 1).Address(True, False), "$")(0)

'Check that there is Current Self-Score column
If Cells(HeaderRow, CsColumn - 1) <> "Current Self-Score" Then
    msgbox "Current Self-Score column missing"
    Exit Sub
End If

'Insert 10 columns in the left of Current score
Columns(CSSColumnLetter & ":" & CSSColumnLetter).Select
For k = 1 To 10
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Next k

CsColumn = CsColumn + 10
CSSColumnLetter = Split(Cells(1, CsColumn - 1).Address(True, False), "$")(0)
LeftBound = CsColumn - 11
RightBound = CsColumn - 2
LeftBoundLetter = Split(Cells(1, LeftBound).Address(True, False), "$")(0)
RightBoundLetter = Split(Cells(1, RightBound).Address(True, False), "$")(0)
Columns(LeftBoundLetter & ":" & RightBoundLetter).Select
Selection.Clear

'Set columns width
For j = 0 To 9
    ColumnLetters(j) = Split(Cells(1, LeftBound + j).Address(True, False), "$")(0)
    Columns(ColumnLetters(j) & ":" & ColumnLetters(j)).ColumnWidth = ColumnWidths(j)
Next j

'Create Header
Cells(HeaderRow, LeftBound).Value = "Self-Score"
Cells(HeaderRow, LeftBound + 1).Value = "Self-Description"
Cells(HeaderRow, LeftBound + 2).Value = "Attachments/Supporting Docs and Location/Link"
Cells(HeaderRow, LeftBound + 3).Value = "SM score"
Cells(HeaderRow, LeftBound + 4).Value = "Analyst notes"
Cells(HeaderRow, LeftBound + 5).Value = "Self-Score (2)"
Cells(HeaderRow, LeftBound + 6).Value = "Reasoning"
Cells(HeaderRow, LeftBound + 7).Value = "Attachments/Supporting Docs and Location/Link"
Cells(HeaderRow, LeftBound + 8).Value = "SM score (2)"
Cells(HeaderRow, LeftBound + 9).Value = "Analyst notes (2)"

Range(Cells(HeaderRow, LeftBound), Cells(HeaderRow, RightBound)).Select
Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
Union(Range(Cells(HeaderRow, LeftBound), Cells(HeaderRow, LeftBound + 2)), Range(Cells(HeaderRow, LeftBound + 5), Cells(HeaderRow, LeftBound + 7))).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
End With
Union(Range(Cells(HeaderRow, LeftBound + 3), Cells(HeaderRow, LeftBound + 4)), Range(Cells(HeaderRow, LeftBound + 8), Cells(HeaderRow, LeftBound + 9))).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 49407
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Format Cells
For i = HeaderRow + 1 To 1000
        If Cells(i, 1).Value <> "" Then
                Range(Cells(i, LeftBound), Cells(i, RightBound)).Select
                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .Orientation = 0
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                With Selection.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                
                Union(Range(Cells(i, LeftBound), Cells(i, LeftBound + 2)), Range(Cells(i, LeftBound + 5), Cells(i, LeftBound + 7))).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent5
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
                
                Union(Cells(i, LeftBound), Cells(i, LeftBound + 5)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                End With
                    With Selection.Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0,1,2,3,4,5"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = "Value must be 0, 1, 2, 3, 4 or 5"
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With
                Union(Cells(i, LeftBound + 3), Cells(i, LeftBound + 8)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                End With
                With Selection.Validation
                    .Delete
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="5"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = "Value must be between 0 and 5"
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With
        End If
Next i

'Update Current Self-score and Current SM score columns formulas

End Sub

Sub UpdateCSSandCS()
Attribute UpdateCSSandCS.VB_ProcData.VB_Invoke_Func = "y\n14"

Dim i As Integer, j As Integer, k As Integer, flag As Integer, HeaderRow As Integer, CsColumn As Integer, NumberOfSS As Integer, NumberOfCS As Integer
Dim SScolumnsNumbers(99) As Integer, CScolumnsNumbers(99) As Integer
Dim SScolumnsLetters(99) As String
Dim CScolumnsLetters(99) As String
Dim Formula As String

HeaderRow = CurrentScoreRow()
CsColumn = CurrentScoreColumn()

'Check that there is Current Self-Score column
If Cells(HeaderRow, CsColumn - 1) <> "Current Self-Score" Then
    msgbox "Current Self-Score column missing"
    Exit Sub
End If

'Count how many Self-Scores there are and where they are
k = 0
For j = 1 To CsColumn
    If Cells(HeaderRow, j) = "Self-score" Or Cells(HeaderRow, j) = "Self-Score" Or Cells(HeaderRow, j) = "Self-score (2)" Or Cells(HeaderRow, j) = "Self-Score (2)" Then
        SScolumnsNumbers(k) = j
        SScolumnsLetters(k) = Split(Cells(1, j).Address(True, False), "$")(0) ' Convert column number to letter
        k = k + 1
    End If
Next j
NumberOfSS = k
If NumberOfSS = 0 Then
    msgbox "No Self-Scores found"
    Exit Sub
End If

'Count how many SM scores there are and where they are
k = 0
For j = 1 To CsColumn
    If Cells(HeaderRow, j) = "SM score" Or Cells(HeaderRow, j) = "SM Score" Or Cells(HeaderRow, j) = "SM score (2)" Or Cells(HeaderRow, j) = "SM score (2)" Then
        CScolumnsNumbers(k) = j
        CScolumnsLetters(k) = Split(Cells(1, j).Address(True, False), "$")(0) ' Convert column number to letter
        k = k + 1
        
    End If
Next j
NumberOfCS = k
If NumberOfCS = 0 Then
    msgbox "No SM scores found"
    Exit Sub
End If

'Write SS and CS formulas
For i = 1 To 1000
    If Cells(HeaderRow + i, 1).Value <> "" Then
        Formula = GetFormula(HeaderRow + i, NumberOfSS, SScolumnsLetters)
        Cells(HeaderRow + i, CsColumn - 1).Value = Formula
        
        Formula = GetFormula(HeaderRow + i, NumberOfCS, CScolumnsLetters)
        Cells(HeaderRow + i, CsColumn).Value = Formula
    End If
Next i

End Sub

Sub DeleteBlueBackground()
Attribute DeleteBlueBackground.VB_ProcData.VB_Invoke_Func = "e\n14"

    Dim i As Integer, j As Integer, k As Integer, HeaderRow As Integer, CsColumn As Integer
    
    HeaderRow = CurrentScoreRow()
    CsColumn = CurrentScoreColumn()
    
    'Check that there is Current Self-Score column
    If Cells(HeaderRow, CsColumn - 1) <> "Current Self-Score" Then
        msgbox "Current Self-Score column missing"
        Exit Sub
    End If
    
    'Check Self-Score is 11 places in the left of Current score
    If Cells(HeaderRow, CsColumn - 11) <> "Self-Score" Then
        msgbox "Self-Score column is not in the right place"
        Exit Sub
    End If
    
    Range(Cells(HeaderRow + 1, CsColumn - 11), Cells(HeaderRow + 1000, CsColumn - 2)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Sub LockCells()
Attribute LockCells.VB_ProcData.VB_Invoke_Func = "j\n14"

Dim i As Integer, j As Integer, k As Integer, HeaderRow As Integer, CsColumn As Integer, CCrow As Integer, CCcolumn As Integer
Dim mode As String, CCrange As String

HeaderRow = CurrentScoreRow()
CsColumn = CurrentScoreColumn()
CCrow = GetCCrow()
CCcolumn = GetCCcolumn()
    
'Check that there is Current Self-Score column
If Cells(HeaderRow, CsColumn - 1) <> "Current Self-Score" Then
    msgbox "Current Self-Score column missing"
    Exit Sub
End If

'Check Self-Score is 11 places in the left of Current score
If Cells(HeaderRow, CsColumn - 11) <> "Self-Score" Then
    msgbox "Self-Score column is not in the right place"
    Exit Sub
End If

'Find CC location and determine if it's only one (SPT, CWS) or it's three categories
If Cells(CCrow, CCcolumn).Value = "Customer count for each category (bubble size)" And ActiveSheet.Name = "P2P" Then
    mode = "P2P"
ElseIf Cells(CCrow, CCcolumn).Value = "Customer count (bubble size)" Then
    mode = "SPTorCWS"
Else
    msgbox "Check Customer Count cell"
    Exit Sub
End If

'Unlock all cells
Cells.Select
Selection.Locked = True
Selection.FormulaHidden = False

'Lock specific cells
If mode = "SPTorCWS" Then
    For j = 1 To 5
        If (Cells(CCrow + 1, CCcolumn + j).Value = "" Or Cells(CCrow + 1, CCcolumn + j).Value = "-") And Application.WorksheetFunction.IsNumber(Cells(CCrow + 1, CCcolumn + j).Value) = False Then
            CCrange = Split(Cells(1, CCcolumn + j).Address(True, False), "$")(0) & CCrow + 1
            Exit For
        End If
        If j = 6 Then
            msgbox "Check Customer Count cell"
            Exit Sub
        End If
    Next j
ElseIf mode = "P2P" Then
    For j = 1 To 5
        If ((Cells(CCrow + 1, CCcolumn + j).Value = "" And Cells(CCrow + 2, CCcolumn + j).Value = "" And Cells(CCrow + 3, CCcolumn + j).Value = "") Or _
        (Cells(CCrow + 1, CCcolumn + j).Value = "-" And Cells(CCrow + 2, CCcolumn + j).Value = "-" And Cells(CCrow + 3, CCcolumn + j).Value = "-")) And _
        (Application.WorksheetFunction.IsNumber(Cells(CCrow + 1, CCcolumn + j).Value) = False And _
        Application.WorksheetFunction.IsNumber(Cells(CCrow + 2, CCcolumn + j).Value) = False And _
        Application.WorksheetFunction.IsNumber(Cells(CCrow + 2, CCcolumn + j).Value) = False) Then
                CCrange = Split(Cells(1, CCcolumn + j).Address(True, False), "$")(0) & CCrow + 1 & ":" & Split(Cells(1, CCcolumn + j).Address(True, False), "$")(0) & CCrow + 3
                Exit For
        End If
        If j = 6 Then
                msgbox "Check Customer Count cell"
                Exit Sub
        End If
    Next j
Else
    msgbox "mode neither SPTorCWS or P2P??"
    Exit Sub
End If

Union(Range(CCrange), Range(Cells(HeaderRow + 1, CsColumn - 11), Cells(HeaderRow + 1000, CsColumn - 2))).Select
Selection.Locked = False
Selection.FormulaHidden = False

End Sub

Function GetCCrow() As Integer
'Customer Count Header Row

Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 1 To 50      'number of rows downwards
    For j = 1 To 100        'number of columns to the right
        If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Customer count (bubble size)" Or Cells(i, j).Value = "Customer count for each category (bubble size)" Then
                GetCCrow = i
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 101 Then
    msgbox "Customer count not found"
End If
End Function
Function GetCCcolumn() As Integer
'Customer Count Header Column

Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 1 To 50      'number of rows downwards
    For j = 1 To 100        'number of columns to the right
        If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Customer count (bubble size)" Or Cells(i, j).Value = "Customer count for each category (bubble size)" Then
                GetCCcolumn = j
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 101 Then
    msgbox "Customer count not found"
End If
End Function

Function GetPArow() As Integer
'Provider Average Row

Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 1 To 50      'number of rows downwards
    For j = 1 To 100        'number of columns to the right
        If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Provider Average" Then
                GetPArow = i
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 101 Then
    msgbox "Provider Average not found"
End If
End Function
Function GetPAcolumn() As Integer
'Provider Average Row

Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 1 To 50      'number of rows downwards
    For j = 1 To 100        'number of columns to the right
       If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Provider Average" Then
                GetPAcolumn = j
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 101 Then
    msgbox "Provider Average not found"
End If
End Function

Sub SummaryColumns()
Attribute SummaryColumns.VB_ProcData.VB_Invoke_Func = "g\n14"

Dim i As Integer, j As Integer, k As Integer, flag As Integer, HeaderRow As Integer, CsColumn As Integer, NumberOfSS As Integer, PArow As Integer, PAcolumn As Integer, NumberOfRows As Integer
Dim SScolumnsNumbers(99) As Integer
Dim SScolumnsLetters(99) As String
Dim AuxArr
Dim Formula As String, CsColumnLetter As String

If ActiveSheet.Name = "P2P" Then
    NumberOfRows = 13
ElseIf ActiveSheet.Name = "Sourcing" Then
    NumberOfRows = 12
ElseIf ActiveSheet.Name = "Spend Analytics" Then
    NumberOfRows = 7
ElseIf ActiveSheet.Name = "SXM" Then
    NumberOfRows = 8
ElseIf ActiveSheet.Name = "CLM" Then
    NumberOfRows = 7
ElseIf ActiveSheet.Name = "Temp Staffing" Then
    NumberOfRows = 20
ElseIf ActiveSheet.Name = "Contracted Services SOW" Then
    NumberOfRows = 17
ElseIf ActiveSheet.Name = "Independent Contract Workers" Then
    NumberOfRows = 22
Else
    msgbox "Something wrong in the Sheet name"
End If

'Find "Provider Average" location
PArow = GetPArow()
PAcolumn = GetPAcolumn()

'Insert 3 columns in the left of Provider Average
For k = 1 To 3
    Range(Cells(PArow, PAcolumn), Cells(PArow + NumberOfRows, PAcolumn)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Next k

Cells(PArow, PAcolumn + 3).Value = "Current Provider Average"

'Move Benchmark Average to the left
Range(Cells(PArow, PAcolumn + 4), Cells(PArow + NumberOfRows, PAcolumn + 4)).Select
Selection.Cut
Cells(PArow, PAcolumn).Select
ActiveSheet.Paste
Cells(PArow, PAcolumn).Value = "Last Quarter Benchmark Average"
For i = 1 To NumberOfRows
    Cells(PArow + i, PAcolumn) = "-"
Next i

'Copy values and format of Current Provider Average 2 columns to the left
Range(Cells(PArow, PAcolumn + 3), Cells(PArow + NumberOfRows, PAcolumn + 3)).Select
Selection.Copy
Cells(PArow, PAcolumn + 1).Select
Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Cells(PArow, PAcolumn + 1).Value = "Last Quarter Provider Average"

'Copy formulas from Current Provider Average one to the left
Range(Cells(PArow, PAcolumn + 3), Cells(PArow + NumberOfRows, PAcolumn + 3)).Select
Selection.Copy
Cells(PArow, PAcolumn + 2).Select
ActiveSheet.Paste

'Format Current SS column
Cells(PArow, PAcolumn + 2).Value = "Current Self-Score Average"
Cells(PArow, PAcolumn + 2).Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColor = 0
    .ThemeColor = xlThemeColorAccent2
    .TintAndShade = 0.599993896298105
    .PatternTintAndShade = 0
End With

'All with borders
Range(Cells(PArow, PAcolumn - 1), Cells(PArow + NumberOfRows, PAcolumn + 3)).Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With


msgbox "Check global average is the average of all the elements, not of all the subcategory averages"
End Sub


Sub CheckEmptyColumns()
Attribute CheckEmptyColumns.VB_ProcData.VB_Invoke_Func = "r\n14"

Dim i As Integer, j As Integer, k As Integer, HeaderRow As Integer, CsColumn As Integer, FirstSScolumn As Integer

HeaderRow = CurrentScoreRow()
CsColumn = CurrentScoreColumn()
FirstSScolumn = GetFirstSScolumn()

For j = FirstSScolumn To CsColumn
        For i = HeaderRow + 1 To HeaderRow + 1000
                If Cells(i, j).Value <> "" Then
                    Exit For
                End If
        If i = HeaderRow + 1000 Then
                Cells(HeaderRow - 1, j).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
        End If
        Next i
Next j
End Sub

Function GetFirstSScolumn() As Integer

Dim i As Integer, j As Integer, flag As Integer
flag = 0

For i = 1 To 50      'number of rows downwards
    For j = 1 To 100        'number of columns to the right
        If IsError(Cells(i, j).Value) <> True Then
            If Cells(i, j).Value = "Self-score" Or Cells(i, j).Value = "Self-Score" Then
                GetFirstSScolumn = j
                flag = 1
                Exit For
            End If
        End If
    Next j
    If flag = 1 Then Exit For
Next i
If i = 51 And j = 101 Then
    msgbox "First Self-score not found"
End If

End Function

Sub CSV()

Application.ScreenUpdating = False

Dim i As Integer, j As Integer, k As Integer, l As Integer, flag As Integer, HeaderRow As Integer, CsColumn As Integer, FirstSScolumn As Integer, _
NumberOfCategorySheets As Integer, NumberOfQuarters As Integer, NumberOfSMscores As Integer, DataArrayCount As Integer, NumberOfQuarterOrRoundChanges As Integer, _
CurrentCSVrow As Integer
Dim scseID As Integer, quarter As Integer, Year As Integer, Round As Integer, SelfScore As Integer
Dim SMscore As Single
Dim SelfDescription As String, AnalystNotes As String
Dim QuarterColumns(9) As Integer
Dim SheetsNames(99) As String
Dim quarters(9) As String
Dim SMscorePositions(9) As Integer
Dim ws As Worksheet
Dim PositionsColumns(499) As Integer
Dim PositionsQuarters(499) As String
Dim PositionsRounds(499) As Integer
Dim PositionsTypes(499) As Integer 'Where type = 0 if SS, 1 if SD..... 8 if SM (2), 9 if AN (2).
Dim NumberOfColumns As Integer 'Number of columns from the FirstSS to CurrentSS
Dim QuarterOrRoundChanges(49) As Integer
Dim DataArray(4999, 8) As String

'If there is a csv sheet already, delete it and create a new one
'Get Category Sheet names
'Check there is not any not recognized name
k = 0
For Each ws In Worksheets
    If ws.Name = "P2P" Or ws.Name = "Sourcing" Or ws.Name = "Spend Analytics" Or ws.Name = "SXM" Or _
    ws.Name = "CLM" Or ws.Name = "Common Requirements" Or ws.Name = "Temp Staffing" Or ws.Name = "Contracted Services SOW" Or ws.Name = "Independent Contract Workers" Then
        SheetsNames(k) = ws.Name
        k = k + 1
    ElseIf ws.Name = "csv" Then
        Application.DisplayAlerts = False
        Sheets("csv").Select
        ActiveWindow.SelectedSheets.Delete
        Application.DisplayAlerts = True
    ElseIf ws.Name <> "Instructions" And ws.Name <> "Company Information" Then
        msgbox "Sheet name not recognized: " & ws.Name
        Exit Sub
    End If
Next
NumberOfCategorySheets = k

'Create csv sheet
Sheets.Add.Name = "csv"
Worksheets("csv").Move After:=Worksheets(SheetsNames(NumberOfCategorySheets - 1))
Sheets("csv").Select
Cells(1, 1).Value = "VendorID"
Cells(1, 2).Value = "scseID"
Cells(1, 3).Value = "Quarter"
Cells(1, 4).Value = "Year"
Cells(1, 5).Value = "Round"
Cells(1, 6).Value = "UpdateDate"
Cells(1, 7).Value = "Self-score"
Cells(1, 8).Value = "Self-Description_Or_Reasoning"
Cells(1, 9).Value = "AttachmentID"
Cells(1, 10).Value = "SMscore"
Cells(1, 11).Value = "Analyst notes"
Cells(1, 12).Value = "User_ID"


For k = 0 To NumberOfCategorySheets - 1

        Sheets(SheetsNames(k)).Select
        
        'Check first column of CLM doesn't have the old IDs
        If ActiveSheet.Name = "CLM" Then
                If Cells(19, 1).Value = "old scseID" Then
                        Columns(1).Delete
                        Cells(19.1).Value = "scseID"
                ElseIf Cells(19, 1).Value = "new scseID" Or Cells(19, 1).Value = "scseID" Then
                        'do nothing
                Else
                        msgbox "Something weird with the old and new scseIDs columns"
                End If
        End If
        
        HeaderRow = CurrentScoreRow()
        CsColumn = CurrentScoreColumn()
        FirstSScolumn = GetFirstSScolumn()
        
        NumberOfQuarters = 0
        For i = 0 To 9
            QuarterColumns(i) = 0
            quarters(i) = ""
            SMscorePositions(i) = 0
        Next i
        
        'Check there is Current Self-Description column
        If Cells(HeaderRow, CsColumn - 1) <> "Current Self-Score" Then
            msgbox "No Current Self-Description column found"
            Exit Sub
        End If
        
        l = 0
        flag = 0
        'Map quarter columns and check they have the proper format: Q2 17 or Q4 17 or Q1 18 or Q2 18 or Q3 18 or Q4 18
        For j = FirstSScolumn To CsColumn - 2
                If Cells(HeaderRow - 1, j).Value = "" And l = 0 Then
                        msgbox "Error, first quarter cell emprty"
                        Exit Sub
                ElseIf Cells(HeaderRow - 1, j).Value = "" Then
                        'Do nothing
                ElseIf (Cells(HeaderRow - 1, j).Value = "Q2 17" Or Cells(HeaderRow - 1, j).Value = "Q4 17" Or Cells(HeaderRow - 1, j).Value = "Q1 18" Or _
                Cells(HeaderRow - 1, j).Value = "Q2 18" Or Cells(HeaderRow - 1, j).Value = "Q3 18" Or Cells(HeaderRow - 1, j).Value = "Q4 18") Then
                        If l = 0 Then
                                QuarterColumns(l) = j
                                quarters(l) = Cells(HeaderRow - 1, j)
                                l = l + 1
                        Else
                                If quarters(l - 1) = Cells(HeaderRow - 1, j) Then
                                        flag = flag + 1
                                        If flag = 2 Then
                                                msgbox "Same quarter three times in arow?"
                                                Exit Sub
                                        End If
                                Else
                                        QuarterColumns(l) = j
                                        quarters(l) = Cells(HeaderRow - 1, j)
                                        l = l + 1
                                        flag = 0
                                End If
                        End If
                        
                Else
                        msgbox "Something wrong in the Quarters row"
                        Exit Sub
                End If
        Next j
        NumberOfQuarters = l
        
        'Map "SM score" cells, that separate round 1 from 2
        l = 0
        For j = FirstSScolumn To CsColumn - 2
                If Cells(HeaderRow, j).Value = "SM score" Then
                        SMscorePositions(l) = j
                        l = l + 1
                End If
        Next j
        NumberOfSMscores = l
        
        
        'Check Quarters() and SMscorePositions() make sense
        If NumberOfQuarters <> NumberOfSMscores Then
                msgbox "Different NumberOfQuarters than NumberOfSMscores"
                Exit Sub
        End If
        For i = 0 To NumberOfQuarters - 2
                If SMscorePositions(i) >= QuarterColumns(i) And SMscorePositions(i) < QuarterColumns(i + 1) Then
                        'Do nothing
                Else
                        msgbox "Error in Quarters and SMscorePositions"
                        Exit Sub
                End If
        Next i
        If SMscorePositions(NumberOfSMscores - 1) <= QuarterColumns(NumberOfQuarters - 1) Then
                msgbox "Error in last Quarter or SMscore position"
        End If
        
        'Create Position vectors (clear it first)
        For i = 0 To 499
                PositionsColumns(i) = -1
                PositionsQuarters(i) = ""
                PositionsRounds(i) = -1
                PositionsTypes(i) = -1
        Next i
        
        NumberOfColumns = CsColumn - FirstSScolumn - 1
        
        'Find quarters
        For j = 0 To NumberOfColumns - 1
                PositionsColumns(j) = QuarterColumns(0) + j
                For l = NumberOfQuarters - 1 To 0 Step -1
                        If PositionsColumns(j) >= QuarterColumns(l) Then
                                PositionsQuarters(j) = quarters(l)
                                Exit For
                        End If
                Next l
        Next j

        'Find type of column and round
        For j = 0 To NumberOfColumns - 1
                If Cells(HeaderRow, PositionsColumns(j)).Value = "Self-score" Or Cells(HeaderRow, PositionsColumns(j)).Value = "Self-Score" Then
                        PositionsRounds(j) = 1
                        PositionsTypes(j) = 0
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "Self-description" Or Cells(HeaderRow, PositionsColumns(j)).Value = "Self-Description" Or _
                Cells(HeaderRow, PositionsColumns(j)).Value = "Self -Description" Then
                        PositionsRounds(j) = 1
                        PositionsTypes(j) = 1
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "Attachments/Supporting Docs and Location/Link" Then
                        If PositionsTypes(j - 1) = 1 And PositionsRounds(j - 1) = 1 Then
                                PositionsRounds(j) = 1
                                PositionsTypes(j) = 2
                        ElseIf PositionsTypes(j - 1) = 1 And PositionsRounds(j - 1) = 2 Then
                                PositionsRounds(j) = 2
                                PositionsTypes(j) = 2
                        Else
                                msgbox "Something weird going on with Attachement column"
                                Exit Sub
                        End If
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "SM score" Or Cells(HeaderRow, PositionsColumns(j)).Value = "SM Score" Then
                        PositionsRounds(j) = 1
                        PositionsTypes(j) = 3
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "Analyst notes" Then
                        PositionsRounds(j) = 1
                        PositionsTypes(j) = 4
                 ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "Self-score (2)" Or Cells(HeaderRow, PositionsColumns(j)).Value = "Self-Score (2)" Then
                        PositionsRounds(j) = 2
                        PositionsTypes(j) = 0
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "Reasoning" Then
                        PositionsRounds(j) = 2
                        PositionsTypes(j) = 1
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "SM score (2)" Or Cells(HeaderRow, PositionsColumns(j)).Value = "SM Score (2)" Then
                        PositionsRounds(j) = 2
                        PositionsTypes(j) = 3
                ElseIf Cells(HeaderRow, PositionsColumns(j)).Value = "Analyst notes (2)" Then
                        PositionsRounds(j) = 2
                        PositionsTypes(j) = 4
                Else
                        msgbox "Unrecognized header: " & Cells(HeaderRow, PositionsColumns(j)).Value
                        Exit Sub
                End If
        Next j
        
       
        'Clear QuarterOrRoundChanges first
        For i = 0 To 49
                QuarterOrRoundChanges(i) = -1
        Next i
         'Fill QuarterOrRoundChanges
        l = 0
        For i = 0 To NumberOfColumns - 2
                If PositionsQuarters(i) <> PositionsQuarters(i + 1) Or PositionsRounds(i) <> PositionsRounds(i + 1) Then
                        QuarterOrRoundChanges(l) = i
                        l = l + 1
                End If
        Next i
        l = 0
        
       
        'Clear DataArray
        For i = 0 To 4999
            For j = 0 To 8
                    DataArray(i, j) = ""
            Next j
        Next i
        
         'Fill DataArray
        DataArrayCount = 0
         'First Quarter and Round out of the main loop
        For i = 0 To 499
                If Cells(HeaderRow + 1 + i, 1).Value <> "" Then
                        DataArray(DataArrayCount, 0) = Cells(HeaderRow + 1 + i, 1).Value
                        DataArray(DataArrayCount, 1) = Mid(PositionsQuarters(0), 2, 1)
                        DataArray(DataArrayCount, 2) = "20" & Mid(PositionsQuarters(0), 4, 2)
                        DataArray(DataArrayCount, 3) = PositionsRounds(0)
                        
                        For j = 0 To QuarterOrRoundChanges(0)
                                If PositionsTypes(j) <> 2 Then 'WE DON'T DO ATTACHMENTS BY NOW
                                        DataArray(DataArrayCount, 4 + PositionsTypes(j)) = Cells(HeaderRow + 1 + i, PositionsColumns(j)).Value
                                End If
                                
                                'Check that Quarter and Round is the same
                                If PositionsQuarters(j) <> PositionsQuarters(0) Or PositionsRounds(j) <> PositionsRounds(0) Then
                                        msgbox "Quarter or Round is not the same where it should be"
                                        Exit Sub
                                End If
                        Next j
                        DataArrayCount = DataArrayCount + 1
                End If
        Next i
        
        'Count last valid position in QuarterOrRoundChanges : NumberOfQuarterOrRoundChanges
        For i = 0 To 49
                If QuarterOrRoundChanges(i) = -1 Then
                        NumberOfQuarterOrRoundChanges = i - 1
                        Exit For
                End If
        Next i
    
        'Main loop
        For i = 0 To 499
                If Cells(HeaderRow + 1 + i, 1).Value <> "" Then
                        For l = 0 To NumberOfQuarterOrRoundChanges - 1
                                
                                DataArray(DataArrayCount, 0) = Cells(HeaderRow + 1 + i, 1).Value
                                DataArray(DataArrayCount, 1) = Mid(PositionsQuarters(QuarterOrRoundChanges(l) + 1), 2, 1)
                                DataArray(DataArrayCount, 2) = "20" & Mid(PositionsQuarters(QuarterOrRoundChanges(l) + 1), 4, 2)
                                DataArray(DataArrayCount, 3) = PositionsRounds(QuarterOrRoundChanges(l) + 1)
                                        
                                For j = QuarterOrRoundChanges(l) + 1 To QuarterOrRoundChanges(l + 1)
                                        If PositionsTypes(j) <> 2 Then 'WE DON'T DO ATTACHMENTS BY NOW
                                                DataArray(DataArrayCount, 4 + PositionsTypes(j)) = Cells(HeaderRow + 1 + i, PositionsColumns(j)).Value
                                        End If
                                        
                                        'Check that Quarter and Round is the same
                                         If PositionsQuarters(j) <> PositionsQuarters(QuarterOrRoundChanges(l) + 1) Or PositionsRounds(j) <> PositionsRounds(QuarterOrRoundChanges(l) + 1) Then
                                        msgbox "Quarter or Round is not the same where it should be"
                                        Exit Sub
                                End If
                                Next j
                                DataArrayCount = DataArrayCount + 1
                        Next l
                End If
        Next i
        
        'Last quarter out of the main loop as well
        
        'First check, one more time, that there is a "Current Self-Description" column
        If Cells(HeaderRow, CsColumn - 1) <> "Current Self-Score" Then
                msgbox "Current Self-Score column missing"
                Exit Sub
        End If

        For i = 0 To 499
                If Cells(HeaderRow + 1 + i, 1).Value <> "" Then
                        DataArray(DataArrayCount, 0) = Cells(HeaderRow + 1 + i, 1).Value
                        DataArray(DataArrayCount, 1) = Mid(PositionsQuarters(QuarterOrRoundChanges(NumberOfQuarterOrRoundChanges) + 1), 2, 1)
                        DataArray(DataArrayCount, 2) = "20" & Mid(PositionsQuarters(QuarterOrRoundChanges(NumberOfQuarterOrRoundChanges) + 1), 4, 2)
                        DataArray(DataArrayCount, 3) = PositionsRounds(QuarterOrRoundChanges(NumberOfQuarterOrRoundChanges) + 1)
                        
                        For j = QuarterOrRoundChanges(NumberOfQuarterOrRoundChanges) + 1 To NumberOfColumns - 1
                                If PositionsTypes(j) <> 2 Then 'WE DON'T DO ATTACHMENTS BY NOW
                                        DataArray(DataArrayCount, 4 + PositionsTypes(j)) = Cells(HeaderRow + 1 + i, PositionsColumns(j)).Value
                                End If
                                
                                'Check that Quarter and Round is the same
                                If PositionsQuarters(j) <> PositionsQuarters(QuarterOrRoundChanges(NumberOfQuarterOrRoundChanges) + 1) Or _
                                PositionsRounds(j) <> PositionsRounds(QuarterOrRoundChanges(NumberOfQuarterOrRoundChanges) + 1) Then
                                        msgbox "Quarter or Round is not the same where it should be"
                                        Exit Sub
                                End If
                        Next j
                        DataArrayCount = DataArrayCount + 1
                End If
        Next i
        
        Sheets("csv").Select
        For i = 0 To DataArrayCount
                If DataArray(i, 4) <> "" Or DataArray(i, 5) <> "" Or DataArray(i, 6) <> "" Or DataArray(i, 7) <> "" Or DataArray(i, 8) <> "" Then
                        CurrentCSVrow = Application.CountA(Range("B:B")) + 1
                        Cells(CurrentCSVrow, 2).Value = DataArray(i, 0)
                        Cells(CurrentCSVrow, 3).Value = DataArray(i, 1)
                        Cells(CurrentCSVrow, 4).Value = DataArray(i, 2)
                        Cells(CurrentCSVrow, 5).Value = DataArray(i, 3)
                        Cells(CurrentCSVrow, 7).Value = DataArray(i, 4)
                        Cells(CurrentCSVrow, 8).Value = DataArray(i, 5)
                        'Cells(CurrentCSVrow, 9).Value = DataArray(i, 6) WE SKIP ATTACHMENTS BY NOW
                        Cells(CurrentCSVrow, 10).Value = DataArray(i, 7)
                        Cells(CurrentCSVrow, 11).Value = DataArray(i, 8)
                        Cells(CurrentCSVrow, 12).Value = 1
                End If
        Next i
Next k

        'Check that Self-scores are blank or 0, 1, 2 ,3 4 or 5 and that SM scores are blank or 0, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5 or 5
        Sheets("csv").Select
        For i = 2 To CurrentCSVrow
                If Cells(i, 7).Value <> "" And Cells(i, 7).Value <> 0 And Cells(i, 7).Value <> 1 And Cells(i, 7).Value <> 2 And Cells(i, 7).Value <> 3 And Cells(i, 7).Value <> 4 And Cells(i, 7).Value <> 5 Then
                        msgbox "Self-score is not blank or 0-5 in CSV row " & i
                        Exit Sub
                End If
                If Cells(i, 10).Value = "" Or Cells(i, 10).Value = "0" Or Cells(i, 10).Value = "1" Or Cells(i, 10).Value = "2" Or Cells(i, 10).Value = "3" Or Cells(i, 10).Value = "4" Or Cells(i, 7).Value = "5" Or _
                Cells(i, 10).Value = "0.5" Or Cells(i, 10).Value = "1.5" Or Cells(i, 10).Value = "2.5" Or Cells(i, 10).Value = "3.5" Or Cells(i, 10).Value = "4.5" Then
                        'do nothing
                Else 'NOT SURE WHY BUT THICS CHECK DOESN'T WORK SOMETIMES...
                        msgbox "SM score is not blank or 0-5 in CSV row " & i
                        Exit Sub
                End If
        Next i

    Sheets("csv").Select
    Cells.Select
    With Selection
        .WrapText = False
    End With
    
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True


End Sub
