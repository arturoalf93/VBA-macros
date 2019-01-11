Attribute VB_Name = "NewRFI"
Option Explicit
'Important: this Dim line must be at the top of your module


Sub CreateTemplate()
    Dim wb1, wb2 As Excel.Workbook
    Dim ws As Worksheet
    
    'How csv is created. Maybe this method has to be used
'    Sheets.Add.Name = "csv"
'    Worksheets("csv").Move After:=Worksheets(SheetsNames(NumberOfCategorySheets - 1))
'    Sheets("csv").Select
    
    Set wb1 = ThisWorkbook
    Sheets.Add.Name = "RFI"
    Set ws = Sheets("RFI")
    Set wb2 = Workbooks.Open("/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Q2 19 RFI Change/New RFI template.xlsx")
    Cells.Select
    Selection.Copy
    ws.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.DisplayAlerts = False
    wb2.Close
    Application.DisplayAlerts = True
    ws.Activate
    Range("A1").Select
End Sub
Sub CreateIndex()
    Dim wb1, wb2 As Excel.Workbook
    Dim ws As Worksheet
    
    'How csv is created. Maybe this method has to be used
'    Sheets.Add.Name = "csv"
'    Worksheets("csv").Move After:=Worksheets(SheetsNames(NumberOfCategorySheets - 1))
'    Sheets("csv").Select
    
    Set wb1 = ThisWorkbook
    Sheets.Add.Name = "Index & Average Scores"
    Set ws = Sheets("Index & Average Scores")
    Set wb2 = Workbooks.Open("/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Q2 19 RFI Change/Index template.xlsx")
    Cells.Select
    Selection.Copy
    ws.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.DisplayAlerts = False
    wb2.Close
    Application.DisplayAlerts = True
    ws.Activate
    Range("A1").Select
End Sub

Sub CreateInstructions()
    Dim wb1, wb2 As Excel.Workbook
    Dim ws As Worksheet
    
    'How csv is created. Maybe this method has to be used
'    Sheets.Add.Name = "csv"
'    Worksheets("csv").Move After:=Worksheets(SheetsNames(NumberOfCategorySheets - 1))
'    Sheets("csv").Select
    
    Set wb1 = ThisWorkbook
    Sheets.Add.Name = "Instructions"
    Set ws = Sheets("Instructions")
    Set wb2 = Workbooks.Open("/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Q2 19 RFI Change/New Instructions.xlsx")
    Cells.Select
    Selection.Copy
    ws.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.DisplayAlerts = False
    wb2.Close
    Application.DisplayAlerts = True
    ws.Activate
    Range("A1").Select
End Sub

'Function GetsmceIDsFromCell(MyRow, MyColumn) As String
Function GetsmceIDsFromCell(MyRow, MyColumn) As Integer()
    Dim cell, chartest As String
    Dim i, j, z, CommaPositions(100) As Integer
    Dim returnVal(100) As Integer

    j = 0
    
    cell = Trim(Cells(MyRow, MyColumn).Value)
    For i = 1 To Len(cell)
        If Mid(cell, i, 1) = "," Then
            CommaPositions(j) = i
            j = j + 1
        End If
    Next i
    'j is now the number of commas
    If CommaPositions(0) = 0 Then 'If there is only 1 id, no commas
        returnVal(0) = Mid(cell, 3, Len(cell) - 3 + 1)
    Else
        returnVal(0) = Mid(cell, 3, CommaPositions(0) - 3)
        If j >= 2 Then
            For z = 1 To j - 1
                returnVal(z) = Mid(cell, CommaPositions(z - 1) + 2, CommaPositions(z) - (CommaPositions(z - 1) + 2))
            Next z
        End If
        returnVal(j) = Mid(cell, CommaPositions(j - 1) + 2, Len(cell) - (CommaPositions(j - 1) + 1))
    End If
    
    GetsmceIDsFromCell = returnVal()
    
End Function

Function GetQuarterColumn(row, column) As Integer
Dim j As Integer
    GetQuarterColumn = -1
    For j = column To (j - 7) Step -1
        If Cells(row - 1, j).Value = "Q2 17" Then
            GetQuarterColumn = 1
            Exit For
        ElseIf Cells(row - 1, j).Value = "Q4 17" Then
            GetQuarterColumn = 2
            Exit For
        ElseIf Cells(row - 1, j).Value = "Q1 18" Then
            GetQuarterColumn = 3
            Exit For
        ElseIf Cells(row - 1, j).Value = "Q2 18" Then
            GetQuarterColumn = 4
            Exit For
        ElseIf Cells(row - 1, j).Value = "Q3 18" Then
            GetQuarterColumn = 5
            Exit For
        ElseIf Cells(row - 1, j).Value = "Q4 18" Then
            GetQuarterColumn = 6
            Exit For
        End If
    Next j
    If GetQuarterColumn = -1 Then
        msgbox "Error, QuarterColumn not found for row " & row & " and column " & " j"
    End If
End Function
Sub HideAll()

Dim ws As Worksheet

For Each ws In Worksheets
    If ws.Name = "P2P" Or ws.Name = "Sourcing" Or ws.Name = "Spend Analytics" Or ws.Name = "SXM" Or ws.Name = "CLM" Then
        'ws.Unprotect Password:=pw
        Sheets(ws.Name).Visible = False
    End If
Next

End Sub

Sub SetCurrentSSandSMaveragesFormulas()
Dim CurrentSMcolumn, i, LastIndexRow As Integer

LastIndexRow = 167

Sheets("RFI").Activate
CurrentSMcolumn = CurrentScoreColumn()

Sheets("Index & Average Scores").Activate
For i = 2 To LastIndexRow
    Cells(i, 5).Value = GetIndexFormula(Cells(i, 7).Value, Cells(i, 8).Value, CurrentSMcolumn)
    Cells(i, 4).Value = GetIndexFormula(Cells(i, 7).Value, Cells(i, 8).Value, CurrentSMcolumn - 1)
Next i

End Sub
Function GetIndexFormula(ByVal StartRow As Integer, ByVal EndRow As Integer, ByVal CurrentSMcolumn As Integer) As String
Dim j As Integer
'=IF(ISNUMBER(AVERAGE(RFI!AA7:AA9)),AVERAGE(RFI!AA7:AA9),"-")
GetIndexFormula = "=IF(ISNUMBER(AVERAGE(RFI!" & Split(Cells(1, CurrentSMcolumn).Address(True, False), "$")(0) & StartRow & ":" & _
Split(Cells(1, CurrentSMcolumn).Address(True, False), "$")(0) & EndRow & ")),AVERAGE(RFI!" & Split(Cells(1, CurrentSMcolumn).Address(True, False), "$")(0) _
& StartRow & ":" & Split(Cells(1, CurrentSMcolumn).Address(True, False), "$")(0) & EndRow & ")," & Chr(34) & "-" & Chr(34) & ")"

End Function

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function


Sub Main()
    Application.ScreenUpdating = True

    Dim i, j, k, z, smceID, RFIsmceIDcolumn, ProvIDcolumn, AffColumn, Q217column, IndividualsmceIDcolumn, Indi, Indj, Indz, ExitForFlag, _
    quarter, HeaderRow, CsColumn, SDcolumnsIndex, l, m, OldscseidColumn, UnchangedColumn, Oldscseid As Integer
    Dim ids() As Integer
    Dim history(6) As String 'history(0) will always be nothing
    ' 1: Q217
    '2: Q4 17
    '3: Q1 18
    '4: Q2 18
    '5: Q3 18
    '6: Q4 18
    Dim module, element, SDorR, pw, UnchangedSheetName As String
    Dim ws As Worksheet
    Dim SDcolumns(20, 2) As Integer ' First item: column. Second item: 1, 2, 3, 4, 5 or 6 depending on the quarter. Third item: 1=Self-Description, 2=Reasoning
    Dim HideQuarterFlag(6) As Integer 'HideQuarterFlag(0) will always be nothing. The rest will be 0 by default.
    Dim HideLastScoreFlag(2) As Integer '0 by default
    Dim LastSS, LastSMscore As Variant
    
    RFIsmceIDcolumn = 1
    AffColumn = 2
    OldscseidColumn = 3
    UnchangedColumn = 4

    Q217column = 8
    pw = "****"
    
    'Add customer counts in Company Information
    Sheets("Company Information").Select
    Range("A1").Select
    Rows("29:29").Select
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Range("B29").Select
    ActiveCell.FormulaR1C1 = "Sourcing customer count"
    Range("B30").Select
    ActiveCell.FormulaR1C1 = "SXM customer count"
    Range("B31").Select
    ActiveCell.FormulaR1C1 = "Spend Analytics customer count"
    Range("B32").Select
    ActiveCell.FormulaR1C1 = "CLM customer count"
    Range("B33").Select
    ActiveCell.FormulaR1C1 = "eProcurement customer count"
    Range("B34").Select
    ActiveCell.FormulaR1C1 = "I2P customer count"
    Range("B29:B34").Select
    Selection.Font.Bold = True
    If Cells(16, 2).Value = "Please list 3 reference customers and reference customer contact information:" Then
        Rows("16:16").Select
        Selection.Delete Shift:=xlUp
    Else
        msgbox "Row 16 of Company Information different than expected"
        Exit Sub
    End If
    Range("A1").Select

    
    Call UnprotectAll
    Call CreateTemplate
    Call CreateIndex
    Application.DisplayAlerts = False
    Sheets("Instructions").Delete
    Application.DisplayAlerts = True
    Call CreateInstructions
    
    
    'Now we are in the RFI template, in our vendor workbook
    'Premise: column A will be the one with the smceIDs
    'Premise: column D will be the one with the affected old ids: "R ***, ***, *** etc..."
    
    Sheets("RFI").Activate
    For i = 3 To 1200
        ' If Not IsEmpty(Cells(i, RFIsmceIDcolumn).Value) Then
        Sheets("RFI").Activate
        If Cells(i, 1).Value <> "" Then
            history(1) = ""
            history(2) = ""
            history(3) = ""
            history(4) = ""
            history(5) = ""
            history(6) = ""
        End If
        
        If Left(Cells(i, AffColumn).Value, 1) = "R" Then
            ids = GetsmceIDsFromCell(i, AffColumn)
            
            history(1) = ""
            history(2) = ""
            history(3) = ""
            history(4) = ""
            history(5) = ""
            history(6) = ""
            
            z = 0
            Do While ids(z) <> 0
            
            For Each ws In Worksheets
                ExitForFlag = 0
                If ws.Name = "P2P" Or ws.Name = "Sourcing" Or ws.Name = "Spend Analytics" Or ws.Name = "SXM" Or ws.Name = "CLM" Then
                    Sheets(ws.Name).Activate
                    module = ws.Name
                    IndividualsmceIDcolumn = 1
                    If ws.Name = "CLM" Then
                        IndividualsmceIDcolumn = 2
                    End If
                    
                    For Indi = 1 To 1000
                        If Cells(Indi, IndividualsmceIDcolumn).Value = ids(z) Then
                            'msgbox "Found in " & module
                            element = Cells(Indi, IndividualsmceIDcolumn + 1).Value
                            
                            'Reset SDcolumns
                            For l = 0 To 20
                                For m = 0 To 2
                                    SDcolumns(l, m) = 0
                                Next m
                            Next l
                            SDcolumnsIndex = 0
                            
                            'Find Self-Description/Reasoning columns, quarter
                            HeaderRow = CurrentScoreRow()
                            CsColumn = CurrentScoreColumn()
                            For Indj = 1 To CsColumn
                                If Cells(HeaderRow, Indj).Value = "Self-description" Or Cells(HeaderRow, Indj).Value = "Self-Description" Or Cells(HeaderRow, Indj).Value = "Self -Description" Then
                                    SDcolumns(SDcolumnsIndex, 0) = Indj
                                    SDcolumns(SDcolumnsIndex, 1) = GetQuarterColumn(HeaderRow, Indj)
                                    SDcolumns(SDcolumnsIndex, 2) = 1
                                    SDcolumnsIndex = SDcolumnsIndex + 1
                                ElseIf Cells(HeaderRow, Indj).Value = "Reasoning" Then
                                    SDcolumns(SDcolumnsIndex, 0) = Indj
                                    SDcolumns(SDcolumnsIndex, 1) = GetQuarterColumn(HeaderRow, Indj)
                                    SDcolumns(SDcolumnsIndex, 2) = 2
                                    SDcolumnsIndex = SDcolumnsIndex + 1
                                End If
                            Next Indj
                            
                            
                            Indz = 0
                            Do While SDcolumns(Indz, 0) <> 0
                                If Cells(Indi, SDcolumns(Indz, 0)) <> "" Then
                                    HideQuarterFlag(SDcolumns(Indz, 1)) = 1 'If 1 it will not be hidden
                                    quarter = SDcolumns(Indz, 1)
                                    If SDcolumns(Indz, 2) = 1 Then
                                        SDorR = "Self-Description"
                                    ElseIf SDcolumns(Indz, 2) = 2 Then
                                        SDorR = "Reasoning"
                                    Else
                                        msgbox "Error, SDcolumns(indz,2) is not neither 1 or 2, for indi: " & Indi
                                    End If
                                    
                                    history(quarter) = history(quarter) & module & " - " & element & " (" & SDorR & "):" & vbNewLine & Cells(Indi, SDcolumns(Indz, 0)).Value & vbNewLine & vbNewLine
                                
                                
                                End If
                                Indz = Indz + 1
                            Loop
                            
                            ExitForFlag = 1
                            Exit For
                        End If
                    Next Indi
                    If ExitForFlag = 1 Then
                        Exit For 'Premise: one smceID can only be found in one individual module, so once it is found, we jump to the next one
                    End If
                End If
            Next ws
            

            
            
            z = z + 1
            Loop
            

                       Sheets("RFI").Activate
            
            For k = 1 To 6
                If history(k) <> "" Then
                    history(k) = Left(history(k), Len(history(k)) - 2)  'Removing the last 2 vbNewLine
                    Cells(i, Q217column + k - 1).Value = history(k)
                    
                    'Format: yellow and borders
                    Cells(i, Q217column + k - 1).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent4
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                    End With
                    With Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                End If
            Next k
         
        End If 'of  If Left(Cells(i, AffColumn).Value, 1) = "R" Then
        
        'UNCHANGED elements
        If Cells(i, UnchangedColumn).Value <> "" Then
            LastSS = ""
            LastSMscore = ""
            UnchangedSheetName = ""
            Sheets("RFI").Activate
            UnchangedSheetName = Cells(i, UnchangedColumn).Value
            If WorksheetExists(UnchangedSheetName) Then
            
                Oldscseid = Cells(i, OldscseidColumn).Value
                Sheets(UnchangedSheetName).Activate
                If UnchangedSheetName = "P2P" Or UnchangedSheetName = "Sourcing" Or UnchangedSheetName = "Spend Analytics" Or UnchangedSheetName = "SXM" Then
                        IndividualsmceIDcolumn = 1
                ElseIf UnchangedSheetName = "CLM" Then
                            IndividualsmceIDcolumn = 2
                Else
                    msgbox "Error, not sheet found with taht UnchangedSheetName (" & UnchangedSheetName & ")"
                End If
                        
                For Indi = 1 To 1000
                    If Cells(Indi, IndividualsmceIDcolumn).Value = Oldscseid Then
                        If Cells(Indi, CurrentScoreColumn() - 1) <> "" Then
                            LastSS = Cells(Indi, CurrentScoreColumn() - 1)
                            HideLastScoreFlag(0) = 1
                        End If
                        If Cells(Indi, CurrentScoreColumn()) <> "" Then
                            LastSMscore = Cells(Indi, CurrentScoreColumn())
                            HideLastScoreFlag(1) = 1
                        End If
                        Exit For
                    End If
                Next Indi
                
                Sheets("RFI").Activate
                Cells(i, Q217column + 6).Value = LastSS
                Cells(i, Q217column + 7).Value = LastSMscore
                'Format: dark pink and borders
                Range(Cells(i, Q217column + 6), Cells(i, Q217column + 7)).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            End If
        End If
        
        'Format the E column yellow if there is no SD history
        If Cells(i, 1).Value <> "" Then
                If history(1) = "" And history(2) = "" And history(3) = "" And history(4) = "" And history(5) = "" And history(6) = "" Then
                    'Add \n + (NEW)
                    Cells(i, 5).Value = Cells(i, 5).Value & vbCrLf & "(NEW)"
                    'Format: yellow
                    Cells(i, 5).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                Else 'There is at least one SD
                    If Cells(i, 14).Value = "" Then 'There is no SS
                        'Add \n + (REVISED)
                        Cells(i, 5).Value = Cells(i, 5).Value & vbCrLf & "(REVISED)"
                        'Format rose
                        Cells(i, 5).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent2
                            .TintAndShade = 0.599993896298105
                            .PatternTintAndShade = 0
                        End With
                    End If
                End If
        End If
        
    Next i
    
    Sheets("RFI").Activate
    'Hide those History quarter columns which flag is still 0
    For j = 1 To 6
        If HideQuarterFlag(j) = 0 Then Columns(Q217column - 1 + j).EntireColumn.Hidden = True
    Next j
    
    'Hide those LastScore columns which lag is still 0
    If HideLastScoreFlag(0) = 0 Then Columns(Q217column + 6).EntireColumn.Hidden = True
    If HideLastScoreFlag(1) = 0 Then Columns(Q217column + 7).EntireColumn.Hidden = True
    
    'Finish nailing the tab 'RFI'
    '
    '
    '
    
    'Create hyperlinks in Index & Average Scores
    Sheets("Index & Average Scores").Activate
    For i = 2 To 167
        ActiveSheet.Hyperlinks.Add Anchor:=Range("A" & i), Address:="", SubAddress:="RFI!E" & Cells(i, 7).Value
    Next i
        
    Columns("B:B").Select
    Selection.Copy
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    
    
    'Set formulas in Current SS and Current SM score averages columns
    Sheets("Index & Average Scores").Activate
    Call SetCurrentSSandSMaveragesFormulas

    
    'Hide link column, two MAP columns
    Sheets("Index & Average Scores").Activate
    Columns(6).EntireColumn.Hidden = True
    Columns(7).EntireColumn.Hidden = True
    Columns(8).EntireColumn.Hidden = True
    
    'Freeze Index & Average Scores first row
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    'Hide first columns and Freeze RFI headers
    Sheets("RFI").Activate
    Columns(1).EntireColumn.Hidden = True
    Columns(2).EntireColumn.Hidden = True
    Columns(3).EntireColumn.Hidden = True
    Columns(4).EntireColumn.Hidden = True
    Range("E1").Select
    Range("F3").Select
    ActiveWindow.FreezePanes = True
    
    'Lock RFI and Index & Average Scores
    Sheets("RFI").Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    Sheets("Index & Average Scores").Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    'Hide P2P, Sourcing, Spend Analytics, SXM and CLM tabs
    Call HideAll
    
    'Lock workbook so they cannot unhide or create sheets
    ActiveWorkbook.Protect Password:=pw, Structure:=True, Windows:=False
    
    
    Application.ScreenUpdating = True
End Sub


Sub Many_files()
'
' Many_files Macro
'

    Dim FileName As String
    Dim NewFileName As String
    Dim i As Integer
    Dim MyFiles(4) As String
    
    Dim wb1, wb2 As Excel.Workbook
    
    MyFiles(0) = "Synertrade Q1 19_S2P_v1.xlsx"
    MyFiles(1) = "Zycus_S2P Q1 19_v1.xlsx"
    MyFiles(2) = ""
    MyFiles(3) = ""
    MyFiles(4) = ""
    

        For i = 0 To 1
            On Error Resume Next

            FileName = MyFiles(i)
    
            Set wb2 = Workbooks.Open("/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Vendors/Q2 19/Transformation folder/" & FileName)
            NewFileName = Left(FileName, Len(FileName) - 5)
            NewFileName = NewFileName & "_Q2_19.xlsx"
            
            Call Main
            
            ActiveWorkbook.SaveAs FileName:= _
            "/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Vendors/Q2 19/Transformation folder/" & NewFileName _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            wb2.Close
            
            On Error GoTo 0
        Next i

End Sub

Sub Many_files_FIX_EPROI2P()
'
' Many_files_FIX_EPROI2P Macro
'

    Dim FileName As String
    Dim NewFileName As String
    Dim i As Integer
    Dim MyFiles(14) As String
    
    Dim wb1, wb2 As Excel.Workbook
    
    MyFiles(0) = "Tradeshift.xlsx"
    MyFiles(1) = "Taulia.xlsx"
    MyFiles(2) = "Oracle.xlsx"
    MyFiles(3) = "Birchstreet.xlsx"
    MyFiles(4) = "BuyerQuest.xlsx"
    MyFiles(5) = "Yooz.xlsx"
    MyFiles(6) = "Claritum.xlsx"
    MyFiles(7) = "Nimbi.xlsx"
    MyFiles(8) = "Prodigo_Solutions.xlsx"
    MyFiles(9) = "OpusCapita.xlsx"
    MyFiles(10) = "Basware.xlsx"
    MyFiles(11) = "wescale.xlsx"
    MyFiles(12) = "ProPocure.xlsx"
    MyFiles(13) = "Vroozi.xlsx"
    MyFiles(14) = "Aquiire_(now Coupa).xlsx"
    

        For i = 0 To 14

            FileName = MyFiles(i)
    
            Set wb2 = Workbooks.Open("/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Vendors/Q2 19/Transformation folder/" & FileName)
            NewFileName = Left(FileName, Len(FileName) - 5)
            NewFileName = NewFileName & "_Q2_19.xlsx"
            
            Call Hide_FIX_EPROI2P
            
            ActiveWorkbook.SaveAs FileName:= _
            "/Users/arodriguez/Dropbox/Work/SpendMatters/SolutionMaps/Vendors/Q2 19/Transformation folder/" & NewFileName _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            wb2.Close
            
        Next i

End Sub
