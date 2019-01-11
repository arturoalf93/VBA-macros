Attribute VB_Name = "Hide_rows"

Sub Hide_Sourcing()
'
' Hide_SXM Macro
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("58:74").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("381:518").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    
    Sheets("Company Information").Select
    Rows("28:28").Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select
    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_SXM()
'
' Hide_SXM Macro
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("75:80").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("519:567").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    
    Sheets("Company Information").Select
    Rows("29:29").Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select
    
    Sheets("Instructions").Select
    Range("A1").Select
End Sub
Sub Hide_SA()
'
' Hide_SA Macro
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("81:85").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("568:616").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    
    Sheets("Company Information").Select
    Rows("30:30").Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select
    Sheets("Instructions").Select
    Range("A1").Select
End Sub
Sub Hide_CLM()
'
' Hide_CLM Macro
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("86:97").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("617:687").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    
    Sheets("Company Information").Select
    Rows("31:31").Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select
    
    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_ePRO()
'
' Hide_ePRO Macro
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("98:144").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("688:949").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    
    Sheets("Company Information").Select
    Rows("32:32").Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select
    
    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_I2P()
Attribute Hide_I2P.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hide_I2P Macro
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("145:167").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("950:1113").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True
    
    Sheets("Company Information").Select
    Rows("33:33").Select
    Selection.EntireRow.Hidden = True
    Range("A1").Select
    
    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_neither_Sourcing_SXM()
'
' Hide_neither_Sourcing_SXM

    Call Hide_Sourcing
    Call Hide_SXM
    
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("29:53").Select
    Selection.EntireRow.Hidden = True
    Rows("57:57").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("222:347").Select
    Selection.EntireRow.Hidden = True
    Rows("375:380").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_only_CLM()
'
' Hide_neither_Sourcing_SXM

    Call Hide_neither_Sourcing_SXM
    Call Hide_SA
    Call Hide_ePRO
    Call Hide_I2P
    
    
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("18:20").Select
    Selection.EntireRow.Hidden = True
    Rows("15:15").Select
    Selection.EntireRow.Hidden = True
    Rows("17:17").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("111:127").Select
    Selection.EntireRow.Hidden = True
    Rows("92:99").Select
    Selection.EntireRow.Hidden = True
    Rows("105:110").Select
    Selection.EntireRow.Hidden = True
    Rows("203:204").Select
    Selection.EntireRow.Hidden = True
    Rows("365:367").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_only_ePRO_or_I2P()
'
' Hide_neither_Sourcing_SXM
    Call Hide_neither_Sourcing_SXM
    Call Hide_SA
    Call Hide_CLM
    
    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("56:56").Select
    Selection.EntireRow.Hidden = True
    Rows("16:16").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("371:374").Select
    Selection.EntireRow.Hidden = True
    Rows("100:104").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
    

End Sub

Sub Hide_FIX_EPROI2P()
'
' Hide_FIX_EPROI2P

    Dim pw As String
    pw = "****"
    
    Sheets("Index & Average Scores").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("18:20").Select
    Selection.EntireRow.Hidden = True
    Rows("12:12").Select
    Selection.EntireRow.Hidden = True
    Range("A2").Select
    Sheets("Index & Average Scores").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("111:127").Select
    Selection.EntireRow.Hidden = True
    Rows("73:80").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
    
End Sub

Sub Hide_Type_1()
    Dim pw As String
    pw = "****"
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("363:364").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_Type_2()
    Dim pw As String
    pw = "****"
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("365:365").Select
    Selection.EntireRow.Hidden = True
    Rows("367:367").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_Type_3()
    Dim pw As String
    pw = "****"
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("368:368").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_Type_4()
    Dim pw As String
    pw = "****"
    
    Sheets("RFI").Select
    ActiveSheet.Unprotect Password:=pw
    Rows("59:59").Select
    Selection.EntireRow.Hidden = True
    Rows("61:62").Select
    Selection.EntireRow.Hidden = True
    Range("E3").Select
    ActiveSheet.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True

    Sheets("Instructions").Select
    Range("A1").Select
End Sub

Sub Hide_Type_ALL()
    Call Hide_Type_1
    Call Hide_Type_2
    Call Hide_Type_3
    Call Hide_Type_4
End Sub
