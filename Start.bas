Attribute VB_Name = "Module1"
Dim Bond As String
Dim ISIN As String
Dim Issuer As String
Dim Crncy As String
Dim Amount As Variant
'Rating
    Dim Collateral_Type As String
    Dim Moody_1, Moody_2, Moody_3, Moody_4, Moody_5, Moody_6 As String
    Dim SNP_1, SNP_2 As String
    Dim Fitch_1, Fitch_2, Fitch_3, Fitch_4 As String
    Dim Moody, SNP, Fitch As String
    Dim Credit_Rating_Array As Variant
    Dim Credit_Rating As String
'Tenor
    Dim Issued_Date As String
    Dim First_Call_Date As String
    Dim Maturity_Date As String
    Dim Total_Tenor, Non_Callable_Tenor As Variant
    Dim Tenor As Variant
Dim Issued_Rate As Variant
Dim Issued_Price As Variant
Dim Issued_Spread As String
Dim Guarantor As String
Dim Industry As String
Dim Coupon_Type As String
Dim Security_Type As String
Dim Series As String
Dim Filter_Type As String

'Main - filter recommended new issued bonds
Sub Recommended_New_Issues()
    
    Application.ScreenUpdating = False
    
    'check if any ISIN is inputted
    Dim ISIN_Amount As Long
    ISIN_Amount = Worksheets("Raw").Cells(Rows.Count, 1).End(xlUp).row - 19
    If ISIN_Amount = 0 Then
        Worksheets("Raw").Range("A1").Select
        MsgBox "No ISIN is inputted!"
        GoTo Skip
    End If
    
    Call Format_Raw
    Call Clear_Sheets
    
    'deal with bonds
    Raw_Last = Worksheets("Raw").Cells(Rows.Count, 1).End(xlUp).row
    For i = 20 To Raw_Last
        Worksheets("Raw").Activate
        Cells(i, 1).Select
        ActiveWindow.ScrollRow = Selection.row
        Call Define_Variables(i)
        Call Select_Sheet
        Call Filter_Bond
        If Filter_Type = "Y" Then Call Paste_Bond
    Next i
    
    Call Format_Sheets
            
    Worksheets("Raw").Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = 20
    MsgBox "Done!"

Skip:
    
    Application.ScreenUpdating = True

End Sub

'Sub - format "Raw" worksheet
Sub Format_Raw()
    
    Industry_Column = WorksheetFunction.Match("Industry", Worksheets("Raw").Range("19:19"), 0)
    Issuer_Column = WorksheetFunction.Match("Issuer", Worksheets("Raw").Range("19:19"), 0)
    ISIN_Column = WorksheetFunction.Match("ISIN", Worksheets("Raw").Range("19:19"), 0)
    Last_Row = Worksheets("Raw").Cells(Rows.Count, 1).End(xlUp).row
    Last_Column = Worksheets("Raw").Cells(1, Columns.Count).End(xlToLeft).Column
    
    'delete same ISIN
    Range("A19", Cells(Last_Row, Last_Column)).Select
    Selection.RemoveDuplicates Columns:=ISIN_Column, Header:=xlYes
    'sort Industry and Issuer
    Range("A19", Cells(Last_Row, Last_Column)).Select
    Selection.Sort Key1:=Worksheets("Raw").Cells(19, Industry_Column), Order1:=xlAscending, Key2:=Worksheets("Raw").Cells(19, Issuer_Column), Order2:=xlAscending, Header:=xlYes
    
    Range("C19", Range("C19").End(xlToRight).End(xlDown)).Columns.AutoFit

End Sub

'Sub - clear sheets
Sub Clear_Sheets()
    
    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Activate
        If sheet.Name <> "Raw" Then
            Range("2:" & Rows.Count).ClearContents
        End If
    Next
      
End Sub

'Sub - define variables
Sub Define_Variables(i)
    
    Last_Row = Worksheets("Raw").Cells(Rows.Count, 1).End(xlUp).row
    Last_Column = Worksheets("Raw").Cells(1, Columns.Count).End(xlToLeft).Column
    
    With Application.WorksheetFunction
        Bond = .HLookup("Bond", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        ISIN = .HLookup("ISIN", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Issuer = .HLookup("Issuer", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Crncy = .HLookup("Currency", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        'Issued Amount
            On Error Resume Next
            Amount = .HLookup("Issued Amount", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False) / 1000000
            On Error GoTo 0
        'Rating
            Collateral_Type = .HLookup("Collateral Type", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            'Moody
            Moody_1 = .HLookup("Moody", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Moody_2 = .HLookup("Moody (Sr)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Moody_3 = .HLookup("Moody (Sub)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Moody_4 = .HLookup("Moody (Jr Sub)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Moody_5 = .HLookup("Moody (LT)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Moody_6 = .HLookup("Moody (Issuer)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            If Moody_1 <> "#N/A N/A" And InStr(Moody_1, "WD") = 0 And InStr(Moody_1, "WR") = 0 And InStr(Moody_1, "NR") = 0 Then
                Moody = Moody_1
            ElseIf Collateral_Type = "SR UNSECURED" And Moody_2 <> "#N/A N/A" And InStr(Moody_2, "WD") = 0 And InStr(Moody_2, "WR") = 0 And InStr(Moody_2, "NR") = 0 Then
                Moody = Moody_2
            ElseIf Collateral_Type = "SUBORDINATED" And Moody_3 <> "#N/A N/A" And InStr(Moody_3, "WD") = 0 And InStr(Moody_3, "WR") = 0 And InStr(Moody_3, "NR") = 0 Then
                Moody = Moody_3
            ElseIf Collateral_Type = "JR SUBORDINATED" And Moody_4 <> "#N/A N/A" And InStr(Moody_4, "WD") = 0 And InStr(Moody_4, "WR") = 0 And InStr(Moody_4, "NR") = 0 Then
                Moody = Moody_4
            ElseIf Moody_5 <> "#N/A N/A" And InStr(Moody_5, "WD") = 0 And InStr(Moody_5, "WR") = 0 And InStr(Moody_5, "NR") = 0 Then
                Moody = Moody_5
            ElseIf Moody_6 <> "#N/A N/A" And InStr(Moody_6, "WD") = 0 And InStr(Moody_6, "WR") = 0 And InStr(Moody_6, "NR") = 0 Then
                Moody = Moody_6
            Else
                Moody = "-"
            End If
            'S&P
            SNP_1 = .HLookup("S&P", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            SNP_2 = .HLookup("S&P (Issuer)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            If SNP_1 <> "#N/A N/A" And InStr(SNP_1, "WD") = 0 And InStr(SNP_1, "WR") = 0 And InStr(SNP_1, "NR") = 0 Then
                SNP = SNP_1
            ElseIf SNP_2 <> "#N/A N/A" And InStr(SNP_2, "WD") = 0 And InStr(SNP_2, "WR") = 0 And InStr(SNP_2, "NR") = 0 Then
                SNP = SNP_2
            Else
                SNP = "-"
            End If
            'Fitch
            Fitch_1 = .HLookup("Fitch", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Fitch_2 = .HLookup("Fitch (Sr)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Fitch_3 = .HLookup("Fitch (Sub)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Fitch_4 = .HLookup("Fitch (Issuer)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            If Fitch_1 <> "#N/A N/A" And InStr(Fitch_1, "WD") = 0 And InStr(Fitch_1, "WR") = 0 And InStr(Fitch_1, "NR") = 0 Then
                Fitch = Fitch_1
            ElseIf Collateral_Type = "SR UNSECURED" And Fitch_2 <> "#N/A N/A" And InStr(Fitch_2, "WD") = 0 And InStr(Fitch_2, "WR") = 0 And InStr(Fitch_2, "NR") = 0 Then
                Fitch = Fitch_2
            ElseIf Collateral_Type = "SUBORDINATED" And Fitch_3 <> "#N/A N/A" And InStr(Fitch_3, "WD") = 0 And InStr(Fitch_3, "WR") = 0 And InStr(Fitch_3, "NR") = 0 Then
                Fitch = Fitch_3
            ElseIf Fitch_4 <> "#N/A N/A" And InStr(Fitch_4, "WD") = 0 And InStr(Fitch_4, "WR") = 0 And InStr(Fitch_4, "NR") = 0 Then
                Fitch = Fitch_4
            Else
                Fitch = "-"
            End If
            'Integrated Rating
            Credit_Rating_Array = Array("(", Moody, "/", SNP, "/", Fitch, ")")
            Credit_Rating = Join(Credit_Rating_Array, "")
        'Tenor
            Issued_Date = .HLookup("Issued Date", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            First_Call_Date = .HLookup("First Call Date", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            Maturity_Date = .HLookup("Maturity Date", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            On Error Resume Next
            Err.Clear
                Total_Tenor = Round((DateValue(Maturity_Date) - DateValue(Issued_Date)) / 365, 1)
                If Err.Number <> 0 Then Total_Tenor = "-"
            Err.Clear
                Non_Callable_Tenor = Round((DateValue(First_Call_Date) - DateValue(Issued_Date)) / 365, 1)
                If Err.Number <> 0 Then Non_Callable_Tenor = "-"
                'delete the tenth number of 0/9/1
                If Right(CStr(Total_Tenor), 2) = ".0" Or Right(CStr(Total_Tenor), 2) = ".9" Or Right(CStr(Total_Tenor), 2) = ".1" Then Total_Tenor = Round(Total_Tenor, 0)
                If Right(CStr(Non_Callable_Tenor), 2) = ".0" Or Right(CStr(Non_Callable_Tenor), 2) = ".9" Or Right(CStr(Non_Callable_Tenor), 2) = ".1" Then Non_Callable_Tenor = Round(Non_Callable_Tenor, 0)
                'delete NC in bonds with callable tenor period of less than six months
                If Total_Tenor <> "-" And Non_Callable_Tenor <> "-" Then
                    If Total_Tenor - Non_Callable_Tenor <= 0.5 Then Non_Callable_Tenor = "-"
                End If
            On Error GoTo 0
            'Integrated Tenor
            If Total_Tenor <> "-" Then
                If Non_Callable_Tenor <> "-" Then
                    Tenor = Total_Tenor & "NC" & Non_Callable_Tenor
                Else
                    Tenor = Total_Tenor
                End If
            Else
                If Non_Callable_Tenor <> "-" Then
                    Tenor = "NC" & Non_Callable_Tenor
                Else
                    Tenor = "-"
                End If
            End If
        'Issued Rate
            Issued_Rate = .HLookup("Fixed Reoffered Rate (%)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            If VarType(Issued_Rate) = vbString Then
                Issued_Rate = .HLookup("Issued Rate (%)", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
                'replace issued rate with the rate within bond name
                If VarType(Issued_Rate) = vbString Then
                    If Right(Bond, 4) = "PERP" Then
                        Issued_Rate = Mid(Bond, InStr(Bond, " ") + 1, Len(Bond) - InStr(Bond, " ") - 5)
                    Else
                        Issued_Rate = Mid(Bond, InStr(Bond, " ") + 1, Len(Bond) - InStr(Bond, " ") - 9)
                    End If
                End If
                'deal with issued rate data
                On Error Resume Next
                Err.Clear
                    If Not IsNumeric(Issued_Rate) And InStr(Issued_Rate, "/") = 0 Then
                        Issued_Rate = "-"
                    ElseIf InStr(Issued_Rate, "/") <> 0 Then
                        Issued_Rate = CInt(Left(Issued_Rate, InStr(Issued_Rate, " ") - 1)) + CInt(Mid(Issued_Rate, InStr(Issued_Rate, " ") + 1, InStr(Issued_Rate, "/") - InStr(Issued_Rate, " ") - 1)) / CInt(Right(Issued_Rate, Len(Issued_Rate) - InStr(Issued_Rate, "/")))
                        If Err.Number <> 0 Then Issued_Rate = "-"
                    Else
                        Issued_Rate = CDbl(Issued_Rate)
                    End If
                On Error GoTo 0
            End If
        'Issued Price
            Issued_Price = .HLookup("Fixed Reoffered Price", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            If VarType(Issued_Price) = vbString Then Issued_Price = .HLookup("Issued Price", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        'Issued Spread
            Issued_Spread = .HLookup("Fixed Reoffered Spread", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
            If VarType(Issued_Spread) = vbString Then Issued_Spread = .HLookup("Issued Spread", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Guarantor = .HLookup("Guarantor", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Industry = .HLookup("Industry", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Coupon_Type = .HLookup("Coupon Type", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Security_Type = .HLookup("Security Type", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
        Series = .HLookup("Series", Range("A19", Cells(Last_Row, Last_Column)), i - 18, False)
    End With
    
End Sub

'Sub - select sheet to paste
Sub Select_Sheet()
    
    If InStr(Collateral_Type, "SUBORDINATED") = 0 And Total_Tenor <> "-" Then
        If Industry <> "Government" Then
            Worksheets("Senior(corp)").Activate
        Else
            Worksheets("Senior(sov)").Activate
        End If
    Else
        If Industry <> "Government" Then
            Worksheets("Sub&Perp(corp)").Activate
        Else
            Worksheets("Sub&Perp(sov)").Activate
        End If
    End If
    
End Sub

'Sub - filter bonds
Sub Filter_Bond()
        
    Dim Rate_LT_Lower As Double
    Dim Rate_ST_Lower As Double
    Dim Boundary As Double
    Dim Amount_Lower As Double
    Dim Year_Upper, Year_Lower As Variant
    Dim Rating_Upper, Rating_Lower As String
    Dim Total_Tenor_Altered As Variant
    Dim Moody_Array, SNP_Array, Fitch_Array As Variant
    Dim Moody_Altered, SNP_Altered, Fitch_Altered As String
    Dim Upper_Index, Lower_Index As String
    Dim Moody_Filter_Type, SNP_Filter_Type, Fitch_Filter_Type As String
    
    Filter_Type = "N"
    
    Rate_LT_Lower = Worksheets("Raw").Range("B4") * 100
    Rate_ST_Lower = Worksheets("Raw").Range("B5") * 100
    Boundary = Worksheets("Raw").Range("B6")
        If Worksheets("Raw").Range("B6") = "" Then Boundary = 10
    Amount_Lower = Worksheets("Raw").Range("B8")
    Year_Upper = Worksheets("Raw").Range("B10")
        If Worksheets("Raw").Range("B10") = "" Or Worksheets("Raw").Range("B10") = "PERP" Then Year_Upper = 1000000
    Year_Lower = Worksheets("Raw").Range("B11")
        If Worksheets("Raw").Range("B11") = "" Then Year_Lower = 0
        If Worksheets("Raw").Range("B11") = "PERP" Then Year_Lower = 1000000
    Rating_Upper = Worksheets("Raw").Range("B13")
        If Worksheets("Raw").Range("B13") = "" Then Rating_Upper = "AAA"
    Rating_Lower = Worksheets("Raw").Range("B14")
        If Worksheets("Raw").Range("B14") = "" Then Rating_Lower = "C"
    
    'embedded filtering criteria
        'Series
            If UCase(Series) <> "REGS" And UCase(Series) <> "EMTN" And UCase(Series) <> "GMTN" And UCase(Series) <> "MTN" And Series <> "#N/A Field Not Applicable" Then GoTo Next_One
        'Security Type
            If Security_Type = "PRIV PLACEMENT" Or Security_Type = "ASSET-BASED REV" Or Security_Type = "TERM" Or Security_Type = "REV" Then GoTo Next_One
        'CouponType
            If Coupon_Type <> "FIXED" And Coupon_Type <> "VARIABLE" And Coupon_Type <> "#N/A Field Not Applicable" Then GoTo Next_One
            If InStr(Bond, "Float") Then GoTo Next_One
    'non-embedded filtering criteria
        'Amount
            If VarType(Amount) = vbString Or Amount < Amount_Lower Then GoTo Next_One
        'Issued Rate
            If VarType(Total_Tenor) = vbString Or Total_Tenor >= Boundary Then
                    If VarType(Issued_Rate) = vbString Or Issued_Rate < Rate_LT_Lower Then GoTo Next_One
            Else
                    If VarType(Issued_Rate) = vbString Or Issued_Rate < Rate_ST_Lower Then GoTo Next_One
            End If
        'Tenor
            Total_Tenor_Altered = CDbl(Replace(Total_Tenor, "-", 1000000))
            If Total_Tenor_Altered > CDbl(Year_Upper) Or Total_Tenor_Altered < CDbl(Year_Lower) Then GoTo Next_One
        'Credit Rating
            Moody_Array = Array("Aaa", "Aa1", "Aa2", "Aa3", "A1", "A2", "A3", "Baa1", "Baa2", "Baa3", "Ba1", "Ba2", "Ba3", "B1", "B2", "B3", "Caa1", "Caa2", "Caa3", "Ca", "C")
            SNP_Array = Array("AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C")
            Fitch_Array = Array("AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C")
            Moody_Altered = Replace(Moody, " *-", ""): Moody_Altered = Replace(Moody_Altered, " *+", ""): Moody_Altered = Replace(Moody_Altered, " (hyb)", "")
            Moody_Altered = Replace(Moody_Altered, "(P)", ""): Moody_Altered = Replace(Moody_Altered, "(EXP)", ""): Moody_Altered = Replace(Moody_Altered, "u", "")
            SNP_Altered = Replace(SNP, " *-", ""): SNP_Altered = Replace(SNP_Altered, " *+", ""): SNP_Altered = Replace(SNP_Altered, " (hyb)", "")
            SNP_Altered = Replace(SNP_Altered, "(P)", ""): SNP_Altered = Replace(SNP_Altered, "(EXP)", ""): SNP_Altered = Replace(SNP_Altered, "u", "")
            Fitch_Altered = Replace(Fitch, " *-", ""): Fitch_Altered = Replace(Fitch_Altered, " *+", ""): Fitch_Altered = Replace(Fitch_Altered, " (hyb)", "")
            Fitch_Altered = Replace(Fitch_Altered, "(P)", ""): Fitch_Altered = Replace(Fitch_Altered, "(EXP)", ""): Fitch_Altered = Replace(Fitch_Altered, "u", "")
            For i = LBound(SNP_Array) To UBound(SNP_Array)
                If SNP_Array(i) = Rating_Upper Then Upper_Index = i
                If SNP_Array(i) = Rating_Lower Then Lower_Index = i
            Next i
            Moody_Filter_Type = "N": SNP_Filter_Type = "N": Fitch_Filter_Type = "N"
            For i = Upper_Index To Lower_Index
                If Moody_Altered = "-" Or Moody_Altered = Moody_Array(i) Then Moody_Filter_Type = "Y"
                If SNP_Altered = "-" Or SNP_Altered = SNP_Array(i) Then SNP_Filter_Type = "Y"
                If Fitch_Altered = "-" Or Fitch_Altered = Fitch_Array(i) Then Fitch_Filter_Type = "Y"
            Next i
            If ActiveSheet.Name = "Senior(corp)" Or ActiveSheet.Name = "Senior(sov)" Then
                If Moody_Altered = "-" And SNP_Altered = "-" And Fitch_Altered = "-" Then GoTo Next_One
                If Moody_Filter_Type = "N" Or SNP_Filter_Type = "N" Or Fitch_Filter_Type = "N" Then GoTo Next_One
            End If
    
    Filter_Type = "Y"
    
Next_One:

End Sub

'Sub - paste recommended bonds
Sub Paste_Bond()
    
    Last_Row = Cells(Rows.Count, 1).End(xlUp).row + 1
    
    Cells(Last_Row, 1) = Issuer
    Cells(Last_Row, 2) = Crncy
    Cells(Last_Row, 3) = Amount
    Cells(Last_Row, 4) = Credit_Rating
    Cells(Last_Row, 5) = Tenor
    Cells(Last_Row, 6) = Bond
    Cells(Last_Row, 7) = ISIN
    Cells(Last_Row, 8) = Issued_Rate
    Cells(Last_Row, 9) = Issued_Price
    Cells(Last_Row, 10) = Issued_Spread
    Cells(Last_Row, 11) = Collateral_Type
    Cells(Last_Row, 12) = Guarantor
    Cells(Last_Row, 13) = Industry

End Sub
    
'Sub - format sheets
Sub Format_Sheets()
    
    Dim i As Range
    
    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Activate
        If sheet.Name <> "Raw" Then
            For Each i In Range("A1").CurrentRegion
                i = Replace(i, "#N/A N/A", "-")
                i = Replace(i, "#N/A Field Not Applicable", "-")
                i = Replace(i, "#N/A Real Time", "-")
                i = Replace(i, "(-/-/-)", "-")
                If i = "" Then i = "-"
            Next i
            Range("A1").CurrentRegion.Columns.AutoFit
            Range("A1").CurrentRegion.Rows.AutoFit
            Range("A1").Select
        End If
    Next sheet
        
End Sub
