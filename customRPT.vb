Imports Microsoft.Office.Interop

Public Class customRPT

    Public Sub dump2TXT(ByRef myXLS3d As Object, ByVal numRows As Long, rArgs As reportingArgs)
        Dim F As New Collection
        F = rArgs.someColl

        Dim hugE$ = ""
        Dim roW As Long = 0
        Dim coL As Integer = 0

        Dim tempL$ = ""

        For coL = 0 To F.Count - 1
            tempL += F(coL + 1)
            If coL = F.Count - 1 Then tempL += vbCrLf Else tempL += ","
        Next
        hugE += tempL

        For roW = 0 To numRows - 1
            tempL = ""
            For coL = 0 To F.Count - 1
                tempL += csvObj(myXLS3d(roW, coL))
                If coL = F.Count - 1 Then tempL += vbCrLf Else tempL += ","
            Next
            hugE += tempL
        Next

        Dim fileN$ = rArgs.s1

        Call streamWriterTxt(fileN, hugE)
        hugE = ""

    End Sub

    Private Function csvObj(ByRef a$) As String
        If a = "0" Or a = "" Or Val(a) > 0 Then
            csvObj = a
        Else
            a = Replace(Replace(a, Chr(34), ""), ",", "_")
            csvObj = Chr(34) + a + Chr(34)
        End If
    End Function
    Public Sub dump2XLS(ByRef myXLS3d As Object, ByVal numRows As Long, rArgs As reportingArgs, Optional dontFreeze As Boolean = False, Optional ByVal doPivot As Boolean = False)
        Dim F As New Collection
        F = rArgs.someColl


        If rArgs.booL1 = False Then
            dump2TXT(myXLS3d, numRows, rArgs)
            Exit Sub
        End If

        Dim appXL As New Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet

        appXL.Visible = True

        xlWB = appXL.Workbooks.Add
        xlWS = xlWB.Sheets(1)

        xlWS.Activate()

        Dim colNUM As Integer = 0

        'column titles
        For Each field In F
            colNUM += 1
            xlWS.Cells(1, colNUM) = field
        Next

        Dim startRow$ = "A2"
        'MainUI.addLOG("Copying 3D to Excel")

        xlWS.Range(startRow + ":" + xlsColName(F.Count) + Trim(Str(numRows + 1))).Value = myXLS3d

        xlWB.RefreshAll()
        xlWS.Columns.AutoFit()



        If doPivot = True Then
            xlWS.Name = "RAW_DATA"
            xlWS = xlWB.Sheets.Add
            xlWS.Name = "Pivot"
            xlWS.Activate()
            Call xlWSdoPivot(xlWS, numRows + 1, F, "TF", appXL)
        End If

        Dim numCopies As Integer = 0

        Dim a$ = ""

        'xlWS.Rows("1:1").Select


        appXL.DisplayAlerts = True
        xlWB.SaveAs2(rArgs.s1, FileFormat:=51)

        xlWS = Nothing
        xlWB = Nothing
        appXL = Nothing
        myXLS3d = Nothing
        GC.Collect()

    End Sub





    Public Sub xlWSdoPivot(ByRef xlWS As Excel.Worksheet, ByVal rowNum As Long, ByRef fieldS As Collection, tData$, ByRef appXL As Excel.Application)

        '        Call xlWSdoPivot(xlWS, numRows, F, "TF", appXL)

        xlWS.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="RAW_DATA!R1C1:R" + Trim(Str(rowNum)) + "C" + Trim(Str(fieldS.Count)), TableDestination:=xlWS.Range("A5"), TableName:="ResultsSummary")

        xlWS.Select()

        Dim PT As Excel.PivotTable = xlWS.PivotTables(1)

        'GoTo skipToHere

        Select Case tData$
            Case "TF"
                PT.AddDataField(PT.PivotFields("#_ISSUES"))

                Dim rPF = CType(PT.PivotFields("APP_NAME"), Microsoft.Office.Interop.Excel.PivotField)
                rPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField

                Dim cPF = CType(PT.PivotFields("CATEGORY"), Microsoft.Office.Interop.Excel.PivotField)
                cPF.Orientation = Excel.XlPivotFieldOrientation.xlColumnField


        End Select

skipToHere:
        PT.RowGrand = True
        PT.TableStyle2 = "PivotStyleDark2"
        xlWS.Columns.AutoFit()

    End Sub
End Class

Public Class reportingArgs
    Public rptName$
    Public booL1 As Boolean = False
    Public booL2 As Boolean = False
    Public s1$ = ""
    Public s2$ = ""
    Public s3$ = ""
    Public numeriC As Integer = 0
    Public numeriC2 As Integer = 0
    Public numeriC3 As Integer = 0
    Public someColl As Collection

End Class




