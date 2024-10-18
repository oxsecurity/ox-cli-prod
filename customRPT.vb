Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

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

        Dim tData$ = "TF"
        If Len(rArgs.s3) Then tData = rArgs.s3

        If doPivot = True Then
            xlWS.Name = "RAW_DATA"
            xlWS = xlWB.Sheets.Add
            xlWS.Name = "Pivot"
            xlWS.Activate()
            Call xlWSdoPivot(xlWS, numRows + 1, F, tData, appXL)
        End If

        Dim numCopies As Integer = 0

        Dim a$ = ""

        'xlWS.Rows("1:1").Select


        appXL.DisplayAlerts = False
        xlWB.SaveAs2(rArgs.s1, FileFormat:=51)

        xlWS = Nothing
        xlWB = Nothing
        appXL = Nothing
        myXLS3d = Nothing
        GC.Collect()

    End Sub

    Public Sub doROIrpt(ByRef rArgs As reportingArgs)
        ' With rAR
        '.s1 = fileN
        '.s2 = Path.Combine(ogDir, "roi_tags_template.xlsx")
        '
        '   If IO.File.Exists(.s2) = False Then
        '       Console.WriteLine("ERROR: Unable to find file " + .s2 + " - aborting")
        '       End
        'End If
        '
        '    .someColl = New Collection
        '    .someColl2 = New Collection
        '    For Each S In allSFs
        '         .someColl.Add(S.shortName)
        ''         .someColl2.Add(S.numOccurrences)
        '         If InStr(S.shortName, " Business Priority") > 0 Then bizPRI += S.numOccurrences
        '     Next
        '     .s3 = bizPRI.ToString
        '     .numeriC = numR
        '     .numeriC2 = numE
        '     .numeriC3 = numD
        ' End With

        Dim appXL As New Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet

        xlWB = appXL.Workbooks.Open(rArgs.s2)
        xlWS = xlWB.ActiveSheet

        appXL.Visible = True

        Dim xlRow As Integer = 5 ' row starts at 5, columns B and F have descript, columns C and G hold count of SFs
        Dim K As Integer


        For K = xlRow To 25
            Dim sF$ = xlGet(xlWS, "B" + K.ToString)
            Dim ndX = grpNDX(rArgs.someColl, sF, False)

            If ndX > 0 Then
                xlWS.Cells(K, 3) = rArgs.someColl2(ndX)
            End If

            If sF$ = "Business Priority" Then xlWS.Cells(K, 3) = Val(rArgs.s3)

            Dim sG$ = xlGet(xlWS, "F" + K.ToString)
            Dim ndX2 = grpNDX(rArgs.someColl, sG, False)

            If ndX2 > 0 Then
                xlWS.Cells(K, 7) = rArgs.someColl2(ndX2)
            End If

        Next

        appXL.DisplayAlerts = False

        Console.WriteLine("Saving " + rArgs.s1)
        xlWB.SaveAs2(rArgs.s1, FileFormat:=51)


        xlWS = Nothing
        xlWB = Nothing
        appXL = Nothing
        GC.Collect()
    End Sub



    Public Sub xlWSdoPivot(ByRef xlWS As Excel.Worksheet, ByVal rowNum As Long, ByRef fieldS As Collection, tData$, ByRef appXL As Excel.Application)

        '        Call xlWSdoPivot(xlWS, numRows, F, "TF", appXL)

        xlWS.PivotTableWizard(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="RAW_DATA!R1C1:R" + Trim(Str(rowNum)) + "C" + Trim(Str(fieldS.Count)), TableDestination:=xlWS.Range("A5"), TableName:="ResultsSummary")

        xlWS.Select()

        Dim PT As Excel.PivotTable = xlWS.PivotTables(1)

        'GoTo skipToHere

        Console.WriteLine("Pivot created with pattern: " + tData)

        Select Case tData$
            Case "TF"
                PT.AddDataField(PT.PivotFields("#_ISSUES"))

                Dim rPF = CType(PT.PivotFields("APP_NAME"), Microsoft.Office.Interop.Excel.PivotField)
                rPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField

                Dim cPF = CType(PT.PivotFields("CATEGORY"), Microsoft.Office.Interop.Excel.PivotField)
                cPF.Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            Case "TAGS"
                PT.AddDataField(PT.PivotFields("#_TAGS"))

                'Dim rPF = CType(PT.PivotFields("APP_NAME"), Microsoft.Office.Interop.Excel.PivotField)
                'rPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField

                Dim cPF = CType(PT.PivotFields("APP_TAG"), Microsoft.Office.Interop.Excel.PivotField)
                cPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField


            Case "IR"
                PT.AddDataField(PT.PivotFields("#_ISSUES"))
                'dF.Orientation=Excel.XlPivotFieldOrientation.

                Dim rPF = CType(PT.PivotFields("CATEGORY"), Microsoft.Office.Interop.Excel.PivotField)
                rPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField

                Dim cPF = CType(PT.PivotFields("SEVERITY"), Microsoft.Office.Interop.Excel.PivotField)
                cPF.Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            Case "SF"
                PT.AddDataField(PT.PivotFields("#_SF"))

                Dim rPF = CType(PT.PivotFields("SF_NAME"), Microsoft.Office.Interop.Excel.PivotField)
                rPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField

                Dim cPF = CType(PT.PivotFields("RED"), Microsoft.Office.Interop.Excel.PivotField)
                cPF.Orientation = Excel.XlPivotFieldOrientation.xlColumnField

            Case "DEV"

                PT.AddDataField(PT.PivotFields("#_30"))
                PT.AddDataField(PT.PivotFields("#_90"))
                PT.AddDataField(PT.PivotFields("#_180"))
                PT.AddDataField(PT.PivotFields("#_YEAR"))

                'Dim rPF = CType(PT.PivotFields("SF_NAME"), Microsoft.Office.Interop.Excel.PivotField)
                'rPF.Orientation = Excel.XlPivotFieldOrientation.xlRowField

                'Dim cPF = CType(PT.PivotFields("RED"), Microsoft.Office.Interop.Excel.PivotField)
                'cPF.Orientation = Excel.XlPivotFieldOrientation.xlColumnField


        End Select

skipToHere:
        PT.RowGrand = True
        PT.TableStyle2 = "PivotStyleDark2"
        xlWS.Columns.AutoFit()

    End Sub

    Public Function xlGet(xlWS As Excel.Worksheet, range$) As String
        Dim a$ = ""
        If IsNothing(xlWS.Range(range).Value) = True Then
            a = ""
        Else
            a = xlWS.Range(range).Value.ToString
        End If
        Return a
    End Function

    Public Function pullXLSfields(rAR As reportingArgs, inputF$) As reportingArgs
        Dim appXL As New Excel.Application
        Dim xlWB As Excel.Workbook
        Dim xlWS As Excel.Worksheet

        xlWB = appXL.Workbooks.Open(inputF)
        xlWS = appXL.ActiveSheet

        Console.WriteLine("Analyzing worksheet " + xlWS.Name)

        Dim roW As Integer = 2
        Dim a$ = ""
        If IsNothing(xlWS.Range(rAR.s1 + Trim(Str(roW))).Value) = True Then
            a = ""
        Else
            a = xlWS.Range(rAR.s1 + Trim(Str(roW))).Value.ToString
        End If

        Dim b$ = ""
        If IsNothing(xlWS.Range(rAR.s3 + Trim(Str(roW))).Value) = True Then
            b = ""
        Else
            b = xlWS.Range(rAR.s3 + Trim(Str(roW))).Value.ToString
        End If

        Do Until a$ = ""
            'Console.WriteLine(a + " - " + b)
            rAR.someColl.Add(Trim(a))
            rAR.someColl2.Add(Trim(b))

            roW += 1
            a$ = ""
            b$ = ""
            If IsNothing(xlWS.Range(rAR.s1 + Trim(Str(roW))).Value) = False Then a = xlWS.Range(rAR.s1 + Trim(Str(roW))).Value.ToString
            If IsNothing(xlWS.Range(rAR.s3 + Trim(Str(roW))).Value) = False Then b = xlWS.Range(rAR.s3 + Trim(Str(roW))).Value.ToString
        Loop

        xlWB.Close()
        appXL.Quit()

        xlWS = Nothing
        xlWB = Nothing
        appXL = Nothing
        GC.Collect()

        Return rAR
    End Function
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
    Public someColl2 As Collection
End Class




