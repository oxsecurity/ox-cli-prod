Imports System
Imports System.IO

Module Program
    Public aTimer As New System.Timers.Timer
    Public OX As oxWrapper

    Public currOffset = 0
    Public issueLimit = 30

    Public numResponseFiles As Integer = 0

    Sub Main(args As String())
        If UBound(args) = -1 Then
            Console.WriteLine("You must enter a command. Try 'help'.")
            End
        End If

        Dim actioN$ = args(0)
        Console.WriteLine("ACTION: " + actioN)



        Select Case LCase(actioN)
            Case "help"
                Console.WriteLine(fLine("help", "shows this list of commands"))
                Console.WriteLine(fLine("policycsv", "create CSV of policies (requires that policy JSON files are saved manually)"))
                Console.WriteLine(fLine("getjson", "uses python engine to pull API JSON args --API (name of API), --file (output filename)"))
                Console.WriteLine(fLine("issuesxls", "retrieves issues using python engine args --FILE (name of XLS file to create)"))
                End

            Case "policycsv"
                Call policyCSV()
                End

            Case "getjson"
                Dim apiCall$ = argValue("api", args)
                Dim fileN$ = argValue("file", args)

                If Len(apiCall) = 0 Then
                    Console.WriteLine("This command's parameters:  getjson --api apiname --file 'file name.json'" + vbCrLf + "api  [req] : the name of the API call (action getAPIs)" + vbCrLf + "file [opt]: the output filename to dump JSON")
                    End
                End If

                If Len(fileN) = 0 Then
                    Call setUpAPICall(apiCall, "", True)
                Else
                    Console.WriteLine("File created: " + fileN)
                End If
                End

                    Case "issuesxls"
                Dim toFilename$ = argValue("file", args)
                If Len(toFilename) = 0 Then
                    Console.WriteLine("This command's parameters:  makexls --file 'filename.xlsx'" + vbCrLf + "file : the output Excel filename.")
                    Console.WriteLine("You must specify a filename for the Excel .xlsx.")
                    End
                End If
                Call getAllIssues()
                Dim allIssues As List(Of issueS)
                allIssues = buildIssues("getissues.json", numResponseFiles - 1)
                Call issueRpt(allIssues, CurDir() + "\" + toFilename)
        End Select

        End



    End Sub
    Public Function getAllIssues() As Integer
        numResponseFiles = 0

        Dim fileName$ = "getissues.json"
        Call setGetIssuesVars(0)

        Call setUpAPICall("getissues", Replace(fileName, ".json", "0.json"))

        Dim respIssue As listIssues = New listIssues
        respIssue = OX.getListIssues(Replace(fileName, ".json", "0.json"))
        With respIssue
            Console.WriteLine("Total Issues: " + .totalIssues.ToString)
            'Console.WriteLine("Filtered Issues: " + .totalFilteredIssues.ToString)
            'Console.WriteLine("Offset: " + .offset.ToString)
        End With

        'this was the first file

        Do Until numResponseFiles = numIssueRequests(respIssue.totalFilteredIssues)
            numResponseFiles += 1
            Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + numIssueRequests(respIssue.totalFilteredIssues).ToString + " requests")
            Call setGetIssuesVars(issueLimit * numResponseFiles)
            Call setUpAPICall("getissues", Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
            Console.SetCursorPosition(0, Console.CursorTop - 1)
            Console.WriteLine(spaces(150))
            Console.SetCursorPosition(0, Console.CursorTop - 2)

        Loop

        Call setGetIssuesVars(0)

        Return numResponseFiles
    End Function
    Public Sub setGetIssuesVars(offSet As Integer)
        Dim newIssueVar As issueRequestVARS = New issueRequestVARS
        With newIssueVar
            .dateRange.from = 1
            .dateRange.to = dateToJS(Now)
            .getIssuesInput.limit = issueLimit
            .getIssuesInput.offset = offSet
        End With
        Dim newJson$ = ""
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetIssuesVars(newIssueVar)
        Call saveJSONtoFile(newJson, "./python/getissues.variables.json")

    End Sub

    Public Function numIssueRequests(totalIssues As Long) As Long
        numIssueRequests = 0

        Dim tlCalled As Integer = 0

        Do Until tlCalled >= totalIssues
            numIssueRequests += 1
            tlCalled += issueLimit
        Loop

    End Function

    Public Sub setUpAPICall(apiCall$, Optional ByVal fileN$ = "", Optional ByVal showJSON As Boolean = False)
        OX = New oxWrapper("", "")
        'Console.WriteLine("Retrieving JSON From OX API: " + apiCall)

        Dim getFile$ = "./python/" + apiCall + "_response.json"
        safeKILL(getFile)
        If Len(fileN) Then safeKILL(fileN)

        Dim succesS As Boolean = OX.getJSON(apiCall)

        If succesS = False Then
            Console.WriteLine("Check the underlying Python connector, make sure Python is in the path. Try 'python ./python/python_examp.py ./python/" + apiCall)
            End
        Else
            If Len(fileN) Then
                FileCopy(getFile, fileN)
            Else
                If showJSON = True Then Console.WriteLine(streamReaderTxt(getFile))
            End If
        End If
    End Sub

    Public Sub policyCSV()

        Dim pWrap As policyWrapper = New policyWrapper

        Dim allPolicies As List(Of oxPolicy) = New List(Of oxPolicy)

        For Each P In pWrap.policyTypes
            Console.WriteLine(P)
            Dim nextPolicy As List(Of oxPolicy)
            nextPolicy = pWrap.loadPolicy(P)

            For Each nP In nextPolicy
                allPolicies.Add(nP)
            Next
        Next

        Console.WriteLine("WRITING POLICY.CSV...")

        Dim csvString$ = ""
        For Each policyRule In allPolicies
            With policyRule
                csvString += qT(.categorY) + "," + qT(.name) + "," + qT(.description) + "," + qT(.detailedDescription) + vbCrLf
            End With
        Next

        Call safeKILL("policies.csv")
        Call streamWriterTxt("policies.csv", csvString)
        End
    End Sub


    Public Function buildIssues(fileN$, numFiles As Integer) As List(Of issueS)
        Dim allIssues As List(Of issueS) = New List(Of issueS)

        Dim currFile As Integer
        Dim cFile$

        For currFile = 0 To numFiles
            cFile = Replace(fileN, ".json", currFile.ToString + ".json")
            Dim tempIssues As List(Of issueS) = New List(Of issueS)
            tempIssues = OX.returnIssues(streamReaderTxt(cFile))
            For Each T In tempIssues
                allIssues.Add(T)
            Next
        Next
        Console.WriteLine("Total # of files (" + issueLimit.ToString + " issues per file): " + numFiles.ToString)
        Console.WriteLine("# of Issues: " + allIssues.Count.ToString)
        Return allIssues
    End Function

    Private Sub issueRpt(allIssues As List(Of issueS), fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        rAR.someColl = New Collection
        rAR.s1 = fileN

        Call safeKILL(fileN)

        With rAR.someColl
            .Add("APP_NAME")
            .Add("PRIORITY")
            .Add("CATEGORY")
            .Add("POLICY")
            .Add("ISSUE_TITLE")
            .Add("SEVERITY")
            .Add("DATE_FOUND")
            .Add("#_OCCURRENCES")
            .Add("#_ISSUES")
        End With

        Dim xls3d(allIssues.Count - 1, 8) As Object

        Dim roW As Long = 0
        For Each aA In allIssues
            xls3d(roW, 0) = aA.app.name
            xls3d(roW, 1) = Math.Round(aA.app.businessPriority, 0)
            xls3d(roW, 2) = aA.category.name
            xls3d(roW, 3) = aA.policy.name
            xls3d(roW, 4) = aA.mainTitle
            xls3d(roW, 5) = aA.severity
            xls3d(roW, 6) = jStoDate(aA.created)
            xls3d(roW, 7) = aA.occurrences
            xls3d(roW, 8) = 1
            roW += 1
        Next

        Console.WriteLine("Throwing object of " + (allIssues.Count * 8).ToString + " elements at Excel..")
        rAR.booL1 = True ' create XLS not CSV

        Call newRpt.dump2XLS(xls3d, allIssues.Count, rAR,, True)

        End
    End Sub

End Module
