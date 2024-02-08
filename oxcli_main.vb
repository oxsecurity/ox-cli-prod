Imports System.Text.RegularExpressions
Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Reflection.Emit

Module Program
    Public aTimer As New System.Timers.Timer
    Public OX As oxWrapper

    Public currOffset = 0
    Public issueLimit = 1000

    Public ogDir$
    Public pyDir$
    Public osType$

    Public numResponseFiles As Integer = 0

    Sub Main(args As String())
        If UBound(args) = -1 Then
            Console.WriteLine("You must enter a command. Try 'help'.")
            End
        End If

        Dim actioN$ = args(0)
        Console.WriteLine("ACTION: " + actioN)

        ' this will generate errors if no python folder exists
        ogDir$ = FileSystem.CurDir
        ChDir("python")
        pyDir$ = FileSystem.CurDir
        ChDir(ogDir)


        If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) = True Then osType = "Windows"
        If RuntimeInformation.IsOSPlatform(OSPlatform.OSX) = True Then osType = "MacOSX"
        If RuntimeInformation.IsOSPlatform(OSPlatform.Linux) = True Then osType = "Linux"

        Console.WriteLine("Detecting " + osType + " environment")

        Select Case LCase(actioN)
            Case "help"
                Console.WriteLine(vbCrLf + "Usage  : ACTION --PARAM1 value --PARAM2 value                        >>>> Actions and parameters are not case sensitive")
                Console.WriteLine("Example: oxcli getjson --api getapplications --file applist.json     >>>> Performs 'getjson' action using param values of 'api' and 'file'")
                Console.WriteLine("=======================================================================================================================================================")
                Console.WriteLine(fLine("help", "shows this list of commands"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("checkme", "self-inspection of environment"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("setenv", "sets environment vars for Python API calls"))
                Console.WriteLine(fLine("", "[REQUIRED] --KEY (OX API Key)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --KEY (OX API Key defaults to https://api.cloud.ox.security/api/apollo-gateway)"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("policycsv", "create CSV of policies (requires that policy JSON files are saved manually)"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("getjson", "uses python engine to pull API JSON args"))
                Console.WriteLine(fLine("", "[REQUIRED] --API (name of API)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --file (output filename)"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("issuesxls", "retrieves issues using python engine args and creates pivot Excel doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name of XLS file to create)"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("issuescsv", "retrieves issues Using python engine args And creates CSV doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name Of CSV file To create)"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("addtag", "adds a New tag - name must be unique"))
                Console.WriteLine(fLine("", "[REQUIRED] --NAME (name Of the tag)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --DISPLAY (display name, defaults To same As --NAME), --TYPE (defaults To 'simple' - recommend default)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("edittags", "Loop through apps and Add/Remove tags using string and/or regex match"))
                Console.WriteLine(fLine("", "[REQUIRED] --ADDTAG (tag displayname) OR --REMTAG (name)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --STR (app contains string), --REGEX (app matches regex), --COMMIT true (otherwise will only preview)"))
                Console.WriteLine(fLine("", "[TEST]     --MATCH (submit test app name) - will confirm string and/or regex match without looping through apps"))
                Console.WriteLine("=======================================================================================================================================================")
                End
            Case "checkme"
                ' removed folder structure due to nuances between different versions of *NIX (tested good on MacOS & Windows only, failed on DEBIAN BOOKWORM and BULLSEYE)

                Dim findFile$ = Path.Combine(ogDir, "Newtonsoft.Json.dll")

                Console.WriteLine("File System check - filesystem.curdir: " + ogDir)
                Console.WriteLine("Path.GetFullPath(Directory.GetCurrentDirectory()= " + Path.GetFullPath(Directory.GetCurrentDirectory()))

                If File.Exists(findFile) = True Then
                    Console.WriteLine("OXcli files present")
                Else
                    Console.WriteLine("OX dependencies missing - check folder contents or download new version")
                    End
                End If

                FileSystem.ChDir(pyDir)

                If File.Exists(".env") Then
                    Console.WriteLine("Environment file exists - credentials not verified")
                Else
                    Console.WriteLine("Environment file (.env) is not present and is needed for credentials")
                End If


                Console.WriteLine("Python directory:  " + pyDir)
                If File.Exists("python_examp.py") Then
                    Console.WriteLine("Python executable exists")
                Else
                    Console.WriteLine("Python script to call APIs must be present - obtain python folder that accompanies this DOTNET executable")
                    End
                End If
                '                Console.WriteLine("Path.GetFullPath(Directory.GetCurrentDirectory()= " + Path.GetFullPath(Directory.GetCurrentDirectory()))

                FileSystem.ChDir(ogDir)

                Console.WriteLine("Changing back to parent folder - " + ogDir)
                If File.Exists(findFile) = True Then
                    Console.WriteLine("Environment check complete")
                Else
                    Console.WriteLine("Unable to revert to previous folder")
                    End
                End If

                End

            Case "setenv"
                Dim urL$ = argValue("url", args)
                Dim apiKey$ = argValue("key", args)
                If apiKey = "" Then
                    Console.WriteLine("To set environment, OXCLI needs the URL and API KEY. Submit using --URL and --KEY params")
                End If
                If urL = "" Then
                    urL = "https://api.cloud.ox.security/api/apollo-gateway"
                    Console.WriteLine("Will default to https://api.cloud.ox.security/api/apollo-gateway")
                End If
                Dim newEfile$ = "oxUrl='" + urL + "'" + Chr(13) + "oxKey='" + apiKey + "'"

                FileSystem.ChDir(pyDir)
                'Console.WriteLine("Checking for " + Path.Combine(pyDir, ".env"))
                If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) <> True Then
                    If File.Exists(Path.Combine(pyDir, ".env")) = True Then
                        Console.WriteLine("On *NIX systems, either delete /python/.env and run this command again or edit your existing /python/.env file to keep only the latest values")
                    End If
                Else
                    safeKILL(".env")
                End If
                Call streamWriterTxt(".env", newEfile)
                Console.WriteLine("New environment variables set for " + urL)
                FileSystem.ChDir(ogDir)
                End

            Case "dvusage_beta_filters"
                Dim fileN As Collection = New Collection
                fileN.Add("DV0523F")
                fileN.Add("DV0623F")
                fileN.Add("DV0723F")
                fileN.Add("DV0823F")
                fileN.Add("DV0923F")
                fileN.Add("DV1023F")
                fileN.Add("DV1123F")
                fileN.Add("DV1223F")
                fileN.Add("DV0124F")
                OX = New oxWrapper("", "")

                Dim fullCSV$ = ""

                For Each J In fileN
                    Dim jsoN$ = streamReaderTxt(J)
                    Dim oxFiltersForMonth As oxUserLogFilter = OX.getUserFilterEntries(jsoN)
                    With oxFiltersForMonth
                        Console.WriteLine(J + " logTypes: " + .logTypes.Count.ToString + " logNames: " + .logNames.Count.ToString + " emails: " + .userEmails.Count.ToString)

                        Dim newS$ = "2023-" + Mid(J, 3, 2) + ",ACTIVITY_TYPE,"
                        For Each L In .logTypes
                            If L.label <> "" Then fullCSV += newS + L.label + "," + L.count.ToString + vbCrLf
                        Next
                        newS$ = "2023-" + Mid(J, 3, 2) + ",ACTIVITY,"
                        For Each A In .logNames
                            If A.label <> "" Then fullCSV += newS + A.label + "," + A.count.ToString + vbCrLf
                        Next
                        newS$ = "2023-" + Mid(J, 3, 2) + ",USER_ACTIVITY,"
                        For Each A In .userEmails
                            If InStr(A.label, "double") > 0 Then fullCSV += newS + A.label + "," + A.count.ToString + vbCrLf
                        Next

                    End With
                Next
                streamWriterTxt("dvusagecsv.csv", fullCSV)
                End

            Case "dvusage_beta_actions"
                ' This does not work as requests from browser are paged 50x
                Console.WriteLine("Action requires API not yet exposed")
                Dim fileN As Collection = New Collection
                fileN.Add("DV0523A")
                fileN.Add("DV0623A")
                fileN.Add("DV0723A")
                fileN.Add("DV0823A")
                fileN.Add("DV0923A")
                fileN.Add("DV1023A")
                fileN.Add("DV1123A")
                fileN.Add("DV1223A")
                fileN.Add("DV0124A")
                Dim fullList As List(Of oxUserLogEntry) = New List(Of oxUserLogEntry)
                OX = New oxWrapper("", "")

                For Each J In fileN
                    'If File.Exists(J) = False Then Console.WriteLine("Does not exist")
                    Dim jsoN$ = streamReaderTxt(J)
                    'Console.WriteLine(jsoN)
                    Dim monthList As List(Of oxUserLogEntry) = New List(Of oxUserLogEntry)
                    monthList = OX.getUserLogEntries(jsoN)
                    For Each A In monthList
                        fullList.Add(A)

                    Next
                    Console.WriteLine("Items in this list: " + monthList.Count.ToString)
                    Console.WriteLine("Total number items: " + fullList.Count.ToString)

                Next
                End

            Case "policycsv"
                Call policyCSV()
                End
            Case "addtag"
                Call addTag(argValue("name", args), argValue("display", args), argValue("type", args))
                End
            Case "edittags"
                Call editTags(args)
                End
            Case "getjson"
                Dim apiCall$ = argValue("api", args)
                Dim fileN$ = argValue("file", args)
                If Len(apiCall) = 0 Then
                    Console.WriteLine("This command's parameters:  getjson --api apiname --file 'file name.json'" + vbCrLf + "api  [req] : the name of the API call (action getAPIs)" + vbCrLf + "file [opt]: the output filename to dump JSON")
                    End
                End If

                If Len(fileN) = 0 Then
                    Console.WriteLine(setUpAPICall(apiCall, "", True))
                Else
                    Call setUpAPICall(apiCall, fileN)
                    Console.WriteLine("File created: " + Path.Combine(ogDir, fileN))
                End If

                End

            Case "issuesxls", "issuescsv"

                If LCase(actioN) = "issuesxls" And osType <> "Windows" Then
                    Console.WriteLine("This command will only work on a Windows machine with Excel locally installed")
                    End
                End If
                Dim toFilename$ = argValue("file", args)
                If Len(toFilename) = 0 Then
                    Console.WriteLine("This command's parameters:  issuesxls OR issuescsv --file 'filename.xlsx'" + vbCrLf + "file : the output Excel filename.")
                    Console.WriteLine("You must specify a filename for the CSV or Excel .xlsx.")
                    End
                End If
                Call getAllIssues()
                Dim allIssues As List(Of issueS)
                allIssues = buildIssues("getIssues.json", numResponseFiles - 1)
                toFilename = Path.Combine(ogDir, toFilename)

                If LCase(actioN) = "issuesxls" Then
                    Call issueRpt(allIssues, toFilename)
                Else
                    Call issueCSV(allIssues, toFilename)
                End If
        End Select

        End



    End Sub


    Public Function addTag(tagName$, Optional ByVal dName$ = "", Optional ByVal tType$ = "simple") As String
        addTag = "" ' returns empty if unsuccessful otherwise tagid of new tag
        If dName = "" Then dName = tagName
        If tType = "" Then tType = "simple"

        Console.WriteLine("Adding tag:")
        Call setAddTagVars(tagName, dName, tType)
        Dim jSon$ = setUpAPICall("addTag",, True)
        addTag = OX.getTagId(jSon)
        If addTag = "" Then
            Console.WriteLine("ERROR: Could not add tag - return JSON=" + vbCrLf + jSon)
        Else
            Console.WriteLine("New Tag: " + tagName + " >> TagID: " + addTag)
        End If
    End Function
    Public Sub editTags(args() As String)
        Dim doRegExMatch As Boolean = False
        Dim toMatch$ = argValue("match", args)
        Dim regX As Regex
        Dim regXmatch As Match
        Dim matchStr$ = argValue("str", args)
        Dim testingOnly As Boolean = False

        Dim addedTag$ = argValue("addtag", args)
        Dim remTag$ = argValue("remtag", args)

        Dim newModTag As editTagsRequestVARS = New editTagsRequestVARS

        Dim commitChanges As Boolean = False

        If LCase(argValue("commit", args)) = "true" Then commitChanges = True

        If Len(toMatch) > 0 Then
            testingOnly = True
            Console.WriteLine("STR=" + matchStr + " TOMATCH=" + toMatch)
        End If

        If Len(argValue("regex", args)) Then
            doRegExMatch = True
            regX = New Regex(argValue("regex", args))
            Console.WriteLine("Performing REGEX matching using " + qT(argValue("regex", args)))
            If testingOnly Then
                Console.WriteLine("Testing match on " + qT(toMatch) + " using Regular Expression: " + qT(argValue("regex", args)))
                regXmatch = regX.Match(toMatch)
                Console.WriteLine("REGX_MATCH: " + CStr(regXmatch.Success) + "  VALUE: " + regXmatch.Value)
            End If
        End If

        Dim doStringMatch As Boolean = False
        If Len(matchStr) Then
            doStringMatch = True
            Console.WriteLine("Performing STRING matching using " + qT(matchStr))
            If testingOnly Then
                Console.WriteLine("Testing match on " + qT(toMatch) + " by looking for string: " + qT(matchStr))
                If InStr(toMatch, matchStr, CompareMethod.Text) Then
                    Console.WriteLine("STR_MATCH: True ")
                Else
                    Console.WriteLine("MATCH: False ")
                End If
            End If
        End If

        If testingOnly = True Then
            End
        End If

        Console.WriteLine("Getting applications")

        Dim appsWithTag As Integer = 0
        Dim allAppsWithTag As Integer = 0

        Dim allApps As List(Of oxAppshort) = getAppListShort()
        Console.WriteLine("# of Applications: " + allApps.Count.ToString)
        Dim allTags As List(Of oxTag) = getAllTags()
        Console.WriteLine("# of Tags: " + allTags.Count.ToString)

        Dim tId$ = OX.returnTagId(addedTag, allTags)
        If tId = "" Then
            Console.WriteLine("This tag must be created before it can be applied")
            If commitChanges = True Then tId = addTag(addedTag)
            If tId = "" And commitChanges = True Then
                Console.WriteLine("ERROR:Could not add this tag - exiting without changes")
                End
            Else
                newModTag.addedTagsIds.Add(tId)
            End If
        Else
            Console.WriteLine("Found TAG: " + tId)
            newModTag.addedTagsIds.Add(tId)
        End If


        ' for now - remTAG needs to be accounted for.. Is API smart enough to ignore REMOVE commands when TAG doesnt exist in first place?
        ' May need to separate ADD and REMOVE and separate operations, although API appears to account for both across multiple apps with a single call

        If newModTag.addedTagsIds.Count + newModTag.removedTagsIds.Count = 0 Then
            Console.WriteLine("You must either add or remove a tag for this operation to run, using --addtag and/or --remtag")
            End
        End If


        Console.WriteLine(vbCrLf + "These applications to receive new tags:" + vbCrLf)
        Console.WriteLine(fLine("Application Name" + spaces(44), "Link" + spaces(76) + "# Tags"))
        Console.WriteLine(fLine("================" + spaces(44), "====" + spaces(76) + "======"))
        For Each app In allApps
            Dim addTag As Boolean = True

            If doRegExMatch = True And addTag = True Then
                regXmatch = regX.Match(app.appName)
                If regXmatch.Success = False Then addTag = False
            End If

            If doStringMatch = True And addTag = True Then
                If InStr(app.appName, matchStr, CompareMethod.Text) = 0 Then addTag = False
            End If

            If app.tagExist(, addedTag) Then
                allAppsWithTag += 1

                If addTag = True Then
                    appsWithTag += 1
                    addTag = False
                End If
            End If

            ' not yet - later ' If app.tagExist(, remTag) Then

            If addTag Then
                Console.WriteLine(fLine(app.appName + spaces(60 - Len(app.appName)), app.link + spaces(80 - Len(app.link)) + app.tags.Count.ToString))
                newModTag.appIds.Add(app.appId)
            End If
        Next

        Console.WriteLine(vbCrLf)

        Console.WriteLine("# Matching Apps with Tag:  " + appsWithTag.ToString)
        Console.WriteLine("# Total Apps with Tag:     " + allAppsWithTag.ToString)

        Console.WriteLine("# Matching Apps to Modify: " + newModTag.appIds.Count.ToString + vbCrLf)

        ' setEditTagsVarsRequests
        If newModTag.appIds.Count Then
            Call setEditTagsVarsRequests(newModTag)
            Console.WriteLine("This action will modify tags of " + newModTag.appIds.Count.ToString + " apps")

            If LCase(argValue("commit", args)) = "true" Then
                Call setUpAPICall("modifyAppsTags")
                Console.WriteLine("Action completed")
            Else
                Console.WriteLine("In order to commit these changes, call command with '--commit true'")
            End If
        Else
            Console.WriteLine("Exiting without changes")
        End If
        End
    End Sub




    ' these get* funcs also need to move to the wrapper

    Public Function getAllTags() As List(Of oxTag)
        getAllTags = New List(Of oxTag)
        Dim fName$ = "getAllTags.json"

        Dim jsoN$ = ""
        jsoN = setUpAPICall("getAllTags",, True)

        If Len(jsoN) Then getAllTags = OX.getAllTags(jsoN)
    End Function



    Public Function getAppListShort() As List(Of oxAppshort)
        getAppListShort = New List(Of oxAppshort)
        Dim jsoN$ = ""
        jsoN = setUpAPICall("getAppsShort",, True)

        getAppListShort = OX.getAppInfoShort(jsoN)
    End Function

    Public Function getAllIssues() As Integer
        numResponseFiles = 0

        Dim fileName$ = "getIssues.json"
        Call setGetIssuesVars(0)

        Console.WriteLine("Pulling first page of issues")

        Call setUpAPICall("getIssues", Replace(fileName, ".json", "0.json"))

        Dim respIssue As listIssues = New listIssues
        respIssue = OX.getListIssues(Replace(fileName, ".json", "0.json"))
        With respIssue
            Console.WriteLine("Total Issues: " + .totalIssues.ToString)
            'Console.WriteLine("Filtered Issues: " + .totalFilteredIssues.ToString)
            'Console.WriteLine("Offset: " + .offset.ToString)
        End With

        numResponseFiles = 1
        'this was the first file

        Do Until numResponseFiles = numIssueRequests(respIssue.totalFilteredIssues)
            numResponseFiles += 1
            Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + numIssueRequests(respIssue.totalFilteredIssues).ToString + " requests")
            Call setGetIssuesVars(issueLimit * numResponseFiles)
            Call setUpAPICall("getIssues", Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
            Console.SetCursorPosition(0, Console.CursorTop - 1)
            Console.WriteLine(spaces(150))
            Console.SetCursorPosition(0, Console.CursorTop - 2)

        Loop

        Call setGetIssuesVars(0)

        Return numResponseFiles
    End Function

    ' Consider putting these in wrapper <summary>
    ' writing new vars files

    Public Sub setAddTagVars(name$, displayName$, tagType$)
        Dim newTag As newTagRequestVARS = New newTagRequestVARS(displayName, name, tagType)
        OX = New oxWrapper("", "")
        Dim newJson$ = OX.jsonGetNewTagVars(newTag)
        Console.WriteLine(newJson)
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, "addTag.variables.json"))

    End Sub



    Public Sub setGetIssuesVars(offSet As Integer)
        Dim newIssueVar As issueRequestVARS = New issueRequestVARS
        With newIssueVar
            .dateRange.from = 1
            .dateRange.to = dateToJS(Now)
            .getIssuesInput.limit = issueLimit
            .getIssuesInput.offset = offSet
        End With
        Dim newJson$ = ""
        ' this is sloppy pls figure out why you did this, this way
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetIssuesVars(newIssueVar)
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, "getIssues.variables.json"))

    End Sub

    Public Sub setEditTagsVarsRequests(evReq As editTagsRequestVARS)
        Dim newJson$ = ""
        ' more sloppy - make OX global for main
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetEditTagsVars(evReq)
        'Console.WriteLine(newJson)
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, "modifyAppsTags.variables.json"))
    End Sub


    Public Function numIssueRequests(totalIssues As Long) As Long
        numIssueRequests = 0

        Dim tlCalled As Integer = 0

        Do Until tlCalled >= totalIssues
            numIssueRequests += 1
            tlCalled += issueLimit
        Loop

    End Function

    Public Function setUpAPICall(apiCall$, Optional ByVal fileN$ = "", Optional ByVal showJSON As Boolean = False) As String
        setUpAPICall = ""
        OX = New oxWrapper("", "")
        'Console.WriteLine("Retrieving JSON From OX API: " + apiCall)

        Dim getFile$ = apiCall + "_response.json"
        If fileN <> "" Then fileN = fileN

        safeKILL(getFile)
        If Len(fileN) Then safeKILL(fileN)

        Console.WriteLine("Executing Python request for '" + apiCall + "'")

        Dim succesS As Boolean = OX.getJSON(apiCall)

        If succesS = False Then
            Console.WriteLine("Check the underlying Python connector, make sure Python is in the path. Try 'python python_examp.py " + apiCall)
            End
        Else
            If Len(fileN) Then
                FileCopy(Path.Combine(pyDir, getFile), Path.Combine(ogDir, fileN))
            Else
                If showJSON = True Then setUpAPICall = streamReaderTxt(Path.Combine(pyDir, getFile)) ' Console.WriteLine(streamReaderTxt(getFile))
            End If
        End If
    End Function

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
            Console.WriteLine("Deserializing JSON " + cFile)
            tempIssues = OX.returnIssues(streamReaderTxt(cFile))

            If numFiles = 0 Then
                allIssues = tempIssues
            Else
                ' is there a better way? test 
                For Each T In tempIssues
                    allIssues.Add(T)
                Next
            End If
        Next


        Console.WriteLine("Total # of files (" + issueLimit.ToString + " issues per file): " + (numFiles + 1).ToString)
        Console.WriteLine("# of Issues: " + allIssues.Count.ToString)
        Return allIssues
    End Function

    Private Sub issueRpt(allIssues As List(Of issueS), fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        rAR.someColl = New Collection
        rAR.s1 = fileN

        Call safeKILL(fileN)

        'Console.WriteLine("Building output from results")

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

    Private Sub issueCSV(allIssues As List(Of issueS), fileN$)
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

        Console.WriteLine("Throwing object of " + (allIssues.Count * 8).ToString + " elements into CSV..")
        rAR.booL1 = True ' create XLS not CSV

        Call newRpt.dump2TXT(xls3d, allIssues.Count, rAR)

        End
    End Sub


End Module
