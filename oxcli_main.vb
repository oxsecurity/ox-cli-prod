Imports System.Text.RegularExpressions
Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Net.WebRequestMethods

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Imports System.ComponentModel
Imports System.Threading


Module Program
    Public aTimer As New System.Timers.Timer
    Public OX As oxWrapper

    Public currOffset = 0
    Public issueLimit = 1000

    Public ogDir$
    Public pyDir$
    Public cacheDir$
    Public osType$

    Public numResponseFiles As Integer = 0

    Public WithEvents issueCacheLoader As BackgroundWorker

    Public issuesCache As List(Of singleIssue)
    Public fileNames As List(Of String)
    Public currentlyLoading As Boolean

    Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

    '    Public Sub loadJSONissues() As 

    Sub Main(args As String())
        If UBound(args) = -1 Then
            Console.WriteLine("You must enter a command. Try 'help'.")
            End
        End If

        Dim actioN$ = args(0)
        Console.WriteLine("ACTION: " + actioN)

        ' this will generate errors if no python folder exists
        ogDir$ = FileSystem.CurDir

        Call System.IO.Directory.CreateDirectory("cache")
        ChDir("cache")
        cacheDir = FileSystem.CurDir

        ChDir(ogDir)

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

                Console.WriteLine(fLine("apptagsxls", "retrieves all APP TAGS and creates pivot Excel doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name of XLS file to create)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("issuesdetailed", "retrieves all issue data and creates pivot Excel doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name of XLS file to create)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --CACHE false  (by default, issues will be stored locally for caching)"))
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

                If IO.File.Exists(findFile) = True Then
                    Console.WriteLine("OXcli files present")
                Else
                    Console.WriteLine("OX dependencies missing - check folder contents or download new version")
                    End
                End If

                FileSystem.ChDir(pyDir)

                If IO.File.Exists(".env") Then
                    Console.WriteLine("Environment file exists - credentials not verified")
                Else
                    Console.WriteLine("Environment file (.env) is not present and is needed for credentials")
                End If


                Console.WriteLine("Python directory:  " + pyDir)
                If IO.File.Exists("python_examp.py") Then
                    Console.WriteLine("Python executable exists")
                Else
                    Console.WriteLine("Python script to call APIs must be present - obtain python folder that accompanies this DOTNET executable")
                    End
                End If
                '                Console.WriteLine("Path.GetFullPath(Directory.GetCurrentDirectory()= " + Path.GetFullPath(Directory.GetCurrentDirectory()))

                FileSystem.ChDir(ogDir)

                Console.WriteLine("Changing back to parent folder - " + ogDir)
                If IO.File.Exists(findFile) = True Then
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
                    If IO.File.Exists(Path.Combine(pyDir, ".env")) = True Then
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

                ' manual stuff
                fileN = New Collection
                fileN.Add("sbpoc.json")

                OX = New oxWrapper("", "")

                Dim fullCSV$ = ""

                Dim writeFileN$ = ""
                writeFileN = argValue("file", args)
                If writeFileN = "" Then writeFileN = "sbusagecsv.csv"

                Dim detailLvl$
                detailLvl = argValue("detail", args)

                If detailLvl = "" Then detailLvl = "summary"


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
                streamWriterTxt(writeFileN, fullCSV)
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

                ' manual stuff
                fileN = New Collection
                fileN.Add("sbpoc.json")

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

            Case "apptagsxls"
                If LCase(actioN) = "apptagsxls" And osType <> "Windows" Then
                    Console.WriteLine("This command will only work on a Windows machine with Excel locally installed")
                    End
                End If
                Dim toFilename$ = argValue("file", args)
                toFilename = Path.Combine(ogDir, toFilename)
                Dim allApps As List(Of oxAppshort) = New List(Of oxAppshort)
                allApps = getAppListShort()
                Console.WriteLine("# of Applications: " + allApps.Count.ToString)
                Call appTagRPT(allApps, toFilename)

            Case "issuesdetailed"
                ' If osType <> "Windows" Then
                'Console.WriteLine("This command will only work on a Windows machine with Excel locally installed")
                'nd
                'End If
                Dim toFilename$ = argValue("file", args)
                If Len(toFilename) = 0 Then
                    Console.WriteLine("This command's parameters:  issuesdetailed --file 'filename.xlsx'" + vbCrLf + "file : the output Excel filename.")
                    Console.WriteLine("You must specify a filename for the Excel .xlsx - if you do not have Excel, issues will still be stored unless --cache false")
                    End
                End If
                Dim allIssues As List(Of issueS)

                Dim loadList As Boolean = False
                If LCase(argValue("loadlist", args)) = "true" Then loadList = True

                If loadList = True Then
                    Call getAllIssues(True)
                Else
                    allIssues = loadCachedList()
                    If allIssues.Count Then GoTo gotCached
                End If

                If loadList = False Then Call getAllIssues(True) ' this means tried to load cache and nothing there
                allIssues = buildShortIssues("getIssuesShort.json", numResponseFiles - 1)

gotCached:
                toFilename = Path.Combine(ogDir, toFilename)

                fileNames = New List(Of String)

                ' building cache
                Dim numCache As Long = 0
                Dim numWrite As Long = 0
                For Each I In allIssues
                    Dim cFile$ = I.issueId + ".json"
                    cFile = safeFilename(cFile)



                    ' opportunity here to capture new scanIDs 
                    If IO.File.Exists(Path.Combine(cacheDir, cFile)) = True Then
                        ' Console.WriteLine("Found Issue cached -> " + cFile)
                        numCache += 1
                        GoTo skipThisOne
                    End If

                    Call setIssueReqVars(I.issueId)
tryAgain:
                    Dim newJson$ = setUpAPICall("getSingleIssue",, True, True)

                    If newJson = "ERROR" Then
                        Console.WriteLine("Trying again in 15 minutes")
                        Sleep(900000)
                        GoTo tryagain
                    End If

                    'Console.WriteLine(newJson)
                    Console.WriteLine(numCache.ToString + "/" + allIssues.Count.ToString + "  " + I.issueId + " -----> " + cFile)
                    Call saveJSONtoFile(newJson, Path.Combine(cacheDir, cFile))
                    numWrite += 1

skipthisone:
                    fileNames.Add(Path.Combine(cacheDir, cFile))
                Next

                Console.WriteLine(vbCrLf + "# of Objects   : " + allIssues.Count.ToString)
                Console.WriteLine("# New Objects  : " + numWrite.ToString)
                Console.WriteLine("Objects Cached : " + numCache.ToString)

                issuesCache = New List(Of singleIssue)
                Dim issuesJSON As List(Of String)
                issuesJSON = New List(Of String)

                GoTo noMoreThreading

                Call processCache() ', issuesCache)

                Console.WriteLine("Back into main process")
                GC.Collect()

                Console.WriteLine("# IssuesCache " + issuesCache.Count.ToString + " files..")
                Console.WriteLine("Cancelling lost threads")

                issueCacheLoader.CancelAsync()
                issueCacheLoader = New BackgroundWorker
                GC.Collect()

                Console.WriteLine("Grabbing remaining items")

                Dim newfilenameS As List(Of String) = New List(Of String)
                For Each I In issuesCache
                    Dim checkFile$ = Path.Combine(cacheDir, safeFilename(I.issueId) + ".json")
                    If ndxFilenames(checkFile) = -1 Then

                        Dim nD As JObject = JObject.Parse(streamReaderTxt(checkFile))
                        Dim newI As singleIssue = New singleIssue

                        newI = JsonConvert.DeserializeObject(Of singleIssue)(nD.SelectToken("data").SelectToken("getSingleIssueInfo").ToString)

                        issuesCache.Add(newI)
                        Console.WriteLine("Added " + newI.issueId)


                    End If
                Next

noMoreThreading:
                Dim pCnt As Integer = 0

                For Each F In fileNames
                    pCnt += 1
                    issuesJSON.Add(streamReaderTxt(F))
                    If pCnt Mod 1000 = 0 Then
                        Thread.Sleep(10)
                        Console.WriteLine(CStr(Now.ToString("hh\:mm\:ss\:ff")) + "> Progress: Read " + pCnt.ToString + " files")
                        Thread.Sleep(10)
                    End If
                Next

                Console.WriteLine(vbCrLf + "Deserializing " + issuesJSON.Count.ToString + " JSON strings")

                pCnt = 0
                For Each S In issuesJSON
                    pCnt += 1
                    Dim nD As JObject = JObject.Parse(S)
                    Dim newI As singleIssue = New singleIssue

                    newI = JsonConvert.DeserializeObject(Of singleIssue)(nD.SelectToken("data").SelectToken("getSingleIssueInfo").ToString)

                    If IsNothing(newI) = True Then
                        Console.WriteLine("Skipping NDX " + (pCnt - 1).ToString + "  -> file > " + fileNames(pCnt - 1))
                        GoTo skipThatOne
                    End If

                    issuesCache.Add(newI)
                    If pCnt Mod 5000 = 0 Then
                        Thread.Sleep(10)
                        Console.WriteLine(CStr(Now.ToString("hh\:mm\:ss\:ff")) + "> Progress: Deserialized " + pCnt.ToString + " JSONs")
                        Thread.Sleep(10)
                    End If
skipThatOne:
                Next
                'fileNames = Nothing
                issuesJSON = Nothing
                GC.Collect()
                Console.WriteLine("Freeing up memory - sending to report construction" + vbCrLf)
                Thread.Sleep(1000)
                Console.WriteLine("# IssuesCache " + issuesCache.Count.ToString + " single issue objects..")
                Console.WriteLine("Item 1: " + issuesCache(0).issueId + vbCrLf + "Last  : " + issuesCache(issuesCache.Count - 1).issueId)

                '                If numCache > issuesCache.Count Then
                '                    Call processCache()
                '                End If

                Call issueDetailRpt(issuesCache, toFilename)

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

        Dim addRepoTag$ = argValue("repotag", args)
        Dim repoOnlyTag As Boolean = False

        If Len(addRepoTag) Then repoOnlyTag = True

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

        If repoOnlyTag = True Then
            Console.WriteLine("Will apply tag '" + addRepoTag + "' to all applications defined as repo folders")
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

        If repoOnlyTag = True Then addedTag = addRepoTag

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

            If repoOnlyTag = True Then
                'Console.WriteLine("Checking " + app.appName)
                If Mid(app.appName, 1, 1) = "*" Then addTag = False
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
        '        Dim jsoN$ = ""
        '        jsoN = setUpAPICall("getAppsShort",, True)

        Call getAllIssues(, "getAppsShort", 500)

        'Dim allIssues As List(Of issueS) = New List(Of issueS)

        Dim currFile As Integer
        Dim cFile$
        Dim fileN$ = "getAppsShort.json"

        For currFile = 0 To numResponseFiles - 1
            cFile = Replace(fileN, ".json", currFile.ToString + ".json")
            Dim tempApps As List(Of oxAppshort) = New List(Of oxAppshort)
            Console.WriteLine("Deserializing JSON " + cFile)
            tempApps = OX.getAppInfoShort(streamReaderTxt(cFile))

            If numResponseFiles = 0 Then
                getAppListShort = tempApps
            Else
                ' is there a better way? test 
                For Each T In tempApps
                    getAppListShort.Add(T)
                Next
            End If
        Next


        Console.WriteLine("Total # of apps (500 apps per file): " + (numResponseFiles - 1).ToString)
        Console.WriteLine("# of Apps: " + getAppListShort.Count.ToString)



        '  getAppListShort = OX.getAppInfoShort(jsoN)
    End Function

    Public Function getAllIssues(Optional ByVal doShortIssues As Boolean = False, Optional ByVal differentCall$ = "", Optional ByVal diffOffset As Long = 0) As Integer
        numResponseFiles = 0

        Dim apiCall$ = "getIssues"
        Dim fileName$ = "getIssues.json"

        If doShortIssues = True Then
            apiCall = "getIssuesShort"
            fileName = "getIssuesShort.json"
        End If

        If Len(differentCall) Then
            apiCall = differentCall
            fileName = apiCall + ".json"
        End If

        Select Case apiCall
            Case "getIssues", "getIssuesShort"
                Call setGetIssuesVars(0, doShortIssues)
            Case "getAppsShort"
                Call setGetAppsShortVars(0) ', doShortIssues)
        End Select

        Console.WriteLine("Pulling first page > " + fileName + " > " + Replace(fileName, ".json", "0.json"))

        Call setUpAPICall(apiCall, Replace(fileName, ".json", "0.json"))

        Select Case apiCall
            Case "getIssues", "getIssuesShort"
                Dim respIssue As listIssues = New listIssues
                respIssue = OX.getListIssues(Replace(fileName, ".json", "0.json"))
                With respIssue
                    Console.WriteLine("Total Issues: " + .totalIssues.ToString)
                    'Console.WriteLine("Filtered Issues: " + .totalFilteredIssues.ToString)
                    'Console.WriteLine("Offset: " + .offset.ToString)
                End With

                numResponseFiles = 0
                'this was the first file

                Do Until numResponseFiles = numIssueRequests(respIssue.totalFilteredIssues)
                    numResponseFiles += 1
                    Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + numIssueRequests(respIssue.totalFilteredIssues).ToString + " requests")
                    Call setGetIssuesVars(issueLimit * numResponseFiles, doShortIssues)
                    Call setUpAPICall(apiCall, Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
                    goUpALine()

                Loop

                Call setGetIssuesVars(0, doShortIssues)

            Case "getAppsShort"
                Dim appPageLimit As Long = 50
                Dim respApps As listApps = New listApps
                respApps = OX.getListAppsPaging(Replace(fileName, ".json", "0.json"))
                With respApps
                    Console.WriteLine("Total: " + .total.ToString)
                    ' Console.WriteLine("Total Filtered Apps: " + .totalFilteredApps.ToString)
                    Console.WriteLine("Total Irrelevant: " + .totalIrrelevantApps.ToString)
                    'Console.WriteLine("Offset: " + .offset.ToString)
                    appPageLimit = .offset
                End With

                numResponseFiles = 0
                'this was the first file

                Dim issReq As Long = numIssueRequests(respApps.total, appPageLimit) 'hardcoding!! UGH

                Do Until numResponseFiles = issReq
                    numResponseFiles += 1
                    Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + issReq.ToString + " requests")
                    ' hardcoding

                    Call setGetAppsShortVars(appPageLimit * numResponseFiles)
                    Call setUpAPICall(apiCall, Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
                    goUpALine()

                Loop

                Call setGetAppsShortVars(0) ', doShortIssues)
        End Select
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

    Public Sub setIssueReqVars(issueId$)
        Dim reqVars As newIssueDetailRequestVARS = New newIssueDetailRequestVARS(issueId)
        OX = New oxWrapper("", "")
        Dim newJson$ = OX.jsonGetNewIssueDetailVars(reqVars)
        'Console.WriteLine(newJson)
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, "getSingleIssue.variables.json"))
    End Sub

    Public Sub setGetAppsShortVars(offSet As Integer)
        Dim toFileN$ = "getAppsShort.variables.json"
        Dim appReqVar As appsRequestVARS = New appsRequestVARS(offSet)
        Dim newJson$ = ""
        ' this is sloppy pls figure out why you did this, this way
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetAppsVars(appReqVar)
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, toFileN))

    End Sub


    Public Sub setGetIssuesVars(offSet As Integer, Optional ByVal doIssuesShort As Boolean = False)
        Dim toFileN$ = "getIssues.variables.json"
        If doIssuesShort = True Then toFileN = "getIssuesShort.variables.json"
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
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, toFileN))

    End Sub

    Public Sub setEditTagsVarsRequests(evReq As editTagsRequestVARS)
        Dim newJson$ = ""
        ' more sloppy - make OX global for main
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetEditTagsVars(evReq)
        'Console.WriteLine(newJson)
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, "modifyAppsTags.variables.json"))
    End Sub


    Public Function numIssueRequests(totalIssues As Long, Optional ByVal iLimit As Long = 0) As Long
        numIssueRequests = 0
        Dim issLimit As Long = issueLimit
        If iLimit > 0 Then issLimit = iLimit

        Dim tlCalled As Integer = 0

        Do Until tlCalled >= totalIssues
            numIssueRequests += 1
            tlCalled += issLimit
        Loop

    End Function

    Public Function setUpAPICall(apiCall$, Optional ByVal fileN$ = "", Optional ByVal showJSON As Boolean = False, Optional noConsole As Boolean = False) As String
        setUpAPICall = ""
        OX = New oxWrapper("", "")
        'Console.WriteLine("Retrieving JSON From OX API: " + apiCall)

        Dim getFile$ = apiCall + "_response.json"
        If fileN <> "" Then fileN = fileN

        safeKILL(getFile)
        safeKILL(Path.Combine(pyDir, getFile))
        If Len(fileN) Then safeKILL(fileN)

        If noConsole = False Then Console.WriteLine("Executing Python request for '" + apiCall + "'")

        Dim succesS As Boolean = OX.getJSON(apiCall)

        If IO.File.Exists(Path.Combine(pyDir, getFile)) = False Then succesS = False

        If succesS = False Then
            Console.WriteLine("Check the underlying Python connector, make sure Python is in the path. Try 'python python_examp.py " + apiCall + "'")
            setUpAPICall = "ERROR"
            Exit Function
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


        Console.WriteLine("Total # of files (" + issueLimit.ToString + " issues per file): " + (numFiles - 1).ToString)
        Console.WriteLine("# of Issues: " + allIssues.Count.ToString)
        Return allIssues
    End Function

    Public Function buildShortIssues(fileN$, numFiles As Integer) As List(Of issueS)
        Dim allIssues As List(Of issueS) = New List(Of issueS)

        Dim currFile As Integer
        Dim cFile$

        For currFile = 0 To numFiles
            cFile = Replace(fileN, ".json", currFile.ToString + ".json")
            safeKILL(Path.Combine(cacheDir, cFile))
            FileCopy(cFile, Path.Combine(cacheDir, cFile))

            Dim tempIssues As List(Of issueShort) = New List(Of issueShort)
            Console.WriteLine("Deserializing JSON " + cFile)
            tempIssues = OX.returnShortIssues(streamReaderTxt(cFile))

            For Each tI In tempIssues
                Dim nI As issueS = New issueS
                With nI
                    .id = tI.id
                    .issueId = tI.issueId
                    .created = tI.created
                    .createdAt = tI.createdAt
                    .scanId = tI.scanId
                End With
                allIssues.Add(nI)
            Next
        Next


        Console.WriteLine("Total # of files (" + issueLimit.ToString + " issues per file): " + (numFiles - 1).ToString)
        Console.WriteLine("# of Issues: " + allIssues.Count.ToString)
        Return allIssues
    End Function

    Public Function loadCachedList() As List(Of issueS)
        Dim allIssues As List(Of issueS) = New List(Of issueS)

        Dim currFile As Integer = 0
        Dim cFile$ = Path.Combine(cacheDir, "getIssuesShort" + currFile.ToString + ".json")

        OX = New oxWrapper("", "")

        Do Until IO.File.Exists(cFile) = False
            Dim tempIssues As List(Of issueShort) = New List(Of issueShort)
            Console.WriteLine("Deserializing JSON " + cFile)
            tempIssues = OX.returnShortIssues(streamReaderTxt(cFile))

            For Each tI In tempIssues
                Dim nI As issueS = New issueS
                With nI
                    .id = tI.id
                    .issueId = tI.issueId
                    .created = tI.created
                    .createdAt = tI.createdAt
                    .scanId = tI.scanId
                End With
                allIssues.Add(nI)
            Next
            currFile += 1
            cFile$ = Path.Combine(cacheDir, "getIssuesShort" + currFile.ToString + ".json")
        Loop

        Console.WriteLine("Total # of files (" + issueLimit.ToString + " issues per file): " + (currFile + 1).ToString)
        Console.WriteLine("# of Issues: " + allIssues.Count.ToString)
        Return allIssues
    End Function


    Private Sub appTagRPT(allApps As List(Of oxAppshort), fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        rAR.someColl = New Collection
        rAR.s1 = fileN

        Call safeKILL(fileN)

        Dim totalNum As Integer = 0
        Console.WriteLine("Calculating total number of tags")

        For Each aA In allApps
            For Each aTag In aA.tags

                If aTag.isOxTag = True Then
                    totalNum += 1
                End If
            Next
        Next

        'Console.WriteLine("Building output from results")

        With rAR.someColl
            .Add("APP_TAG")
            .Add("APP_NAME")
            .Add("#_TAGS")
        End With

        Dim xls3d(totalNum - 1, 2) As Object

        Dim roW As Long = 0
        For Each aA In allApps
            For Each aTag In aA.tags

                If aTag.isOxTag = True Then
                    xls3d(roW, 0) = aTag.name
                    xls3d(roW, 1) = aA.appName
                    xls3d(roW, 2) = 1
                    roW += 1
                End If
            Next
        Next

        Console.WriteLine("Throwing object of " + (roW * 3).ToString + " elements at Excel..")
        rAR.booL1 = True ' create XLS not CSV

        If roW = 0 Then
            Console.WriteLine("Nothing to output")
            End
        End If

        Call newRpt.dump2XLS(xls3d, roW - 1, rAR,, True)

        End
    End Sub
    Private Sub issueDetailRpt(ByRef allI As List(Of singleIssue), fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        Console.WriteLine("Building report for " + allI.Count.ToString + " single issues")

        rAR.someColl = New Collection
        rAR.s1 = fileN + "_issues.xlsx"

        Call safeKILL(fileN)
        ' first the issues for dedup testing

        With rAR.someColl
            .Add("ISSUE_ID")
            .Add("TITLE")
            .Add("CATEGORY")
            .Add("MATCH_REPO")
            .Add("RESOURCE")
            .Add("SEVERITY")
            .Add("ORIG_SEVERITY")
            .Add("#_OCCURRENCES")
            .Add("CREATED")
            .Add("SCANID")
            .Add("SOURCETOOLS_1")
            .Add("SOURCETOOLS_2")
            .Add("SOURCETOOLS_3")
            .Add("FILEN")
            .Add("#_SF")
            .Add("#_SF_R")
            .Add("#_SF_E")
            .Add("#_SF_D")
            .Add("#_SEV_PLUS")
            .Add("#_SEV_MINUS")
            .Add("#_ISSUES")
        End With

        Dim sfColl As Collection = New Collection
        With sfColl
            .Add("SF_NAME")
            .Add("REACHABILITY")
            .Add("EXPLOITABILITY")
            .Add("DAMAGE")
            .Add("RESOURCE")
            .Add("WEIGHT")
            .Add("ISSUE_ID")
            .Add("ISSUE_TITLE")
            .Add("REASON")
            .Add("FILEN")
            .Add("#_SF")
            .Add("#_SF_R")
            .Add("#_SF_E")
            .Add("#_SF_D")
            .Add("RED")
            .Add("MATCH")
        End With


        Dim xls3d(allI.Count - 1, 20) As Object

        Dim allSF As Long = 0
        Console.WriteLine(allI(0).issueId)
        Console.WriteLine(allI(allI.Count - 1).issueId)

        Dim K As Long = 0

        Dim roW As Long = 0
        For K = 0 To allI.Count - 1 ' In allI
            Dim aA As singleIssue = allI(K)
            'xls3d(0, 0) = "some string"
            'Console.WriteLine(aA.issueId)
            xls3d(roW, 0) = aA.issueId
            xls3d(roW, 1) = aA.mainTitle
            xls3d(roW, 2) = aA.category.name
            Dim match = False
            If Mid(aA.app.name, 1, 1) <> "*" Then match = True
            xls3d(roW, 3) = CStr(match)
            Dim rID$ = ""
            If aA.sourceTools(0) = "WIZ" And Val(Mid(aA.resource.id, 1, 1)) > 0 Then rID = "WIZ" Else rID = aA.resource.id
            xls3d(roW, 4) = rID
            xls3d(roW, 5) = aA.severity
            xls3d(roW, 6) = aA.originalToolSeverity
            xls3d(roW, 7) = aA.occurrences
            xls3d(roW, 8) = jStoDate(aA.created)
            xls3d(roW, 9) = aA.scanId
            xls3d(roW, 10) = aA.sourceTools(0)
            Dim sT2$ = ""
            Dim sT3$ = ""
            If aA.sourceTools.Count > 1 Then sT2 = aA.sourceTools(1)
            If aA.sourceTools.Count > 2 Then sT3 = aA.sourceTools(2)
            xls3d(roW, 11) = sT2
            xls3d(roW, 12) = sT3
            xls3d(roW, 13) = fileNames(K)
            xls3d(roW, 14) = aA.numSevFactors()
            allSF += aA.numSevFactors
            xls3d(roW, 15) = aA.numSevFactors(True)
            xls3d(roW, 16) = aA.numSevFactors(, True)
            xls3d(roW, 17) = aA.numSevFactors(,, True)
            If aA.increasedSev = True Then xls3d(roW, 18) = 1 Else xls3d(roW, 18) = 0
            If aA.decreasedSev = True Then xls3d(roW, 19) = 1 Else xls3d(roW, 19) = 0
            xls3d(roW, 20) = 1
            roW += 1
        Next

        Dim issueRows As Long = roW
        roW = 0

        Dim xls3d2(allSF - 1, 15) As Object
        For K = 0 To allI.Count - 1 ' In allI
            Dim aA As singleIssue = allI(K)
            'issue stuff
            Dim rID$ = ""
            If aA.sourceTools(0) = "WIZ" And Val(Mid(aA.resource.id, 1, 1)) > 0 Then rID = "WIZ" Else rID = aA.resource.id
            Dim issueID$ = aA.issueId
            Dim mainTitle$ = aA.mainTitle
            Dim fN$ = fileNames(K)

            For Each SF In aA.severityChangedReason
                xls3d2(roW, 0) = SF.shortName
                xls3d2(roW, 4) = rID
                xls3d2(roW, 5) = SF.changeNumber
                xls3d2(roW, 6) = SF.reason
                xls3d2(roW, 7) = issueID
                xls3d2(roW, 8) = mainTitle
                xls3d2(roW, 9) = fN
                xls3d2(roW, 10) = 1
                xls3d2(roW, 14) = LCase(SF.changeCategory)
                Dim match = False
                If Mid(aA.app.name, 1, 1) <> "*" Then match = True
                xls3d2(roW, 15) = CStr(match)
                Select Case LCase(SF.changeCategory)
                    Case "reachable"
                        xls3d2(roW, 11) = 1
                        xls3d2(roW, 12) = 0
                        xls3d2(roW, 13) = 0
                        xls3d2(roW, 1) = "TRUE"
                        xls3d2(roW, 2) = "FALSE"
                        xls3d2(roW, 3) = "FALSE"

                    Case "exploitable"
                        xls3d2(roW, 11) = 0
                        xls3d2(roW, 12) = 1
                        xls3d2(roW, 13) = 0
                        xls3d2(roW, 1) = "FALSE"
                        xls3d2(roW, 2) = "TRUE"
                        xls3d2(roW, 3) = "FALSE"

                    Case "damage"
                        xls3d2(roW, 11) = 0
                        xls3d2(roW, 12) = 0
                        xls3d2(roW, 13) = 1
                        xls3d2(roW, 1) = "FALSE"
                        xls3d2(roW, 2) = "FALSE"
                        xls3d2(roW, 3) = "TRUE"

                End Select
                roW += 1
            Next
        Next



        Console.WriteLine("Throwing object of " + (allSF * 14).ToString + " elements at Excel..")
        rAR.booL1 = True ' create XLS not CSV

        rAR.s3 = "IR"
        Call newRpt.dump2XLS(xls3d, allI.Count, rAR,, True)

        rAR.s3 = "SF"
        rAR.someColl = sfColl
        rAR.s1 = fileN + "_sevFactors.xlsx"


        Call newRpt.dump2XLS(xls3d2, allSF, rAR,, True)


    End Sub

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

    Public Sub processCache() ', objList As List(Of singleIssue))
        '        Console.WriteLine("Kicking off background threads for # files.. " + fileNames.Count.ToString)

        '        Dim tasK As taskArgs = New taskArgs
        '       Task.listString = fileNames
        '      Console.WriteLine("Kicking off background threads for # files.. " + tasK.listString.Count.ToString)
        Dim numHolding As Integer = 0
        Dim oldCount As Long = 0
        Dim mustRetry As Boolean = False

startOver:

        currentlyLoading = True

        issueCacheLoader = New BackgroundWorker
        issueCacheLoader.WorkerSupportsCancellation = True
        issueCacheLoader.WorkerReportsProgress = True
        issueCacheLoader.RunWorkerAsync()

        '        Console.WriteLine("Watching load from cache" + vbCrLf + vbCrLf + vbCrLf)

        '        Dim fnCount As Long = fileNames.Count
        '       Dim numIterations As Long = 0

        Do Until currentlyLoading = False
            '      Dim iCount As Long = issuesCache.Count
            '     numIterations += 1
            '    Console.WriteLine("Processed: " + iCount.ToString + "/" + fnCount.ToString + "     FPS: " + Math.Round((iCount - oldCount) / 10, 2).ToString + "    ")
            Threading.Thread.Sleep(5000)
            '  If oldCount > 0 And iCount = oldCount Then numHolding += 1 Else numHolding = 0
            '
            'If numHolding > 5 Then
            'Dim newfilenameS As List(Of String) = New List(Of String)
            ''         For Each I In issuesCache
            '             Dim checkFile$ = safeFilename(Path.Combine(cacheDir, I.issueId))
            '             If ndxFilenames(checkFile) = -1 Then
            '                 newfilenameS.Add(checkFile)
            '             End If
            '         Next
            '         fileNames = newfilenameS
            '         Console.WriteLine("Built new list of " + fileNames.Count.ToString + " files..")
            '         mustRetry = True
            '    Else
            '        If numHolding Then Console.WriteLine(vbCrLf + vbCrLf + "Detecting rogue threads " + numHolding.ToString + "/5")
            '    End If
            '    oldCount = iCount
        Loop

escapeLoop:
        'If mustRetry = True Then GoTo startOver

        'Console.WriteLine("Completed processCache - exiting back")
    End Sub
    Public Sub issueThreadDeserialize(fileN$)
        'Dim W As New WaitCallback(AddressOf issueCacheLoader)
        'Console.WriteLine("Received " + fileN)
        On Error GoTo errorcatch

        Threading.Thread.Sleep(20)

        If issueCacheLoader.CancellationPending = True Then GoTo endHere

        Dim nD As JObject = JObject.Parse(streamReaderTxt(fileN))
        Dim I As singleIssue = New singleIssue

        I = JsonConvert.DeserializeObject(Of singleIssue)(nD.SelectToken("data").SelectToken("getSingleIssueInfo").ToString)
        If issueCacheLoader.CancellationPending = True Then GoTo endHere

        issuesCache.Add(I)
        Dim iC As Long = issuesCache.Count
        Dim fN As Long = fileNames.Count

        issueCacheLoader.ReportProgress(Math.Round(100 - ((fN - iC) / fN) * 100), 2)
        Threading.Thread.Sleep(100)
        'Console.WriteLine("Received " + I.id)

endHere:
errorcatch:
        GC.Collect()
        'Threading.Thread.Sleep(50)
    End Sub



    Private Sub issueCacheLoader_DoWork(sender As Object, e As DoWorkEventArgs) Handles issueCacheLoader.DoWork


        Console.WriteLine("1000x thread concurrency for # files.. " + fileNames.Count.ToString)

        'Console.WriteLine("Do this!")
        Dim W As WaitCallback = New WaitCallback(AddressOf issueThreadDeserialize)
        ThreadPool.SetMaxThreads(200, 50)

        Console.WriteLine("# files: " + fileNames.Count.ToString)

        Dim K As Integer = 0


        For Each F In fileNames
            K += 1
            'For K = 1 To 150
            ThreadPool.QueueUserWorkItem(W, F)
            'Thread.Sleep(50)
            '            Console.WriteLine("Adding to threadpool " + K.ToString)
            'If K Mod 1000 Then Console.WriteLine("Queueing: " + K.ToString)

            Threading.Thread.Sleep(100)

        Next
        '        Console.WriteLine("Done Queueing")



    End Sub
    Public Class taskArgs
        Public listString As List(Of String)
    End Class
    '
    Private Sub issueCacheLoader_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles issueCacheLoader.RunWorkerCompleted
        Console.WriteLine("All Done! > " + issuesCache.Count.ToString + "/" + fileNames.Count.ToString + vbCrLf + issuesCache(0).id + vbCrLf + issuesCache(issuesCache.Count - 1).id)
        currentlyLoading = False
    End Sub

    Private Function ndxFilenames(ByVal fileN$) As Long
        ndxFilenames = 0
        For Each F In fileNames
            If F = fileN Then
                Return ndxFilenames
            Else
                ndxFilenames += 1
            End If
        Next
        ndxFilenames = -1
    End Function

    Private Sub issueCacheLoader_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles issueCacheLoader.ProgressChanged
        Static oldP As Long = 0

        Dim pC As Long = e.ProgressPercentage
        If oldP = pC Then Exit Sub

        If pC Mod 0.5 = 0 Then Console.WriteLine("Current %: " + e.ProgressPercentage.ToString)

        oldP = pC
    End Sub
End Module
