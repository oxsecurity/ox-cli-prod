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
Imports System.Runtime.Intrinsics.Arm


Module Program
    Public aTimer As New System.Timers.Timer
    Public OX As oxWrapper

    Public currOffset = 0
    Public issueLimit = 10000

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

    Function Main(args As String()) As Integer
        ' Environment.ExitCode = 1
        If UBound(args) = -1 Then
            Console.WriteLine("You must enter a command. Try 'help'.")
            End
        End If

        OX = New oxWrapper("", "")
        '        Main = 1
        '       Console.WriteLine("Returning Exit Code " + Environment.ExitCode.ToString)
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

                Console.WriteLine(fLine("getconnectors", "prints all connectors to screen or CSV"))
                Console.WriteLine(fLine("", "[OPTIONAL] --FILE (name of CSV file to create)"))
                Console.WriteLine("-----------------")


                Console.WriteLine(fLine("apptagsxls", "retrieves all APP TAGS and creates pivot Excel doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name of XLS file to create)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("getapps", "prints all irrelevant apps to screen or CSV"))
                Console.WriteLine(fLine("", "[OPTIONAL] --FILE (name of CSV file to create)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("getirrelevantapps", "prints all irrelevant apps to screen or CSV"))
                Console.WriteLine(fLine("", "[OPTIONAL] --FILE (name of CSV file to create)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("issuesdetailed", "retrieves all issue data and creates pivot Excel doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name of XLS file to create)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --CACHE false  (by default, issues will be stored locally for caching)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("issuesxls", "retrieves all issues and creates pivot Excel doc"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (name of XLS file to create)"))
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("issuescsv", "retrieves all issues and creates CSV doc"))
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
                Console.WriteLine("-----------------")
                Console.WriteLine(fLine("devdetail", "Takes in JSON of Dev Detail and presents 30/90/180 commit stats, plus creates pivot"))
                Console.WriteLine(fLine("", "[REQUIRED] --INFILE (filename of JSON)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --FILE (Excel report filename)"))
                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("tagfromxls", "Takes in EXCEL of columns with APP NAMES or APP IDs and their respective tag, adds TAGS to apps"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (XLS filename containing App and Tag info)"))
                Console.WriteLine(fLine("", "[REQUIRED] --APPNAME (Excel column representing App Name)"))
                Console.WriteLine(fLine("", "[REQUIRED] --TAG (Excel column representing TAG to apply)"))
                Console.WriteLine(fLine("", "[OPTIONAL] --COMMIT (set to true to apply tags, otherwise test only)"))
                Console.WriteLine("-----------------")


                Console.WriteLine(fLine("gatecheck", "Checks most recent results for HIGH+ vulns across all sources based on recent scan"))
                Console.WriteLine(fLine("", "[REQUIRED] --APPNAME (Name of application to check)"))
                Console.WriteLine(fLine("", "[REQUIRED] --APPNAME (Failcode to return eg 1 when HIGH+ encountered)"))

                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("jsonfromxls", "Creates JSON used to define mapping of integration projects -> OX applications"))
                Console.WriteLine(fLine("", "[REQUIRED] --XLS (XLS filename containing App and Tag info)"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (Excel column representing App Name)"))
                Console.WriteLine(fLine("", "[REQUIRED] --APPNAME (Excel column representing App Name)"))
                Console.WriteLine(fLine("", "[REQUIRED] --MAP (Excel column representing App Name)"))
                Console.WriteLine(fLine("", "[REQUIRED] -- MAPTYPE (Type of integration eg blackduck)"))

                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("readout", "Creates multiple documents for readout - requires MS Excel Powerpoint"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (Filenames of reports will begin with this prefix)"))

                Console.WriteLine("-----------------")

                Console.WriteLine(fLine("cvecsv", "Creates CSV of all CVEs"))
                Console.WriteLine(fLine("", "[REQUIRED] --FILE (Filename of CSV output file)"))


                Console.WriteLine("=======================================================================================================================================================")
                End


            Case "cvecsv"
                Call cveList(args)
                End

            Case "gatecheck"
                'Dim appId$ = argValue("appid", args)
                Dim appName$ = argValue("appname", args)
                Dim failCode$ = argValue("failcode", args)

                Console.WriteLine(vbCrLf + "OX SECURITY GATE CHECK - WILL LOOK FOR RESULTS FROM RECENT SCAN FOR --appname SPECIFIED" + vbCrLf)

                If appName = "" Then
                    Console.WriteLine("You must specify --appname")
                    Environment.ExitCode = 1
                    End
                End If
                Dim allIssues As List(Of issuesMedium) = New List(Of issuesMedium)

                Console.WriteLine("Getting all issues High Severity & above for " + appName)
                If failCode = "" Then
                    Console.WriteLine("No --failcode entered - this action will log vulnerabilities only")
                Else
                    Console.WriteLine("--failcode provided - this action will return an exit code of " + failCode + " if issues exist in recent scheduled scan")
                End If

                'Console.WriteLine(vbCrLf)

                Dim numFiles As Integer = getAllIssues(, "getIssuesMedium",, appName)
                'Console.WriteLine(vbCrLf + vbCrLf + "Pulling " + numFiles.ToString + " files")
                allIssues = buildMediumIssues("getIssuesMedium.json", numFiles)
                'Console.WriteLine("# of issues: " + allIssues.Count.ToString)

                Call consoleDump(allIssues)

                If Len(failCode) And allIssues.Count > 0 Then
                    If Val(failCode) > 0 Then
                        Environment.ExitCode = Val(failCode)
                        Console.WriteLine(vbCrLf + "ERROR: EXIT CODE=" + failCode + " for " + allIssues.Count.ToString + " vulnerabilities")
                    End If
                End If
                End



            Case "readout"

                Call readoutSub(args)
                End


            Case "tagfromxls"
                Call tagFromXLS(args)
                End

            Case "getirrelevantapps"
                Dim allApps As List(Of oxAppIrrelevant) = New List(Of oxAppIrrelevant)
                allApps = getAppListIrrelevant()
                Console.WriteLine("# Of Apps: " + allApps.Count.ToString)
                Dim fileN$ = argValue("file", args)

                Dim csV$ = ""

                Dim allReasons As Collection = New Collection
                For Each I In allApps
                    For Each R In I.irrelevantReasons
                        If grpNDX(allReasons, R) = 0 Then allReasons.Add(R)
                    Next
                Next

                For Each I In allApps
                    Dim lDate$ = ""
                    Dim aName$ = ""
                    Dim iReason$ = ""

                    aName = I.appName + " [" + I.appId + "]"
                    lDate = CStr(jStoDate(CLng(I.lastCodeChange)))

                    For Each reasoN In allReasons
                        Dim foundReason As Boolean = False
                        For Each R In I.irrelevantReasons
                            If reasoN = R Then foundReason = True
                        Next
                        If foundReason = True Then
                            iReason += "1,"
                        Else
                            iReason += ","
                        End If
                    Next

                    iReason = Mid(iReason, 1, Len(iReason) - 1)
                    If fileN = "" Then
                        Console.WriteLine(aName + spaces(60 - Len(aName)) + lDate + spaces(25 - Len(lDate)) + iReason)
                    Else
                        csV += qT(I.appName) + "," + I.appId + "," + I.link + "," + lDate + "," + iReason + vbCrLf
                    End If
                Next
                If Len(csV) Then
                    safeKILL(fileN)
                    Dim hdR$ = "APP_NAME,APP_ID,LINK,LAST_CHANGE,"
                    'IRRELEVANT_REASON" + vbCrLf
                    For Each rsN In allReasons
                        hdR += rsN + ","
                    Next
                    hdR = Mid(hdR, 1, Len(hdR) - 1) + vbCrLf
                    safeKILL(fileN)
                    streamWriterTxt(fileN, hdR + csV)
                    Console.WriteLine("File written to " + fileN)
                End If
                End

            Case "jsonfromxls"
                Call createMappingJSON(args)
                End


            Case "strcompare"
                Dim matcH As Single = 0
                matcH = GetSimilarity("four score and seven years ago", "for scor and sevn yeres ago")
                Console.WriteLine("Match = " + matcH.ToString)
                End


            Case "gitlab_tag_groups"
                Dim reportOnly As Boolean = False
                If LCase(argValue("reportonly", args)) = "true" Then reportOnly = True
                Dim allApps As List(Of oxAppshort) = New List(Of oxAppshort)

                If reportOnly = False Then
                    allApps = getAppListShort()
                    Console.WriteLine("# of Applications: " + allApps.Count.ToString)
                End If

                Dim fileJSON$ = argValue("glabjson", args)
                Dim fileN$ = argValue("file", args)

                Dim csV$ = ""
                Dim glabRepos As List(Of glabRepo) = New List(Of glabRepo)
                glabRepos = OX.returnGitLabRepos(fileJSON)

                Console.WriteLine("# of OX Apps: " + allApps.Count.ToString)
                Console.WriteLine("# of GitLab JSON Repos: " + glabRepos.Count.ToString)
                'For Each gL In glabRepos
                'Console.WriteLine(gL.name_with_namespace + "," + gL.name + "," + gL.namespace.name)
                'Next
                csV = "OX_ID,OX_APP_NAME,TAG,NS_KIND,WEB_URL,COUNT,EXCEPTION" + vbCrLf
                For Each oApp In allApps
                    If Mid(oApp.appName, 1, 1) = "*" Then GoTo skipFakeApp
                    ' If InStr(oApp.appName, "terraform-okta-group") > 0 Then
                    '     Dim K As Integer
                    '     K = 12
                    ' End If

                    Dim cLine$ = qT(oApp.appId) + "," + qT(oApp.appName) + ","
                    Dim numEntries As Integer = 0

                    Dim groupsString$ = ""
                    For Each gL In glabRepos
                        ' old way, resulted in numerous instances of no match or multiple matches
                        '   If oApp.appName = gL.name Then
                        '   groupsString += gL.namespace.name + ","
                        '   numEntries += 1
                        '   End If

                        ' new way
                        ' match OX link to GL web url
                        If oApp.link = gL.web_url Then
                            groupsString += qT(gL.namespace.name) + "," + gL.namespace.kind + ","
                            numEntries += 1
                        End If
                    Next
                    'If Len(groupsString) Then
                    '    If Mid(groupsString, Len(groupsString), 1) = "," Then groupsString = Mid(groupsString, 1, Len(groupsString) - 1)
                    'End If
                    If groupsString = "" Then groupsString = "N/A,N/A,"
                    cLine += groupsString + qT(oApp.link) + ",1"
                    If numEntries = 0 Then cLine += ",NO EXACT MATCH"
                    If numEntries > 1 Then cLine += ",MORE THAN ONE MATCH"
                    csV += cLine + vbCrLf
                    Console.WriteLine(cLine)
skipFakeApp:
                Next
                Call safeKILL(fileN)
                Call streamWriterTxt(fileN, csV)

                End


            Case "devdetail"
                Call devDetailXLS(args)
                End


            Case "getapps"
                Dim allApps As List(Of oxAppshort) = getAppListShort()
                Console.WriteLine("# of Applications: " + allApps.Count.ToString)

                Dim fileN$ = argValue("file", args)

                Dim csV$ = ""

                Dim maxNumTags As Integer = 0

                For Each I In allApps
                    Dim aName$ = ""
                    Dim tagCsv$ = ""

                    aName = I.appName + " [" + I.appId + "]"
                    For Each R In I.tags
                        tagCsv$ += R.displayName + ","
                    Next
                    If I.tags.Count > maxNumTags Then maxNumTags = I.tags.Count

                    tagCsv$ = Mid(tagCsv$, 1, Len(tagCsv$) - 1)

                    If fileN = "" Then
                        Console.WriteLine(aName + spaces(60 - Len(aName)) + I.tags.Count.ToString + spaces(10) + tagCsv$)
                    Else
                        csV += I.appName + "," + I.appId + "," + tagCsv + vbCrLf
                    End If
                Next

                If Len(csV) Then
                    Dim hdR$ = "APP_NAME,APP_ID,"
                    Dim K As Integer

                    For K = 1 To maxNumTags
                        hdR += "TAG_" + K.ToString + ","
                    Next
                    hdR = Mid(hdR, 1, Len(hdR) - 1) + vbCrLf
                    Console.WriteLine("HEADERS:" + vbCrLf + hdR)

                    safeKILL(fileN)
                    streamWriterTxt(fileN, hdR + csV)
                    Console.WriteLine("File written to " + fileN)

                End If

                End



            Case "getconnectors"
                Dim allConnections As List(Of connectorFamily) = New List(Of connectorFamily)

                allConnections = getOxConnectors()

                Dim isConfig As Boolean = False
                If LCase(argValue("configonly", args)) = "true" Then isConfig = True
                Dim fileN$ = argValue("file", args)
                Dim csV$ = ""
                Dim cLine$ = ""

                cLine = qT("FAMILY") + "," + qT("NAME") + vbCrLf ' + "," + qT("DESCRIPTION") + "," + qT("CREDENTIAL_TYPES") + vbCrLf
                Console.WriteLine(cLine)

                For Each F In allConnections
                    Dim aLine$ = ""
                    For Each CONN In F.connectors
                        If isConfig = True And CONN.connector.isConfigured = False Then GoTo skip

                        aLine = qT(F.familyDisplayName) + ","
                        Dim C As oxConnection = CONN.connector
                        aLine += qT(C.displayName) ' + "," + qT(C.description)
                        'Dim creD$ = ""

                        'If C.credentialsTypes IsNot Nothing Then
                        '    If C.credentialsTypes.Count Then
                        '        For Each CT In C.credentialsTypes
                        '            creD += CT + ","
                        '        Next
                        '    End If
                        'End If

                        'aLine += "," + qT(creD) ' + vbCrLf

                        Console.WriteLine(aLine)
                        cLine += aLine + vbCrLf
skip:
                    Next
                    'cLine += aLine + vbCrLf
                    'Console.WriteLine(aLine)
                Next


                If Len(fileN) Then streamWriterTxt(fileN, cLine)
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
                End

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
                        GoTo skipthisone
                    End If

                    Call setIssueReqVars(I.issueId)
tryAgain:
                    Dim newJson$ = setUpAPICall("getSingleIssue",, True, True)

                    If newJson = "ERROR" Then
                        Console.WriteLine("Trying again in 15 minutes")
                        Sleep(900000)
                        GoTo tryAgain
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

                ' Call issueDetailRpt(issuesCache, toFilename)

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



    End Function


    Public Function addTag(tagName$, Optional ByVal dName$ = "", Optional ByVal tType$ = "simple") As String
        addTag = "" ' returns empty if unsuccessful otherwise tagid of new tag
        If dName = "" Then dName = tagName
        If tType = "" Then tType = "simple"

        Console.WriteLine("Adding tag:")
        Call setAddTagVars(tagName, dName, tType)
        Thread.Sleep(1000)
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

    Public Sub cveList(args() As String)
        Dim rFiles As Integer = 0
        Dim allIssues = New List(Of issueS)

        Dim done As Collection = New Collection
        done = CSVFiletoCOLL(Path.Combine(ogDir, "issues_sofar.txt"))

        Console.WriteLine(done.Count.ToString + " issues already cached")

        'Call safeKILL(Path.Combine(ogDir, "issues_sofar.txt"))

        Dim fileN$ = argValue("file", args)
        If fileN = "" Then
            Console.WriteLine("You must provide the filename of the CSV to write --FILE (csv filename)")
            Exit Sub
            '        Else
            '           Call safeKILL(fileN)
        End If

        Console.WriteLine("Getting issue list")

        If argValue("numfiles", args) <> "" Then
            rFiles = Val(argValue("numfiles", args))
        End If

        If rFiles = 0 Then
            Call getAllIssues(True)
        Else
            numResponseFiles = rFiles
        End If

        'Console.WriteLine("# resp files: " + (numResponseFiles).ToString)
        allIssues = buildIssues("getIssuesShort.json", numResponseFiles)

        Dim csV$ = ""
        Dim totalNum As Integer = 0

        For Each i In allIssues
            If i.category.name <> "Open Source Security" Then GoTo countMe
            totalNum += 1
countMe:
        Next

        Dim currCtr As Integer = 1
        For Each i In allIssues
            csV = ""
            ' call single issue for all cves/vulns
            If i.category.name <> "Open Source Security" Then GoTo skipIt

            If grpNDX(done, i.issueId) > 0 Then
                Console.WriteLine("Already have " + i.issueId)
                currCtr += 1
                GoTo skipIt
            End If

            Call setIssueReqVars(i.issueId, "getSCAvulns.variables.json")
            Dim jsoN$ = ""


            Console.WriteLine("[" + currCtr.ToString + "/" + totalNum.ToString + "] " + i.app.name + " | " + i.issueId)
            jsoN = setUpAPICall("getSCAvulns",, True)
            currCtr += 1
            'Console.WriteLine(jsoN)

            Dim issueVulns As scaCVE = New scaCVE
            Dim nD As JObject = JObject.Parse(jsoN)
            issueVulns = JsonConvert.DeserializeObject(Of scaCVE)(nD.SelectToken("data").SelectToken("getSingleIssueInfo").ToString)

            For Each V In issueVulns.scaVulnerabilities
                Dim parentLib$ = ""

                If IsNothing(issueVulns.sbom) = False Then
                    parentLib = issueVulns.sbom.libId
                Else
                    parentLib = V.libName + "|" + V.libVersion
                End If

                Dim a$ = qT(i.app.name) + "," + qT(parentLib) + "," + qT(V.libName + "@" + V.libVersion) + "," + qT(V.cve)
                '                Console.WriteLine(a)
                csV += a + vbCrLf
            Next

            Call streamWriterTxt(fileN, csV)
            Call streamWriterTxt(Path.Combine(ogDir, "issues_sofar.txt"), i.issueId)
skipIt:
        Next


    End Sub

    Public Sub readoutSub(args() As String)
        Dim allIssues = New List(Of issueS)
        Dim allApps As List(Of oxAppshort) = New List(Of oxAppshort)

        allApps = getAppListShort()
        Console.WriteLine("# of Applications: " + allApps.Count.ToString)

        ' if cached.. skip API calls by specifying # of response files in dir
        Dim rFiles As Integer = 0
        If argValue("numfiles", args) <> "" Then
            rFiles = Val(argValue("numfiles", args))
        End If

        If rFiles = 0 Then
            Call getAllIssues(True)
        Else
            numResponseFiles = rFiles
        End If
        'Console.WriteLine("# resp files: " + (numResponseFiles).ToString)
        allIssues = buildIssues("getIssuesShort.json", numResponseFiles)

        'Console.WriteLine("# of issues: " + allIssues.Count.ToString)
        With allIssues(0)
            Console.WriteLine(.severityChangedReason.Count.ToString)
        End With

        Console.WriteLine("Analyzing Severity Factors..")

        Dim allSFs As List(Of sevF) = New List(Of sevF)
        Dim numR As Integer = 0
        Dim numE As Integer = 0
        Dim numD As Integer = 0

        For Each I In allIssues
            For Each S In I.severityChangedReason
                Dim sNdx = getSFndx(allSFs, S.shortName)
                If sNdx = -1 Then
                    'new
                    Dim SF As sevF = New sevF
                    SF.shortName = S.shortName
                    SF.numOccurrences = 1
                    SF.changeCategory = S.changeCategory
                    SF.changeNumber = S.changeNumber
                    allSFs.Add(SF)
                Else
                    allSFs(sNdx).numOccurrences += 1
                End If
                Select Case S.changeCategory
                    Case "Damage"
                        numD += 1
                    Case "Reachable"
                        numR += 1
                    Case "Exploitable"
                        numE += 1
                End Select
            Next
        Next

        Console.WriteLine("# unique Severity Factors     : " + allSFs.Count.ToString)
        Console.WriteLine("# Severity Factors applied    : " + (numR + numE + numD).ToString)
        Console.WriteLine("# Reachable/Exploitable/Damage: " + numR.ToString + "/" + numE.ToString + "/" + numD.ToString)
        Console.WriteLine(vbCrLf + "Making reports..")

        Dim fileN$ = argValue("file", args)

        If Len(fileN) Then
            'safeKILL(fileN)
            Call issueDetailRpt(allIssues, fileN$)
            Call appTagRPT(allApps, Path.Combine(ogDir, fileN + "_tags.xlsx"))
            Call sfRoiRPT(allSFs, numR, numE, numD, Path.Combine(ogDir, fileN + "_SFroi.xlsx"))

        End If

    End Sub

    Public Sub tagFromXLS(args() As String)

        Dim fileN$ = argValue("file", args)

        Dim commitChanges As Boolean = False
        If LCase(argValue("commit", args)) = "true" Then commitChanges = True

        If fileN = "" Then
            Console.WriteLine("You must provide the filename of the Excel file to read from using --FILE")
            End
        Else
            If IO.File.Exists(fileN) = False Then
                Console.WriteLine("The file '" + fileN + "' does not exist")
                End
            End If
        End If

        Dim tagCol$ = argValue("tag", args)
        Dim appCol$ = argValue("appname", args)

        If tagCol = "" Or appCol = "" Then
            Console.WriteLine("You must provide an Excel column containing the --APPNAME and --TAG strings" + vbCrLf + "EXAMPLE:   --FILE example.xlsx --APPNAME A --TAG B" + vbCrLf + "-------     Would use Excel file 'example.xlsx' to tag Applications named in first column A with values in second column B")
            End
        End If


        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs
        '        Dim newrpt A
        rAR.s1 = appCol ' s1 = APP NAME.. s2 = APP ID
        rAR.s3 = tagCol ' J = field to map to s1 or s2
        rAR.someColl = New Collection
        rAR.someColl2 = New Collection
        rAR = newRpt.pullXLSfields(rAR, fileN)

        Console.WriteLine("# of Apps in column " + appCol + ": " + rAR.someColl.Count.ToString)

        Dim allApps As List(Of oxAppshort) = getAppListShort()
        'Console.WriteLine("# of Applications: " + allApps.Count.ToString)
        Dim allTags As List(Of oxTag) = getAllTags()
        Console.WriteLine("# of Tags: " + allTags.Count.ToString)

        Dim appCount As Integer = 0

        For appCount = 1 To rAR.someColl.Count
            Dim apP As oxAppshort = OX.returnAppShortByName(rAR.someColl(appCount), allApps)

            Dim actioN$ = ""

            If apP.appName = "" Then
                Console.WriteLine("Cannot find app: " + rAR.someColl(appCount))
                GoTo skipApp
            End If

            actioN = "APP: " + apP.appName + " [" + apP.appId + "] - "

            Dim appExist As Boolean = True
            Dim tagExist As Boolean = False

            Dim taG$ = rAR.someColl2(appCount)
            taG = Trim(taG)

            If taG = "" Then GoTo skipApp 'nothing requested to be tagged

            ' determine if tag needs to be added
            Dim tId$ = OX.returnTagId(taG, allTags)
            If tId$ = "" Then
                tagExist = False
                actioN += "ADD TAG [" + taG + "] TO MASTER LIST -"

            Else
                actioN += "TAG EXISTS ID: " + tId + " - "
                tagExist = True
            End If

            If commitChanges = True Then
                'here add tag to list
                'reload all tags
                'now we have id
                If tagExist = False Then
                    Call addTag(taG)
                    allTags = getAllTags()
                    tId$ = OX.returnTagId(taG, allTags)
                    If tId = "" Then
                        Console.WriteLine("Problem adding tag " + taG + " - ABORTING.. Fix this tag in the XLS (illegal char?)")
                        GoTo skipApp
                    Else
                        Console.WriteLine("New tag '" + taG + "' [" + tId + "] added")
                        actioN += " [ADDED] -"
                        tagExist = True
                    End If
                End If
            End If

            Dim tagExistOnApp As Boolean = False

            For Each appTag In apP.tags
                If appTag.tagId = tId Then
                    tagExistOnApp = True
                End If
            Next

            If tagExistOnApp = True Then
                actioN += "TAG ALREADY EXISTS ON APP"
                Console.WriteLine(actioN)
            Else
                ' here we add
                actioN += "COMMITTING TAG TO APP - "
                If commitChanges = False Then
                    Console.WriteLine(actioN + " COMMIT FALSE")
                Else
                    Dim modAppTag As editTagsRequestVARS = New editTagsRequestVARS
                    modAppTag.addedTagsIds.Add(tId)
                    modAppTag.appIds.Add(apP.appId)
                    'here add tag to app
                    Call setEditTagsVarsRequests(modAppTag)
                    Call setUpAPICall("modifyAppsTags")
                    Console.WriteLine(actioN + " COMPLETE")
                End If

            End If

skipApp:
        Next

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

    Public Function getOxConnectors() As List(Of connectorFamily)
        getOxConnectors = New List(Of connectorFamily)

        Dim jsoN$ = setUpAPICall("getConnectors",, True)
        getOxConnectors = OX.getConnectionsFromJson(jsoN)

    End Function

    Public Function getAppListIrrelevant() As List(Of oxAppIrrelevant)
        getAppListIrrelevant = New List(Of oxAppIrrelevant)
        '        Dim jsoN$ = ""
        '        jsoN = setUpAPICall("getAppsShort",, True)

        Call getAllIssues(, "getAppsIrrelevant", 500)

        'Dim allIssues As List(Of issueS) = New List(Of issueS)

        Dim currFile As Integer
        Dim cFile$
        Dim fileN$ = "getAppsIrrelevant.json"

        For currFile = 0 To numResponseFiles - 1
            cFile = Replace(fileN, ".json", currFile.ToString + ".json")
            Dim tempApps As List(Of oxAppIrrelevant) = New List(Of oxAppIrrelevant)
            Console.WriteLine("Deserializing JSON " + cFile)
            tempApps = OX.getAppIrrelevant(streamReaderTxt(cFile))

            If numResponseFiles = 0 Then
                getAppListIrrelevant = tempApps
            Else
                ' is there a better way? test 
                For Each T In tempApps
                    getAppListIrrelevant.Add(T)
                Next
            End If
        Next


        Console.WriteLine("Total # of files (500 apps per file): " + (numResponseFiles - 1).ToString)
        Console.WriteLine("# of Apps: " + getAppListIrrelevant.Count.ToString)



        '  getAppListShort = OX.getAppInfoShort(jsoN)
    End Function


    Public Function getAllIssues(Optional ByVal doShortIssues As Boolean = False, Optional ByVal differentCall$ = "", Optional ByVal diffOffset As Long = 0, Optional ByVal appName$ = "") As Integer
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
            Case "getAppsIrrelevant"
                Call setGetAppsIrrelevantVars(0) ', doShortIssues)
            Case "getIssuesMedium"
                Call setGetIssuesMediumVars(0, appName)

        End Select

        Console.WriteLine("Pulling first page > " + fileName + " > " + Replace(fileName, ".json", "0.json"))

        Call setUpAPICall(apiCall, Replace(fileName, ".json", "0.json"))

        Select Case apiCall
            Case "getIssues", "getIssuesShort"
                Dim respIssue As listIssues = New listIssues
                respIssue = OX.getListIssues(Replace(fileName, ".json", "0.json"))
                Console.WriteLine("Just got +" + Replace(fileName, ".json", "0.json"))

                With respIssue
                    Console.WriteLine("Total Issues: " + .totalIssues.ToString)
                    Console.WriteLine("Filtered Issues: " + .totalFilteredIssues.ToString)
                    Console.WriteLine("Offset: " + .offset.ToString)
                    If issueLimit <> .offset Then issueLimit = .offset
                End With

                numResponseFiles = 0
                'this was the first file

                ' Console.WriteLine("Starting loop From 0 to " + numIssueRequests(respIssue.totalFilteredIssues).ToString)

                Do Until numResponseFiles = numIssueRequests(respIssue.totalFilteredIssues) - 1
                    numResponseFiles += 1
                    Console.WriteLine("Calling OXAPI: " + (numResponseFiles + 1).ToString + " OF " + numIssueRequests(respIssue.totalFilteredIssues).ToString + " requests")
                    Call setGetIssuesVars(issueLimit * numResponseFiles, doShortIssues)
                    Call setUpAPICall(apiCall, Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
                    goUpALine()
                Loop

                Call setGetIssuesVars(0, doShortIssues)

            Case "getAppsIrrelevant"
                Dim appPageLimit As Long = 1000
                Dim respIssue As listApps = New listApps
                respIssue = OX.getListAppsPaging(Replace(fileName, ".json", "0.json"))
                With respIssue
                    Console.WriteLine("Total: " + .total.ToString)
                    ' Console.WriteLine("Total Filtered Apps: " + .totalFilteredApps.ToString)
                    Console.WriteLine("Total Irrelevant: " + .totalIrrelevantApps.ToString)
                    Console.WriteLine("Offset: " + .offset.ToString)
                    If appPageLimit <> .offset Then appPageLimit = .offset
                End With

                numResponseFiles = 0
                'this was the first file

                Dim issReq As Long = numIssueRequests(respIssue.total, appPageLimit)
                'Console.WriteLine("Limit = " + appPageLimit.ToString)

                Do Until numResponseFiles = issReq
                    numResponseFiles += 1
                    Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + issReq.ToString + " requests")
                    ' hardcoding

                    Call setGetAppsIrrelevantVars(appPageLimit * numResponseFiles)
                    Call setUpAPICall(apiCall, Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
                    goUpALine()
                Loop

                Call setGetAppsIrrelevantVars(0) ', doShortIssues)



            Case "getIssuesMedium"
                Dim respIssue As listIssues = New listIssues
                respIssue = OX.getListIssues(Replace(fileName, ".json", "0.json"))

                With respIssue
                    'Console.WriteLine("Total Issues: " + .totalIssues.ToString)
                    'Console.WriteLine("Filtered Issues: " + .totalFilteredIssues.ToString)
                    'Console.WriteLine("Offset: " + .offset.ToString)
                    If issueLimit <> .offset Then issueLimit = .offset
                End With

                numResponseFiles = 1
                'this was the first file

                Do Until numResponseFiles = numIssueRequests(respIssue.totalFilteredIssues)
                    numResponseFiles += 1
                    'Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + numIssueRequests(respIssue.totalFilteredIssues).ToString + " requests")
                    Call setGetIssuesMediumVars(issueLimit * numResponseFiles, appName) ', doShortIssues)
                    Call setUpAPICall(apiCall, Replace(fileName, ".json", numResponseFiles.ToString + ".json"))
                    ' goUpALine()

                Loop

                Call setGetIssuesMediumVars(0, appName) ', doShortIssues)


            Case "getAppsShort"
                Dim appPageLimit As Long = 1000
                Dim respApps As listApps = New listApps
                respApps = OX.getListAppsPaging(Replace(fileName, ".json", "0.json"))
                'Console.WriteLine("Limit = " + appPageLimit.ToString)
                With respApps
                    Console.WriteLine("Total: " + .total.ToString)
                    ' Console.WriteLine("Total Filtered Apps: " + .totalFilteredApps.ToString)
                    Console.WriteLine("Total Irrelevant: " + .totalIrrelevantApps.ToString)
                    'Console.WriteLine("Offset: " + .offset.ToString)
                    If appPageLimit <> .offset Then appPageLimit = .offset
                End With

                numResponseFiles = 0
                'this was the first file

                Dim issReq As Long = numIssueRequests(respApps.total, appPageLimit)
                'Console.WriteLine("Limit = " + appPageLimit.ToString)

                Do Until numResponseFiles = issReq
                    numResponseFiles += 1
                    Console.WriteLine("Calling OXAPI: " + numResponseFiles.ToString + " OF " + issReq.ToString + " requests")
                    ' hardcoding

                    Call setGetAppsShortVars(appPageLimit * numResponseFiles)
                    Console.WriteLine("OFFSET:" + (appPageLimit * numResponseFiles).ToString)
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

    Public Sub setIssueReqVars(issueId$, Optional ByVal fileN$ = "")
        Dim reqVars As newIssueDetailRequestVARS = New newIssueDetailRequestVARS(issueId)
        OX = New oxWrapper("", "")
        Dim newJson$ = OX.jsonGetNewIssueDetailVars(reqVars)
        'Console.WriteLine(newJson)
        If fileN = "" Then fileN = "getSingleIssue.variables.json"
        Call saveJSONtoFile(newJson, Path.Combine(pyDir, fileN))
    End Sub

    Public Sub setGetAppsShortVars(offSet As Integer)
        Dim toFileN$ = "getAppsShort.variables.json"
        Dim appReqVar As appsRequestVARS = New appsRequestVARS(offSet)
        appReqVar.offset = offSet
        Dim newJson$ = ""
        ' this is sloppy pls figure out why you did this, this way
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetAppsVars(appReqVar)
        newJson = "{" + Chr(34) + "getApplicationsInput" + Chr(34) + ":" + newJson + "}"
        Console.WriteLine(newJson)

        Call saveJSONtoFile(newJson, Path.Combine(pyDir, toFileN))

    End Sub
    Public Sub setGetAppsIrrelevantVars(offSet As Integer)
        Dim toFileN$ = "getAppsIrrelevant.variables.json"
        Dim appReqVar As appsIrrelevantRequestVARS = New appsIrrelevantRequestVARS(offSet)
        Dim newJson$ = ""
        ' this is sloppy pls figure out why you did this, this way
        OX = New oxWrapper("", "")
        newJson = OX.jsonGetIrrelevantAppsVars(appReqVar)
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
    Public Sub setGetIssuesMediumVars(offSet As Integer, Optional ByVal appName$ = "") ', Optional ByVal doIssuesShort As Boolean = False)
        Dim toFileN$ = "getIssuesMedium.variables.json"
        'If doIssuesShort = True Then toFileN = "getIssuesShort.variables.json"
        Dim newIssueVar As issueRequestVARS = New issueRequestVARS
        With newIssueVar

            .dateRange.from = 1
            .dateRange.to = dateToJS(Now)
            .getIssuesInput.limit = 1000 '10000 results in 12MB file and some timeouts
            .getIssuesInput.offset = offSet
            .getIssuesInput.filters.criticality = New List(Of String)
            .getIssuesInput.filters.criticality.Add("Appoxalypse")
            .getIssuesInput.filters.criticality.Add("Critical")
            .getIssuesInput.filters.criticality.Add("High")
            '.getIssuesInput.filters.criticality.Add("Medium")
            '.getIssuesInput.filters.criticality.Add("Critical") ',"High","Medium","Low","Info"
            If Len(appName) > 0 Then
                .getIssuesInput.conditionalFilters = New List(Of requestConditions)
                Dim rQ As requestConditions = New requestConditions
                rQ.condition = "OR"
                rQ.fieldName = "apps"
                rQ.values = New List(Of String)
                rQ.values.Add(appName)
                .getIssuesInput.conditionalFilters.Add(rQ)
                '                .getIssuesInput.
            End If
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



    'fuzzy matching
    Public Function GetSimilarity(string1 As String, string2 As String) As Single
        Dim dis As Single = ComputeDistance(string1, string2)
        Dim maxLen As Single = string1.Length
        If maxLen < string2.Length Then
            maxLen = string2.Length
        End If
        If maxLen = 0.0F Then
            Return 1.0F
        Else
            Return 1.0F - dis / maxLen
        End If
    End Function

    Private Function ComputeDistance(s As String, t As String) As Integer
        Dim n As Integer = s.Length
        Dim m As Integer = t.Length
        Dim distance As Integer(,) = New Integer(n, m) {}
        ' matrix
        Dim cost As Integer = 0
        If n = 0 Then
            Return m
        End If
        If m = 0 Then
            Return n
        End If
        'init1

        Dim i As Integer = 0
        While i <= n
            distance(i, 0) = System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While
        Dim j As Integer = 0
        While j <= m
            distance(0, j) = System.Math.Max(System.Threading.Interlocked.Increment(j), j - 1)
        End While
        'find min distance

        For i = 1 To n
            For j = 1 To m
                cost = (If(t.Substring(j - 1, 1) = s.Substring(i - 1, 1), 0, 1))
                distance(i, j) = Math.Min(distance(i - 1, j) + 1, Math.Min(distance(i, j - 1) + 1, distance(i - 1, j - 1) + cost))
            Next
        Next
        Return distance(n, m)
    End Function






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

    Public Function buildMediumIssues(fileN$, numFiles As Integer) As List(Of issuesMedium)
        Dim allIssues As List(Of issuesMedium) = New List(Of issuesMedium)

        Dim currFile As Integer
        Dim cFile$

        For currFile = 0 To numFiles - 1
            cFile = Replace(fileN, ".json", currFile.ToString + ".json")
            Dim tempIssues As List(Of issuesMedium) = New List(Of issuesMedium)
            'Console.WriteLine("Deserializing JSON " + cFile)
            tempIssues = OX.returnMediumIssues(streamReaderTxt(cFile))

            If numFiles = 0 Then
                allIssues = tempIssues
            Else
                ' is there a better way? test 
                For Each T In tempIssues
                    allIssues.Add(T)
                Next
            End If
        Next


        Console.WriteLine("Total # of files (" + issueLimit.ToString + " issues per file): " + (numFiles).ToString)
        Console.WriteLine("# of Issues: " + allIssues.Count.ToString)
        Return allIssues
    End Function


    Public Function buildShortIssues(fileN$, numFiles As Integer) As List(Of issueS)
        Dim allIssues As List(Of issueS) = New List(Of issueS)

        Dim currFile As Integer
        Dim cFile$

        For currFile = 0 To numFiles
            Console.WriteLine("Loading " + Replace(fileN, ".json", currFile.ToString + ".json"))
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

    Private Sub sfRoiRPT(allSFs As List(Of sevF), numR As Integer, numE As Integer, numD As Integer, fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        Console.WriteLine("Creating SF ROI XLS..")

        Dim bizPRI As Integer = 0
        With rAR
            .s1 = fileN
            .s2 = Path.Combine(ogDir, "roi_tags_template.xlsx")

            If IO.File.Exists(.s2) = False Then
                Console.WriteLine("ERROR: Unable to find file " + .s2 + " - aborting")
                End
            End If

            .someColl = New Collection
            .someColl2 = New Collection
            For Each S In allSFs
                .someColl.Add(S.shortName)
                .someColl2.Add(S.numOccurrences)
                If InStr(S.shortName, " Business Priority") > 0 Then bizPRI += S.numOccurrences
            Next
            .s3 = bizPRI.ToString
            .numeriC = numR
            .numeriC2 = numE
            .numeriC3 = numD
        End With

        Call newRpt.doROIrpt(rAR)
    End Sub

    Private Sub appTagRPT(allApps As List(Of oxAppshort), fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        rAR.someColl = New Collection
        rAR.s1 = fileN
        rAR.s3 = "TAGS"

        Call safeKILL(rAR.s1)

        Dim totalNum As Integer = 0
        Console.WriteLine("Calculating total number of tags")

        For Each aA In allApps
            For Each aTag In aA.tags
                '
                '                'If aTag.isOxTag = True Then
                totalNum += 1
                '                'End If
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

                'If aTag.isOxTag = True Then
                xls3d(roW, 0) = aTag.name
                xls3d(roW, 1) = aA.appName
                xls3d(roW, 2) = 1
                roW += 1
                'End If
            Next
        Next

        Console.WriteLine("Throwing object of " + (roW * 3).ToString + " elements at Excel..")
        rAR.booL1 = True ' create XLS not CSV

        If roW = 0 Then
            Console.WriteLine("Nothing to output")
            End
        End If

        Call newRpt.dump2XLS(xls3d, roW - 1, rAR,, True)

    End Sub
    Private Sub issueDetailRpt(ByRef allI As List(Of issueS), fileN$)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        Console.WriteLine("Building report for " + allI.Count.ToString + " issues")

        rAR.someColl = New Collection
        rAR.s1 = Path.Combine(ogDir, fileN + "_issues.xlsx")

        Call safeKILL(rAR.s1)
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
            Dim aA As issueS = allI(K)
            'xls3d(0, 0) = "some string"
            'Console.WriteLine(aA.issueId)
            xls3d(roW, 0) = aA.issueId
            xls3d(roW, 1) = aA.mainTitle
            xls3d(roW, 2) = aA.category.name
            Dim match = False
            If Mid(aA.app.name, 1, 1) <> "*" Then match = True
            xls3d(roW, 3) = CStr(match)
            Dim rID$ = ""
            rID = aA.sourceTools(0) ' = "WIZ" And Val(Mid(aA.resource.id, 1, 1)) > 0 Then rID = "WIZ" Else rID = aA.resource.id
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
            xls3d(roW, 13) = "" ' wtf is this.. fileNames(K)
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
            Dim aA As issueS = allI(K)
            'issue stuff
            Dim rID$ = ""
            rID = aA.sourceTools(0) ' = "WIZ" ' And Val(Mid(aA.resource.id, 1, 1)) > 0 Then rID = "WIZ" Else rID = aA.resource.id
            Dim issueID$ = aA.issueId
            Dim mainTitle$ = aA.mainTitle
            'Dim fN$ = fileNames(K)

            For Each SF In aA.severityChangedReason
                xls3d2(roW, 0) = SF.shortName
                xls3d2(roW, 4) = rID
                xls3d2(roW, 5) = SF.changeNumber
                xls3d2(roW, 6) = SF.reason
                xls3d2(roW, 7) = issueID
                xls3d2(roW, 8) = mainTitle
                xls3d2(roW, 9) = "" ' fN..wtf
                xls3d2(roW, 10) = 1
                xls3d2(roW, 14) = UCase(SF.changeCategory)
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
        rAR.s1 = Path.Combine(ogDir, fileN + "_sevFactors.xlsx")
        Call safeKILL(rAR.s1)


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

            'new 8/21
            '  .Add("COMPLIANCE")
            '  .Add("COMP_SECTION")

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

    Private Sub devDetailXLS(args() As String)
        Dim fileN$ = argValue("infile", args)
        If System.IO.File.Exists(fileN) = False Then
            Console.WriteLine("You must provide a file to parse and report on.")
        End If
        Dim userDetail As devDetail = New devDetail
        userDetail = OX.getCommittingUsers(fileN)

        ' Console.WriteLine("# of Committing Users: " + userDetail.getOrgUsersByOrgId.users.Count.ToString)

        Dim numYear As Integer = 0
        Dim num180 As Integer = 0
        Dim num90 As Integer = 0
        Dim num30 As Integer = 0

        Console.WriteLine("Org Name: " + userDetail.getOrgUsersByOrgId.display_name)
        Console.WriteLine("Dev Count/ Dev Count API: " + userDetail.getOrgUsersByOrgId.developersCount.ToString + "/" + userDetail.getOrgUsersByOrgId.developersCountAPI.ToString)
        Console.WriteLine("Users with 1yr commit info: " + userDetail.getOrgUsersByOrgId.users.Count.ToString)
        Console.WriteLine(vbCrLf + "# devs commit within: ")


        For Each U In userDetail.getOrgUsersByOrgId.users
            Dim uDate As DateTime = CDate(U.latestCommitDate)
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 365 Then
                'Console.WriteLine("Within year")
                numYear += 1
            End If
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 180 Then
                'Console.WriteLine("Within 180")
                num180 += 1
            End If
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 90 Then
                'Console.WriteLine("Within 90")
                num90 += 1
            End If
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 30 Then
                'Console.WriteLine("Within 30")
                num30 += 1
            End If
        Next
        Console.WriteLine("30 days: " + num30.ToString)
        Console.WriteLine("90 days: " + num90.ToString)
        Console.WriteLine("180 days:" + num180.ToString)
        Console.WriteLine("365 days:" + numYear.ToString)

        Dim outFile$ = argValue("file", args)

        If outFile$ <> "" Then
            Call userDetailRpt(userDetail.getOrgUsersByOrgId.users, outFile, args)
        End If

    End Sub

    Private Sub userDetailRpt(allUsers As List(Of committingUsers), fileN$, args() As String)
        Dim newRpt As customRPT = New customRPT
        Dim rAR As reportingArgs = New reportingArgs

        rAR.someColl = New Collection
        rAR.s1 = fileN

        Dim minDup As Single = 0.81
        minDup = Val(argValue("mindup", args))
        If minDup = 0 Then minDup = 0.81


        Dim last90Emails As Collection = New Collection

        Call safeKILL(fileN)

        'Console.WriteLine("Building output from results")

        With rAR.someColl
            .Add("EMAIL")
            .Add("NAME")
            .Add("COMMIT_DATE")
            .Add("COMMIT_REPO")
            .Add("#_DUP")
            .Add("POSS_MATCH")
            .Add("#_SVCACCT")
            .Add("#_30")
            .Add("#_90")
            .Add("#_180")
            .Add("#_YEAR")
        End With

        rAR.s3 = "DEV"

        Dim xls3d(allUsers.Count - 1, 10) As Object
        ''Dim numYear As Integer = 0
        'D 'im num180 As Integer = 0
        'D'im num90 As Integer = 0
        'Dim num30 As Integer = 0

        Dim tlNum90 As Integer = 0

        Dim roW As Long = 0
        For Each aA In allUsers
            xls3d(roW, 0) = aA.committerEmail
            xls3d(roW, 1) = aA.committerAuthor
            xls3d(roW, 2) = CDate(aA.latestCommitDate).ToShortDateString
            xls3d(roW, 3) = aA.link
            xls3d(roW, 4) = 0
            xls3d(roW, 5) = ""
            xls3d(roW, 6) = 0
            Dim uDate As DateTime = CDate(aA.latestCommitDate)

            Dim numYear As Integer = 0
            Dim num180 As Integer = 0
            Dim num90 As Integer = 0
            Dim num30 As Integer = 0
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 365 Then
                'Console.WriteLine("Within year")
                numYear += 1
            End If
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 180 Then
                'Console.WriteLine("Within 180")
                num180 += 1
            End If
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 90 Then
                'Console.WriteLine("Within 90")
                num90 += 1
                last90Emails.Add(aA.committerEmail)
                tlNum90 += 1
            End If
            If DateDiff(DateInterval.Day, uDate, Date.Now) < 30 Then
                'Console.WriteLine("Within 30")
                num30 += 1
            End If

            xls3d(roW, 7) = num30
            xls3d(roW, 8) = num90
            xls3d(roW, 9) = num180
            xls3d(roW, 10) = numYear
            roW += 1
        Next

        Console.WriteLine("Throwing object of " + (allUsers.Count * 11).ToString + " elements at Excel..")
        rAR.booL1 = True ' create XLS not CSV

        Call newRpt.dump2XLS(xls3d, allUsers.Count, rAR,, True)




        Console.WriteLine("Searching for duplicates..Match% set to >" + minDup.ToString)

        'svc account search
        'Dim svcA As Collection = spotServiceAccounts()

        ' duplicates report
        rAR = New reportingArgs
        rAR.someColl = New Collection

        With rAR.someColl
            .Add("FIRST_ENTRY(COMMIT_IN_90)")
            .Add("MATCH_%")
            .Add("POSSIBLE_DUPS(COMMIT_IN_90)")
            .Add("#_SUBTRACT")
        End With

        rAR.s1 = Replace(fileN, ".xlsx", "") + "_dups.xlsx"
        rAR.s3 = "DEVDUP"

        Dim devNum As Integer = 0
        Dim K As Integer = 0
        Dim loopNum As Integer = 0
        Dim listOfDups As List(Of duplicateDev) = New List(Of duplicateDev)

        Dim totalRows As Integer

        Dim alreadyCounted As Collection = New Collection
        Dim currEntry As duplicateDev = New duplicateDev
        Dim dup90 As Integer = 0

        For K = 0 To allUsers.Count - 1
            Dim addTo90 As Boolean = False

            If grpNDX(alreadyCounted, K.ToString) Then GoTo skipOuterLoop
            If isServiceAccount(LCase(allUsers(K).committerAuthor)) Then GoTo skipOuterLoop

            devNum = K
            currEntry = New duplicateDev
            currEntry.firstEntry = allUsers(K).committerAuthor
            currEntry.firstEmail = allUsers(K).committerEmail
            currEntry.dupUsers = New List(Of possibleMatch)

            If grpNDX(last90Emails, allUsers(K).committerEmail) > 0 Then
                currEntry.firstEmailIn90 = True
                addTo90 = True
            Else
                currEntry.firstEmailIn90 = False
            End If

            If K Mod 100 = 0 Then Console.WriteLine("[" + (K + 1).ToString + "/" + allUsers.Count.ToString + "] Checking for duplicates on " + currEntry.firstEntry)


            For loopNum = 0 To allUsers.Count - 1
                If loopNum = devNum Then GoTo skipMe

                If allUsers(loopNum).committerAuthor = currEntry.firstEntry Then
                    Dim newDup As possibleMatch = New possibleMatch
                    newDup.matchName = allUsers(loopNum).committerAuthor
                    newDup.matchEmail = allUsers(loopNum).committerEmail
                    newDup.matchNum = 1
                    If grpNDX(last90Emails, newDup.matchEmail) > 0 Then
                        newDup.in90 = True
                        addTo90 = True
                        'dup90 += 1
                        'Console.WriteLine("USER: " + newDup.matchEmail)
                    Else
                        newDup.in90 = False
                    End If
                    currEntry.dupUsers.Add(newDup)
                    alreadyCounted.Add(loopNum.ToString)
                    totalRows += 1
                    'Console.WriteLine("Added dup for " + currEntry.firstEmail)
                    GoTo skipMe
                End If


                Dim matchPCT As Single = 0
                matchPCT = GetSimilarity(LCase(currEntry.firstEntry), LCase(allUsers(loopNum).committerAuthor))

                If matchPCT >= minDup Then
                    ' Console.WriteLine("Testing " + allUsers(loopNum).committerAuthor)
                    Dim newDup As possibleMatch = New possibleMatch
                    newDup.matchName = allUsers(loopNum).committerAuthor
                    newDup.matchEmail = allUsers(loopNum).committerEmail
                    newDup.matchNum = Math.Round(matchPCT, 2)
                    If grpNDX(last90Emails, newDup.matchEmail) > 0 Then
                        newDup.in90 = True
                        addTo90 = True
                        'Console.WriteLine("USER: " + newDup.matchEmail)
                        'dup90 += 1
                    Else
                        newDup.in90 = False
                    End If
                    currEntry.dupUsers.Add(newDup)
                    alreadyCounted.Add(loopNum.ToString)
                    totalRows += 1
                End If
skipMe:
            Next
            If addTo90 = False Then GoTo skipOuterLoop

            If currEntry.dupUsers.Count > 0 Then
                listOfDups.Add(currEntry)
                dup90 += (currEntry.numCommittersIn90 - 1)
                totalRows += 1
            End If
skipOuterLoop:

        Next

        Console.WriteLine("Looking for service accounts")


        Dim svcAccount90 As Integer = 0

        currEntry = New duplicateDev
        currEntry.firstEntry = "SERVICE ACCOUNT"
        currEntry.firstEmail = "possible matches"
        currEntry.dupUsers = New List(Of possibleMatch)

        For Each U In allUsers
            'Dim founD As Boolean = False
            Dim a$ = LCase(U.committerAuthor)

            If isServiceAccount(U.committerAuthor) = True Then
                Console.WriteLine("SVC ACCOUNT: " + U.committerAuthor)
                Dim newEntry As possibleMatch = New possibleMatch
                newEntry.matchName = U.committerAuthor
                newEntry.matchEmail = U.committerEmail
                newEntry.matchNum = 1
                If grpNDX(last90Emails, newEntry.matchEmail) > 0 Then
                    newEntry.in90 = True
                    'Console.WriteLine("SVC ACCOUNT: " + newEntry.matchEmail)
                    svcAccount90 += 1
                    totalRows += 1
                Else
                    newEntry.in90 = False
                End If
                currEntry.dupUsers.Add(newEntry)
            End If ' founD = True

        Next
        'If currEntry.dupUsers.Count > 0 Then
        listOfDups.Add(currEntry)
        'End If

        If listOfDups.Count = 0 Then
            Console.WriteLine("No duplicates or service accounts detected")
            Exit Sub
        End If

        Dim dup3d(totalRows - 1, 3) As Object

        '        'Dim totalSubtract As Integer = 0
        '        .Add("FIRST_ENTRY_NAME")
        '        .Add("FIRST_EMAIL")
        '        .Add("POSSIBLE_DUPS(COMMIT_IN_90)")
        '        .Add("#_SUBTRACT")


        roW = 0
        For Each aA In listOfDups
            dup3d(roW, 0) = aA.firstEntry + "/" + aA.firstEmail + " (" + aA.firstEmailIn90.ToString + ")"
            dup3d(roW, 1) = aA.lowestMatchNum  'Else dup3d(roW, 1) = ""
            '            dup3d(roW, 1) = aA.firstEmailIn90.ToString
            Dim d$ = ""
            For Each duP In aA.dupUsers
                d$ += duP.matchName + "/" + duP.matchEmail + " (" + duP.in90.ToString + ")" + vbCrLf
            Next
            d = Mid(d, 1, Len(d) - 1)
            dup3d(roW, 2) = d
            If aA.firstEntry = "SERVICE ACCOUNT" Then dup3d(roW, 3) = aA.numCommittersIn90 Else dup3d(roW, 3) = aA.numCommittersIn90 - 1
            roW += 1
        Next

        Console.WriteLine("Throwing object of " + (totalRows * 5).ToString + " elements at Excel..")
        rAR.booL1 = True ' create XLS not CSV

        Call newRpt.dump2XLS(dup3d, totalRows, rAR)


        Console.WriteLine("=================================")
        Console.WriteLine("90 DAY SUMMARY")
        Console.WriteLine("=================================")
        Console.WriteLine("COMMITTING AUTHORS: " + tlNum90.ToString)
        Console.WriteLine("LIKELY DUPLICATES:  " + dup90.ToString)
        Console.WriteLine("SERVICE ACCOUNTS:   " + svcAccount90.ToString)
        Console.WriteLine("EST NET TOTAL DEVS: " + (tlNum90 - dup90 - svcAccount90).ToString)



    End Sub

    Private Function isServiceAccount(auth$) As Boolean
        Dim S As Collection = New Collection
        S = spotServiceAccounts()

        isServiceAccount = False
        auth = LCase(auth)

        If InStr(auth, "build") > 0 Then
            Dim K As Integer
            K = 1
        End If

        For Each a In S
            If InStr(auth, a) Then
                isServiceAccount = True
                'Console.WriteLine("Found " + a)
            End If
        Next

    End Function

    Private Function spotServiceAccounts() As Collection
        Dim svcA As Collection = New Collection
        svcA.Add("svc")
        svcA.Add("github")
        svcA.Add("bitbucket")
        svcA.Add("gitlab")
        svcA.Add("azure")
        svcA.Add("aws")
        svcA.Add("gcp")
        svcA.Add("service")
        svcA.Add("build")
        svcA.Add("automation")
        svcA.Add("unknown")
        svcA.Add("jenkins")
        svcA.Add("terraform")
        svcA.Add("root")
        svcA.Add("pipeline")
        svcA.Add("robot")
        svcA.Add("circleci")
        svcA.Add("snyk")
        svcA.Add("system")
        svcA.Add("ciserver")
        Return svcA
    End Function


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


    Public Sub createMappingJSON(args As String())
        Dim newRpt As customRPT = New customRPT

        Dim fileN$ = argValue("xls", args)
        If fileN = "" Then
            Console.WriteLine("You must include parameter --XLS (filename) as an input file")
            If IO.File.Exists(fileN) = False Then Console.WriteLine("Excel file '" + fileN + "' does not exist")
            End
        End If
        Dim jsonF = argValue("file", args)
        If jsonF = "" Then
            Console.WriteLine("You must include parameter --FILE (filename) as an output for the json file")
            End
        End If
        Dim mapT = argValue("maptype", args)
        If mapT = "" Then
            Console.WriteLine("You must include parameter --MAPTYPE (type)" + vbCrLf + "    Avail: BLACKDUCK")
            End
        End If
        Dim appCOL = argValue("appname", args)
        If appCOL = "" Then
            Console.WriteLine("You must include parameter --APPNAME (column)" + vbCrLf + "    Avail: BLACKDUCK")
            End
        End If
        Dim mapCOL = argValue("map", args)
        If mapCOL = "" Then
            Console.WriteLine("You must include parameter --MAP (column)" + vbCrLf + "    Avail: BLACKDUCK")
            End
        End If


        ' THIS NEEDS TO BE PARAMETERIZED --MAPCOLUMN --APPNAME or --APPID
        Dim rAR As reportingArgs = New reportingArgs
        rAR.s1 = appCOL '"A" ' s1 = APP NAME.. s2 = APP ID
        rAR.s3 = mapCOL '"I" ' J = field to map to s1 or s2
        rAR.someColl = New Collection
        rAR.someColl2 = New Collection
        rAR = newrpt.pullXLSfields(rAR, fileN)

        Console.WriteLine("Collection1: " + rAR.someColl.Count.ToString)
        Console.WriteLine("Collection2: " + rAR.someColl2.Count.ToString)

        '  Select Case LCase(mapT)
        '      Case "blackduck", "polaris", "veracode"
        Dim jsoN$ = OX.integrationMapping(rAR.someColl, rAR.someColl2, mapT)
                safeKILL(jsonF)
                saveJSONtoFile(jsoN, jsonF)
                Console.WriteLine("File saved: " + jsonF)
                '      Case "prob dont use this"
                '          Dim jsoN$ = OX.integrationMapping(rAR.someColl, rAR.someColl2, mapT)
                '          safeKILL(jsonF)
                '          saveJSONtoFile(jsoN, jsonF)
                '          Console.WriteLine("File saved: " + jsonF)
                '  End Select
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

    Public Sub consoleDump(allIssues As List(Of issuesMedium))
        If allIssues.Count = 0 Then Exit Sub

        Dim iC As issuesClass = New issuesClass(allIssues)

        Dim nStr$ = ""
        With iC
            Dim numS As Integer = .numSev("Appoxalypse")
            If numS Then nStr += numS.ToString + " Appoxalypse,"
            numS = .numSev("Critical")
            If numS Then nStr += numS.ToString + " Critical,"
            numS = .numSev("High")
            If numS Then nStr += numS.ToString + " High,"
        End With

        If Len(nStr) > 0 Then nStr = Mid(nStr, 1, Len(nStr) - 1)
        Dim lI$ = "======================================================================================================"

        Console.WriteLine(vbCrLf + "ISSUE SUMMARY: " + nStr + vbCrLf + iC.sumCats + vbCrLf + lI)

        Dim K As Integer

        For K = 0 To allIssues.Count - 1
            Dim I As issuesMedium = allIssues(K)
            Dim sTool$ = "OX"
            Console.WriteLine("ISSUE NAME: " + I.name + "    OX SEV: " + I.severity + "   ORIG SEV: " + I.originalToolSeverity)
            Console.WriteLine("CATEGORY: " + I.category.name + "       # OCCURRENCES: " + I.occurrences.ToString + "      " + sourceAndCommit(I.aggregations))
            Console.WriteLine(I.mainTitle + vbCrLf + vbCrLf + "DESCRIPTION:" + vbCrLf + I.secondTitle)
            Console.WriteLine(vbCrLf + "POLICY DESCRIPTION:" + vbCrLf + I.policy.detailedDescription + vbCrLf)
            For Each S In I.severityChangedReason
                Console.WriteLine(sfString(S))
            Next
            Console.WriteLine(lI + vbCrLf)

        Next


    End Sub
    Public Function sfString(SF As sevFactor) As String
        Dim pF$ = " "
        If 100 * SF.changeNumber > 0 Then pF = "+"
        If 100 * SF.changeNumber < 0 Then pF = "-"
        Return pF + "[" + SF.changeCategory + "] " + SF.shortName

    End Function
    Public Function sourceAndCommit(A As oxAgg) As String
        On Error Resume Next
        Dim s$ = ""
        Dim c$ = ""
        s$ = A.items(0).source
        c$ = A.items(0).commitBy

        If s$ = "" Then s$ = "OX"
        If c$ = "" Then c$ = "UNKNOWN"
        sourceAndCommit = "SOURCE: " + s$ + "         COMMIT: " + c$
    End Function


    Private Sub issueCacheLoader_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles issueCacheLoader.ProgressChanged
        Static oldP As Long = 0

        Dim pC As Long = e.ProgressPercentage
        If oldP = pC Then Exit Sub

        If pC Mod 0.5 = 0 Then Console.WriteLine("Current %: " + e.ProgressPercentage.ToString)

        oldP = pC
    End Sub

    Private Function getSFndx(sfList As List(Of sevF), sfText$) As Integer
        getSFndx = -1

        Dim ctR As Integer = 0
        For Each S In sfList
            If LCase(sfText) = LCase(S.shortName) Then
                getSFndx = ctR
            End If
            ctR += 1
        Next
    End Function

    Private Class sevF
        Public numOccurrences As Long
        Public changeNumber As Decimal
        ' Public reason As String
        Public shortName As String
        Public changeCategory As String
    End Class
    Private Class duplicateDev
        '.Add("FIRST_ENTRY_NAME")
        '.Add("FIRST_EMAIL")
        '.Add("MATCH_#")
        '.Add("MATCH_NAME")
        '.Add("MATCH_EMAIL")
        '.Add("#_DUPS")
        Public firstEntry As String
        Public firstEmail As String
        Public firstEmailIn90 As Boolean
        Public dupUsers As List(Of possibleMatch)

        Public Function lowestMatchNum() As Single
            Dim x As Single = 99
            For Each D In dupUsers
                If D.matchNum < x Then x = D.matchNum
            Next
            Return x
        End Function
        Public Function numCommittersIn90() As Integer
            Dim numU As Integer = 0
            If Me.firstEmailIn90 = True Then numU += 1
            For Each D In dupUsers
                If D.in90 = True Then numU += 1
            Next
            Return numU
        End Function
    End Class
    Private Class possibleMatch
        Public matchName As String
        Public matchEmail As String
        Public matchNum As Single
        Public in90 As Boolean
    End Class

End Module
