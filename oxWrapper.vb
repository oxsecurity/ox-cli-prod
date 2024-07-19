Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System
Imports System.IO
Imports System.Text


Public Class policyWrapper
    Public policyTypes As List(Of String)



    Public Sub New()
        policyTypes = New List(Of String)
        With policyTypes
            .Add("Git Posture")
            .Add("Code Security")
            .Add("Secret Scan")
            .Add("Open Source Security")
            .Add("SBOM")
            .Add("Infrastructure as Code Scan")
            .Add("CICD Posture")
            .Add("Security Tool Coverage")
            .Add("Container Security")
            .Add("Artifact Integrity")
            .Add("Cloud Security")
        End With
    End Sub

    Public Function loadPolicy(policyName$) As List(Of oxPolicy)
        loadPolicy = New List(Of oxPolicy)

        Dim fileN$ = policyName + ".json"
        If System.IO.File.Exists(fileN) = False Then
            Console.WriteLine("Policy file '" + CurDir() + fileN + "' does not exist")
            Exit Function
        End If

        Dim jsoN$ = streamReaderTxt(fileN)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getPoliciesByCategoryIdAndProfileId").SelectToken("policies").ToString

        loadPolicy = JsonConvert.DeserializeObject(Of List(Of oxPolicy))(jsoN)

        For Each P In loadPolicy
            P.categorY = policyName
        Next

        Console.WriteLine(jsoN)

        Return loadPolicy
    End Function


End Class
Public Class oxWrapper
    Private apiK$
    Private hostnamE$
    Private isConnected As Boolean

    Public Sub New(urL$, apiKey$)
        'hostnamE = "https://api.cloud.ox.security" '/api/apollo-gateway
        hostnamE = urL
        apiK = apiKey
        isConnected = True
    End Sub

    Public Function getJSON(apiCall$) As Boolean
        On Error GoTo errorcatch

        Dim sInfo$ = "python"
        If osType = "MacOSX" Or osType = "Linux" Then sInfo = "python3"

        FileSystem.ChDir(pyDir)

        Dim startInfo As New ProcessStartInfo
        startInfo.FileName = sInfo
        startInfo.Arguments = "python_examp.py " + apiCall
        'startInfo.UseShellExecute = True

        ' Console.WriteLine("Executing>" + vbCrLf + startInfo.FileName + " " + startInfo.Arguments)

        Dim callPython As System.Diagnostics.Process = Process.Start(startInfo)
        ' Process.Start(startInfo)

        If callPython.WaitForExit(30000) = True Then
            getJSON = True
            FileSystem.ChDir(ogDir)
            Exit Function
        Else
            Console.WriteLine("API Process timeout")
            getJSON = False
            FileSystem.ChDir(ogDir)
            Exit Function
        End If

errorcatch:
        FileSystem.ChDir(ogDir)
        getJSON = False
        Console.WriteLine("ERROR: " & ErrorToString())
        'Return getAPIData("/api/apollo-gateway", True, "")
    End Function







    ' Writing the VARS files - consider bringing from main into wrapper
    ' set JSON object as *.variables.json file 
    Public Function jsonGetNewIssueDetailVars(issueInput As newIssueDetailRequestVARS) As String
        Dim jsoN$ = JsonConvert.SerializeObject(issueInput)
        Return jsoN
    End Function
    Public Function jsonGetNewTagVars(newTag As newTagRequestVARS) As String
        Dim jsoN$ = JsonConvert.SerializeObject(newTag)
        Return jsoN
    End Function

    Public Function jsonGetAppsVars(nAppList As appsRequestVARS) As String
        Dim jsoN$ = JsonConvert.SerializeObject(nAppList)
        Return jsoN
    End Function

    Public Function jsonGetIrrelevantAppsVars(nAppList As appsIrrelevantRequestVARS) As String
        Dim jsoN$ = JsonConvert.SerializeObject(nAppList)
        Return jsoN
    End Function

    Public Function jsonGetIssuesVars(giV As issueRequestVARS) As String
        Dim jsoN$ = JsonConvert.SerializeObject(giV)
        Return jsoN
    End Function

    Public Function jsonGetEditTagsVars(evR As editTagsRequestVARS) As String
        Dim fullReq As editTagsReq = New editTagsReq

        fullReq.input = evR

        Dim jsoN$ = JsonConvert.SerializeObject(fullReq)
        Return jsoN

    End Function

    Public Function returnIssues(json$) As List(Of issueS)
        returnIssues = New List(Of issueS)
        Dim nD As JObject = JObject.Parse(json)
        json = nD.SelectToken("data").SelectToken("getIssues").SelectToken("issues").ToString
        returnIssues = JsonConvert.DeserializeObject(Of List(Of issueS))(json)
    End Function

    Public Function returnMediumIssues(json$) As List(Of issuesMedium)
        returnMediumIssues = New List(Of issuesMedium)
        Dim nD As JObject = JObject.Parse(json)
        json = nD.SelectToken("data").SelectToken("getIssues").SelectToken("issues").ToString
        returnMediumIssues = JsonConvert.DeserializeObject(Of List(Of issuesMedium))(json)
    End Function


    Public Function returnShortIssues(json$) As List(Of issueShort)
        returnShortIssues = New List(Of issueShort)
        Dim nD As JObject = JObject.Parse(json)
        json = nD.SelectToken("data").SelectToken("getIssues").SelectToken("issues").ToString
        returnShortIssues = JsonConvert.DeserializeObject(Of List(Of issueShort))(json)
    End Function

    Public Function getTagId(jSon$) As String
        getTagId = ""
        Dim nD As JObject = JObject.Parse(jSon)
        jSon = nD.SelectToken("data").SelectToken("addTags").SelectToken("tags").ToString

        Dim tagObj As List(Of oxTag) = New List(Of oxTag)
        tagObj = JsonConvert.DeserializeObject(Of List(Of oxTag))(jSon)

        If tagObj.Count = 0 Then Exit Function 'add was unsuccessful
        getTagId = tagObj(0).tagId
    End Function

    Public Function getListIssues(fileN$) As listIssues
        getListIssues = New listIssues
        Dim jsoN$ = streamReaderTxt(fileN)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getIssues").ToString

        getListIssues = JsonConvert.DeserializeObject(Of listIssues)(jsoN)

    End Function


    Public Function getListAppsPaging(fileN$) As listApps
        getListAppsPaging = New listApps
        Dim jsoN$ = streamReaderTxt(fileN)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getApplications").ToString

        getListAppsPaging = JsonConvert.DeserializeObject(Of listApps)(jsoN)

    End Function

    Public Function getConnectionsFromJson(jsoN$) As List(Of connectorFamily)
        getConnectionsFromJson = New List(Of connectorFamily)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getConnectorsByFamily").ToString

        getConnectionsFromJson = JsonConvert.DeserializeObject(Of List(Of connectorFamily))(jsoN)

    End Function

    Public Function getAppInfoShort(jsoN$) As List(Of oxAppshort)
        getAppInfoShort = New List(Of oxAppshort)

        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getApplications").SelectToken("applications").ToString

        getAppInfoShort = JsonConvert.DeserializeObject(Of List(Of oxAppshort))(jsoN)
    End Function

    Public Function getAppIrrelevant(jsoN$) As List(Of oxAppIrrelevant)
        getAppIrrelevant = New List(Of oxAppIrrelevant)

        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getApplications").SelectToken("applications").ToString

        getAppIrrelevant = JsonConvert.DeserializeObject(Of List(Of oxAppIrrelevant))(jsoN)
    End Function

    Public Function getAllTags(jsoN$) As List(Of oxTag)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getAllTags").SelectToken("tags").ToString

        getAllTags = JsonConvert.DeserializeObject(Of List(Of oxTag))(jsoN)
    End Function

    Public Function returnTagId(taG$, tagList As List(Of oxTag)) As String
        returnTagId = ""
        taG = LCase(taG)

        For Each T In tagList
            If LCase(T.displayName) = taG Then
                Return T.tagId
            End If
        Next
    End Function

    Public Function getUserLogEntries(jsoN$) As List(Of oxUserLogEntry)
        getUserLogEntries = New List(Of oxUserLogEntry)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getLogs").ToString

        getUserLogEntries = JsonConvert.DeserializeObject(Of List(Of oxUserLogEntry))(jsoN)
    End Function

    Public Function getUserFilterEntries(jsoN$) As oxUserLogFilter
        getUserFilterEntries = New oxUserLogFilter
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getLogsFilters").ToString

        getUserFilterEntries = JsonConvert.DeserializeObject(Of oxUserLogFilter)(jsoN)
    End Function

End Class

Public Class oxUserLogFilter
    Public logTypes As List(Of oxLogFilterProps)
    Public logNames As List(Of oxLogFilterProps)
    Public userEmails As List(Of oxLogFilterProps)
End Class
Public Class oxLogFilterProps
    Public count As Integer
    Public label As String
End Class

Public Class oxUserLogEntry
    Public logType As String
    Public logName As String
    Public userEmail As String
    Public domain As String
    Public [date] As DateTime
End Class

Public Class oxAppIrrelevant
    '   "appId": "51548962",
    ' "appName": "WebGoat",
    ' "lastCodeChange": "1698193900331",
    ' "irrelevantReasons": [
    '     "No code changes in the last 6 months"
    ' ],
    ' "overrideRelevance": "default",
    ' "type": "GitLab",
    ' "fakeApp": false

    Public appId As String
    Public appName As String
    Public lastCodeChange As String
    Public irrelevantReasons As List(Of String)
    Public overrideRelevance As String
    Public [type] As String
    Public fakeApp As Boolean

    Public Sub New()
        irrelevantReasons = New List(Of String)
    End Sub

End Class

Public Class oxAppshort
    Public appId As String
    Public appName As String
    Public link As String
    Public tags As List(Of oxTag)

    Public Function tagExist(Optional ByVal tagId$ = "", Optional ByVal tagDisplayName$ = "") As Boolean
        tagExist = False

        If Len(tagDisplayName) Then GoTo doName

        If Len(tagId) = 0 Then
            Exit Function
        End If

        For Each T In Me.tags
            If T.tagId = tagId Then
                tagExist = True
                Exit Function
            End If
        Next

doName:

        For Each T In Me.tags
            If LCase(T.displayName) = LCase(tagDisplayName) Then
                tagExist = True
                Exit Function
            End If
        Next
    End Function
End Class

Public Class oxTag
    Public tagId As String
    Public name As String
    Public displayName As String
    Public tagType As String
    Public createdBy As String
    Public isOxTag As Boolean
End Class

Public Class oxPolicy
    '    "data": {
    '        "getPoliciesByCategoryIdAndProfileId": {
    '            "policies": [
    '                {
    '                    "id": "64f9c9a7f59f29539740bf86",
    '                    "policyId": "oxPolicy_securityCloudScan_100",
    '                    "ruleId": "oxRule_securityCloudScan_1",
    '                    "name": "Cloud security (CSPM) alerts should not occur",
    '                    "catId": 15,
    '                    "description": "CSPM (Cloud Security Posture Management) issues should not be present.",
    '                    "detailedDescription": "Cloud misconfigurations can lead to catastrophic security issues like breaches, exposure of data and exposure of infrastructure. In 2021, Codecov published a public Docker image containing static credentials for a GCP service account. These credentials were used to replace the install script hosted in Google Cloud Storage with a malicious script stealing environment variables.",
    '                    "severity": null,

    Public categorY As String
    Public id As String
    Public name As String
    Public description As String
    Public detailedDescription As String


End Class

Public Class gQLgetIssues_qry
    Public query As String
    'Public variables As getIssuesInput
End Class

Public Class gqlVars
    Public getIssuesInput As issueFilterClass
    Public Sub New()
        getIssuesInput = New issueFilterClass
    End Sub
End Class
Public Class issueFilterClass
    Public owners As List(Of String)
    Public offset As Integer
    Public limit As Integer
    Public filters As issueFilter
    Public sort As sortFilter
    Public dateRange As gqlDateRange
    Public isDemo As Boolean

    Public Sub New()
        offset = 0
        limit = 1000
        Me.owners = New List(Of String)
        Me.filters = New issueFilter
        Me.sort = New sortFilter
        Me.dateRange = New gqlDateRange
        isDemo = True

        With Me.filters.criticality
            .Add("Critical")
            .Add("High")
            .Add("Medium")
            .Add("Low")
            .Add("Info")
        End With
        With Me.sort
            .fields.Add("Severity")
            .order.Add("DESC")
        End With
        Me.dateRange.from = 1684993734665
        Me.dateRange.to = 9999999999999
    End Sub

End Class
Public Class gqlDateRange
    Public [from] As Long
    Public [to] As Long
End Class
Public Class sortFilter
    Public fields As List(Of String)
    Public order As List(Of String)
    Public Sub New()
        fields = New List(Of String)
        order = New List(Of String)
    End Sub
End Class
Public Class issueFilter
    Public criticality As List(Of String)
    Public Sub New()
        criticality = New List(Of String)
    End Sub
End Class
Public Class listIssues
    '    "totalIssues": 560,
    '      "totalFilteredIssues": 30,
    '      "totalResolvedIssues": 0,
    '      "offset": 50
    ' Public issues As List(Of oxIssueS)
    Public totalIssues As Long
    Public totalFilteredIssues As Long
    Public totalResolvedIssues As Long
    Public offset As Long
End Class
Public Class listApps
    '    "totalIssues": 560,
    '      "totalFilteredIssues": 30,
    '      "totalResolvedIssues": 0,
    '      "offset": 50
    ' Public issues As List(Of oxIssueS)
    Public total As Long
    Public totalFilteredApps As Long
    Public totalIrrelevantApps As Long
    Public offset As Long
End Class



Public Class singleIssue
    'dependencyGraph
    'sbom
    Public id As String
    Public issueId As String
    Public gptInfo As oxGPT
    Public isGPTFixAvailable As Boolean
    Public name As String
    Public scanId As String
    Public created As Long
    Public scanDate As Long
    Public mainTitle As String
    Public secondTitle As String
    Public description As String
    Public severity As String
    Public owners As List(Of String)
    Public ruleId As String
    Public originalToolSeverity As String
    Public exclusionCategory As String
    Public occurrences As Integer
    Public comment As String
    Public learnMore As List(Of String)
    Public exclusionId As String
    Public resource As issueResources
    Public isMonoRepoChild As Boolean
    Public monoRepoParent As String
    Public isFixAvailable As Boolean
    'Public prDeatils As String
    'public autofix
    Public extraInfo As List(Of kvPair)
    'Public lots of APP info here
    Public app As issueApp
    Public policy As oxPolicy
    Public category As oxCategory
    Public isPRAvailable As Boolean
    'public aggregations
    Public recommendation As String
    Public violationInfoTitle As String
    Public sourceTools As List(Of String)
    Public cwe As List(Of String)
    Public cweList As List(Of cweInfo)
    Public severityChangedReason As List(Of sevFactor)
    Public tickets As List(Of String)
    Public oscarData As List(Of oxOscar)
    Public Sub New()
        sourceTools = New List(Of String)
    End Sub
    Public Function numSevFactors(Optional ByVal numReachable As Boolean = False, Optional ByVal numExploitable As Boolean = False, Optional ByVal damagE As Boolean = False) As Integer
        numSevFactors = Me.severityChangedReason.Count

        If numReachable = False And numExploitable = False And damagE = False Then
            Exit Function
        End If
        numSevFactors = 0

        For Each SF In Me.severityChangedReason
            If numReachable = True And SF.changeCategory = "Reachable" Then numSevFactors += 1
            If numExploitable = True And SF.changeCategory = "Exploitable" Then numSevFactors += 1
            If damagE = True And SF.changeCategory = "Damage" Then numSevFactors += 1
        Next

    End Function
    Public Function increasedSev() As Boolean
        increasedSev = False
        If returnSeverityNum(Me.originalToolSeverity) < returnSeverityNum(Me.severity) Then increasedSev = True
    End Function
    Public Function decreasedSev() As Boolean
        decreasedSev = False
        If returnSeverityNum(Me.originalToolSeverity) > returnSeverityNum(Me.severity) Then decreasedSev = True
    End Function

End Class

Public Class issueResources
    Public id As String
    Public [type] As String
End Class
Public Class issueApp
    '  "app": {
    '      "id": "*aem-dispatcher",
    '      "name": "*aem-dispatcher",
    '      "businessPriority": 31.890410958904113,
    '      "type": "Git",
    '      "originBranchName": "",
    '      "repoId": null,
    Public id As String
    Public name As String
    Public businessPriority As Long

End Class
Public Class oxOscar
    Public id As String
    Public name As String
    Public description As String
    Public url As String
End Class
Public Class sevFactor
    Public changeNumber As Decimal
    Public reason As String
    Public shortName As String
    Public changeCategory As String
    Public extraInfo As List(Of extraInfoSF)
End Class
Public Class extraInfoSF
    Public [key] As String
    Public link As String
    Public snippet As oxSnippet
End Class
Public Class oxSnippet
    Public snippetLineNumber As Long
    Public language As String
    Public [text] As String
    Public filename As String
End Class
Public Class cweInfo
    Public name As String
    Public description As String
    Public url As String
End Class
Public Class oxGPT
    Public createdAt As String
    Public user As String
    Public gptResponse As String
End Class
Public Class kvPair
    Public key As String
    Public value As String
End Class
'Public Class oxCategory
'    Public name As String
'    Public categoryId As Integer
'End Class
'Public Class oxPolicy
'    Public id As String
'    Public name As String
'    Public detailedDescription As String
'End Class
Public Class issueShort
    Public id As String
    Public issueId As String
    Public scanId As String
    Public created As Long
    Public createdAt As Long
End Class
Public Class issueS
    ' {
    '     "id": "651110199778b62c06b261b5",
    '     "issueId": "584352228-oxPolicy_securityScan_55-CKV_AWS_20-false",
    '     "mainTitle": "AWS S3 Bucket is configured for PUBLIC read access",
    '     "secondTitle": "S3 buckets that are publically accessible are one of the leading causes of data exposure and loss. An S3 bucket with public read access provides attackers the ability to access stored data.",
    '     "name": "IaC issue",
    '     "created": 1695616812332,
    '     "scanId": "adb3ff84-85cd-4783-9a7d-3df18af8bda5",
    '     "owners": [
    '       "Kostya Zhuruev"
    '     ],
    '     "occurrences": 1,
    '     "comment": null,
    '     "severity": "Critical",
    '     "policy": {
    '     },
    '     "category": {
    '     },
    '     "app": {
    '     },
    '     "createdAt": 1693550422116

    Public id As String
    Public issueId As String
    Public mainTitle As String
    Public secondTitle As String
    Public name As String
    Public created As Long
    Public scanId As String
    Public owners As List(Of String)
    Public occurrences As Integer
    Public comment As String
    Public severity As String
    Public createdAt As Long
    Public policy As oxPolicy
    Public category As oxCategory
    Public app As oxApp

End Class
Public Class oxCategory

    Public name As String
    Public categoryId As Integer

End Class
Public Class oxApp

    Public id As String
    Public name As String
    Public businessPriority As Long
    Public [type] As String
    Public fakeApp As Boolean

End Class

Public Class issuesMedium
    Public id As String
    Public issueId As String
    Public mainTitle As String
    Public secondTitle As String
    Public name As String
    Public created As Long
    Public createdAt As Long
    Public scanId As String
    Public owners As List(Of String)
    Public occurrences As Integer
    Public severity As String
    Public originalToolSeverity As String
    Public aggregations As oxAgg
    Public policy As oxPolicy
    Public category As oxCategory
    Public app As oxApp
    Public severityChangedReason As List(Of sevFactor)
End Class
Public Class oxAgg
    Public summary As oxSumm
    Public [type] As String
    Public items As List(Of oxSource)
End Class
Public Class oxSource
    Public source As String
    Public commitBy As String
End Class
Public Class oxSumm
    Public summary As String
    Public comment As String
End Class

Public Class oxCats
    'blend of source and category
    Public name As String
    Public count As Integer
    Public numA As Integer
    Public numC As Integer
    Public numH As Integer
    Public numM As Integer
    Public numL As Integer
    Public numI As Integer
End Class

Public Class issuesClass
    Public allIssues As List(Of issuesMedium)

    Public Function numSev(criticalitY$) As Integer
        numSev = 0
        For Each I In allIssues
            Dim seV$ = LCase(I.severity)
            If seV = LCase(criticalitY) Then numSev += 1
        Next
    End Function

    Public Function categorieS() As List(Of oxCats)
        categorieS = New List(Of oxCats)
        Dim unqList As Collection = New Collection
        For Each I In allIssues
            If grpNDX(unqList, I.category.name) = 0 Then
                unqList.Add(I.category.name)
            End If
        Next
        For Each N In unqList
            Dim oC As oxCats = New oxCats
            oC.name = N
            categorieS.Add(oC)
        Next

        'now add em all up
        For Each O In categorieS
            For Each I In allIssues
                If O.name = I.category.name Then
                    O.count += 1
                    Select Case (LCase(I.severity))
                        Case "info"
                            O.numI += 1
                        Case "low"
                            O.numL += 1
                        Case "medium"
                            O.numM += 1
                        Case "high"
                            O.numH += 1
                        Case "critical"
                            O.numC += 1
                        Case "appoxalypse"
                            O.numA += 1
                    End Select
                End If
            Next
        Next
    End Function

    Public Function sumCats() As String
        sumCats = ""
        Dim allC As List(Of oxCats) = Me.categorieS

        For Each C In allC
            Dim newStr$ = ""
            sumCats += C.name + ": "
            If C.numA Then newStr += C.numA.ToString + " Appoxalypse,"
            If C.numC Then newStr += C.numC.ToString + " Critical,"
            If C.numM Then newStr += C.numM.ToString + " Medium,"
            If C.numL Then newStr += C.numL.ToString + " Low,"
            If C.numI Then newStr += C.numI.ToString + " Info,"

            sumCats += C.count.ToString + vbCrLf 'Mid(newStr, 1, Len(newStr) - 1) + ")" + vbCrLf
        Next

    End Function

    Public Sub New(issues As List(Of issuesMedium))
        allIssues = issues
    End Sub
End Class

' Group these together

Public Class newIssueDetailRequestVARS
    '    {
    '  "getSingleIssueInput": {
    '    "issueId": "584352228-oxPolicy_securityScan_55-CKV_AWS_20-false"
    '  }
    '}
    Public getSingleIssueInput As nidWrap1

    Public Sub New(issueId$)
        getSingleIssueInput = New nidWrap1
        getSingleIssueInput.issueId = issueId
    End Sub
End Class
Public Class nidWrap1
    Public issueId As String
End Class

Public Class newTagRequestVARS
    '    {
    '  "input": {
    '    "tagsInput": [
    '      {
    '        "displayName": "zzz2zzz",
    '        "name": "zzzz2zz",
    '        "tagType": "simple"
    '      }
    '    ]
    '  }
    '}
    ' mirrors layers of objects here to achieve desired serialization
    Public [input] As ntrWrap1
    Public Sub New(Optional ByVal dN$ = "", Optional ByVal nA$ = "", Optional ByVal tT$ = "")
        input = New ntrWrap1
        If dN <> "" Then
            If tT = "" Then tT = "simple"
            With input.tagsInput
                Dim nT As ntrVars = New ntrVars
                nT.displayName = dN
                nT.name = nA
                nT.tagType = tT
                .Add(nT)
            End With
        End If
    End Sub
End Class

Public Class ntrWrap1
    Public tagsInput As List(Of ntrVars)
    Public Sub New()
        tagsInput = New List(Of ntrVars)
    End Sub
End Class
Public Class ntrVars
    Public displayName As String
    Public name As String
    Public tagType As String
End Class
Public Class appsIrrelevantRequestVARS
    Public getApplicationsInput As appsIrrelevantRequestFields

    Public Sub New(offset As Long)
        getApplicationsInput = New appsIrrelevantRequestFields(offset)
    End Sub
End Class
Public Class appsIrrelevantRequestFields
    '    {
    '  "getApplicationsInput": {
    '    "applicationFilters": [
    '      "Irrelevant"
    '    ],
    '    "irrelevancyFilters": [],
    '    "filters": {},
    '    "offset": 0,
    '    "limit": 50,
    '    "search": "",
    '    "orderBy": {
    '      "direction": "DESC",
    '      "field": "Info"
    '    }
    '  }
    '}
    Public applicationFilters As List(Of String)
    Public irrelevancyFilters As List(Of String)
    Public offset As Long
    Public limit As Long
    Public search As String
    Public orderBy As orderByClause


    Public Sub New(offS As Long)
        offset = offS
        limit = 500
        search = ""

        orderBy = New orderByClause
        irrelevancyFilters = New List(Of String)
        applicationFilters = New List(Of String)
        applicationFilters.Add("Irrelevant")

        orderBy.direction = "DESC"
        orderBy.field = "Info"
    End Sub


End Class
Public Class orderByClause
    Public direction As String
    Public field As String
End Class
Public Class appsRequestVARS
    '    {"getApplicationsInput": {
    '  "offset": 0,
    '  "limit": 200000}}

    Public offset As Long
    Public limit As Long

    Public Sub New(offS As Long)
        offset = offS
        limit = 500
    End Sub
End Class

Public Class requestConditions
    Public condition As String
    Public fieldName As String
    Public values As List(Of String)
End Class
Public Class issueRequestVARS
    '    {"getIssuesInput": {"owners": [],"offset": 0,"limit": 1000,"filters": {"criticality": ["Critical","High","Medium","Low","Info"]}
    '    ',"sort": {"fields": ["Severity"],"order": ["DESC"]},"dateRange": {"from": 1684993734665,"to": 1685598534665}},
    ' "isDemo" true}
    '
    Public getIssuesInput As irvGII
    Public sort As irvSORT
    Public dateRange As irvDR
    Public isDemo As Boolean


    Public Sub New()
        isDemo = True
        getIssuesInput = New irvGII
        sort = New irvSORT
        dateRange = New irvDR

        With getIssuesInput
            .limit = 30
            .offset = 0
            .owners = New List(Of String)
            .filters = New irvFIL
            .filters.criticality = New List(Of String)
            .filters.criticality.Add("Appoxalypse")
            .filters.criticality.Add("Critical")
            .filters.criticality.Add("High")
            .filters.criticality.Add("Medium")
            .filters.criticality.Add("Low")
            .filters.criticality.Add("Info")
        End With

        With sort
            .fields = New List(Of String)
            .fields.Add("Severity")
            .order = New List(Of String)
            .order.Add("DESC")
        End With

        With dateRange
            .from = 0
            .to = dateToJS(Now)
        End With
    End Sub
End Class

Public Class oxConnectors
    '  "getConnectorsByFamily" [
    '  {
    '    "family": "SourceControl",
    '    "familyDisplayName": "Source Control",
    '    "connectors": [
    '      {
    '        "connector": {
    '          "id": "1",
    '          "name": "GitHub",
    '          "displayName": "GitHub",
    '          "description": "GitHub, Inc. is a provider of Internet hosting for software development and version control using Git. It offers the distributed version control and source code management functionality of Git, plus its own features",
    '          "credentialsTypes": [
    '            "GitHubApp",
    '            "IdentityProvider",
    '            "Token"
    '          ]
    '        }
    '      },

    Public getConnectorsByFamily As List(Of connectorFamily)
    Public Sub New()
        getConnectorsByFamily = New List(Of connectorFamily)
    End Sub
End Class
Public Class connectorFamily
    Public family As String
    Public familyDisplayName As String
    Public connectors As List(Of oxConnector)
End Class
Public Class oxConnector
    Public connector As oxConnection
    Public Sub New()
        connector = New oxConnection
    End Sub
End Class
Public Class oxConnection
    Public id As String
    Public name As String
    Public displayName As String
    Public description As String
    Public isConfigured? As Boolean
    Public credentialsTypes As List(Of String)
    Public Sub New()
        credentialsTypes = New List(Of String)
        isConfigured = False
    End Sub
End Class
Public Class editTagsReq
    Public [input] As editTagsRequestVARS
End Class

Public Class editTagsRequestVARS
    '    {
    '  "input": {
    '    "addedTagsIds": [
    '      "ea5b86c0-908d-4c04-92d6-32b267f6bdb5"
    '    ],
    '    "removedTagsIds": [],
    '    "appIds": [
    '      "*Bitbucket-Settings (oxsecurity)",
    '      "{a6d51cf9-4029-4163-89fd-97987351d81d}"
    '    ]
    '  }
    '}

    Public addedTagsIds As List(Of String)
    Public removedTagsIds As List(Of String)
    Public appIds As List(Of String)

    Public Sub New()
        addedTagsIds = New List(Of String)
        removedTagsIds = New List(Of String)
        appIds = New List(Of String)

    End Sub
End Class








Public Class irvGII
    Public owners As List(Of String)
    Public offset As Integer
    Public limit As Integer
    Public filters As irvFIL
    Public conditionalFilters As List(Of requestConditions)
End Class
Public Class irvFIL
    Public criticality As List(Of String)
End Class
Public Class irvSORT
    Public fields As List(Of String)
    Public order As List(Of String)
End Class
Public Class irvDR
    Public from As Long
    Public [to] As Long
End Class

