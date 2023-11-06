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

        Dim startInfo As New ProcessStartInfo
        startInfo.FileName = "python"
        startInfo.Arguments = "./python/python_examp.py ./python/" + LCase(apiCall)

        Dim callPython As System.Diagnostics.Process = Process.Start(startInfo)
        ' Process.Start(startInfo)
        If callPython.WaitForExit(30000) = True Then
            getJSON = True
            Exit Function
        Else
            Console.WriteLine("API Process timeout")
            getJSON = False
            Exit Function
        End If
errorcatch:
        getJSON = False
        Console.WriteLine("ERROR: " & ErrorToString())
        'Return getAPIData("/api/apollo-gateway", True, "")
    End Function

    Public Function jsonGetIssuesVars(giV As issueRequestVARS) As String
        Dim jsoN$ = JsonConvert.SerializeObject(giV)
        Return jsoN
    End Function


    Public Function returnIssues(json$) As List(Of issueS)
        returnIssues = New List(Of issueS)
        Dim nD As JObject = JObject.Parse(json)
        json = nD.SelectToken("data").SelectToken("getIssues").SelectToken("issues").ToString
        returnIssues = JsonConvert.DeserializeObject(Of List(Of issueS))(json)
    End Function


    Public Function getListIssues(fileN$) As listIssues
        getListIssues = New listIssues
        Dim jsoN$ = streamReaderTxt(fileN)
        Dim nD As JObject = JObject.Parse(jsoN)
        jsoN = nD.SelectToken("data").SelectToken("getIssues").ToString

        getListIssues = JsonConvert.DeserializeObject(Of listIssues)(jsoN)

    End Function




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
Public Class irvGII
    Public owners As List(Of String)
    Public offset As Integer
    Public limit As Integer
    Public filters As irvFIL
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

