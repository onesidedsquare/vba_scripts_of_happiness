Sub Fossa_CVEReport()
    '####################################################
    ' RUN Fossa_APPID_extract.vb FIRST
    ' Select/copy out the CVE from Summary
    ' A dialog box will pop up asking for what to report on, provide the box with your CVE
    '####################################################
    Dim dataSheet As Worksheet
    Dim newSheet As Worksheet
    Dim cell As Range
    Dim cveColumn As Range
    Dim headers As Variant
    Dim searchText As String
    Dim uniqueFixes As Object
    Dim uniqueTeams As Object
    Dim teamProjectDict As Object
    Dim uniqueCVEs As Object
    Dim summaryRow As Long
    Dim newRow As Long
    Dim fixValue As String
    Dim teamCell As Range
    Dim projectCell As Range
    Dim vulnCell As Range
    Dim version As String
    Dim completeFix As String
    Dim epssScore As String
    Dim epssPercentile As String
    Dim uniqueVer As Object
    Dim severityColumn As Range
    Dim severity As String
    
    ' Prompt the user for text input
    searchText = InputBox("Enter the CVE to search for:", "Search Text")

    If searchText = "" Then
        MsgBox "No search text entered. Exiting script."
        Exit Sub
    End If

    Set dataSheet = ThisWorkbook.Sheets("Summary")

    Set cveColumn = dataSheet.Rows(1).Find("cve", LookIn:=xlValues, lookat:=xlWhole)
    If cveColumn Is Nothing Then
        MsgBox "Column 'cve' not found in the 'vulnerability' tab."
        Exit Sub
    End If

    ' Find the column with the header "severity"
    Set severityColumn = dataSheet.Rows(1).Find("severity", LookIn:=xlValues, lookat:=xlWhole)
    If severityColumn Is Nothing Then
        MsgBox "Column 'severity' not found in the 'vulnerability' tab."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set newSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Name = searchText & " Findings Report"
    
    Set uniqueFixes = CreateObject("Scripting.Dictionary")
    Set uniqueTeams = CreateObject("Scripting.Dictionary")
    Set teamProjectDict = CreateObject("Scripting.Dictionary")
    Set uniqueCVEs = CreateObject("Scripting.Dictionary")
    Set uniqueVer = CreateObject("Scripting.Dictionary")

    For Each cell In dataSheet.Columns(cveColumn.Column).Cells
        If Trim(cell.Value) = searchText Then
            ' Check if severity is not equal to "unknown"
            Set severityCell = dataSheet.Cells(cell.Row, severityColumn.Column)
            If Not severityCell Is Nothing And severityCell.Value <> "unknown" Then
                ' Handle potential errors when searching for team and project
                On Error Resume Next
                Set teamCell = dataSheet.Cells(cell.Row, Application.Match("teams", dataSheet.Rows(1), 0))
                Set projectCell = dataSheet.Cells(cell.Row, Application.Match("project", dataSheet.Rows(1), 0))
                Set vulnCell = dataSheet.Cells(cell.Row, Application.Match("version", dataSheet.Rows(1), 0))
                Set sevCell = dataSheet.Cells(cell.Row, Application.Match("severity", dataSheet.Rows(1), 0))
                On Error GoTo 0
    
                If Not teamCell Is Nothing And Not projectCell Is Nothing Then
                    If teamCell.Value <> "" And projectCell.Value <> "" Then
                        If Not uniqueTeams.Exists(teamCell.Value & "_" & projectCell.Value) Then
                            uniqueTeams.Add teamCell.Value & "_" & projectCell.Value, True
                            teamProjectDict.Add teamCell.Value & "_" & projectCell.Value, teamCell.Value & "_" & projectCell.Value & "_" & vulnCell.Value
                        End If
                    End If
                End If
                
                epssScore = dataSheet.Cells(cell.Row, Application.Match("epssScore", dataSheet.Rows(1), 0)).Value
                epssPercentile = dataSheet.Cells(cell.Row, Application.Match("epssPercentile", dataSheet.Rows(1), 0)).Value
    
                If Not uniqueCVEs.Exists(searchText) Then
                    uniqueCVEs.Add searchText, searchText & "_" & epssScore & "_" & epssPercentile
                End If
                
                version = dataSheet.Cells(cell.Row, Application.Match("version", dataSheet.Rows(1), 0)).Value
                completeFix = dataSheet.Cells(cell.Row, Application.Match("completeFix", dataSheet.Rows(1), 0)).Value
                
                If Not uniqueVer.Exists(version) Then
                    uniqueVer.Add version, version & "_" & completeFix
                End If
             End If
        End If
    Next cell

    ' Display unique teams and their projects
    newRow = 1
    newSheet.Cells(newRow, 1).Value = "Teams"
    newSheet.Cells(newRow, 2).Value = "Project"
    newSheet.Cells(newRow, 3).Value = "Vuln Ver"
    newSheet.Cells(newRow, 6).Value = "CVEs"
    newSheet.Cells(newRow, 7).Value = "EpssScore"
    newSheet.Cells(newRow, 8).Value = "EpssPercentile"
    
    restartRow = newRow + 1
    newRow = newRow + 1

    For Each Key In teamProjectDict.Keys
        newSheet.Cells(newRow, 1).Value = Split(teamProjectDict(Key), "_")(0)
        newSheet.Cells(newRow, 2).Value = Split(teamProjectDict(Key), "_")(1)
        newSheet.Cells(newRow, 3).Value = Split(teamProjectDict(Key), "_")(2)
        newRow = newRow + 1
    Next Key
    
    For Each Key In uniqueCVEs.Keys
        cve = Split(uniqueCVEs(Key), "_")(0)
        epssScore = Split(uniqueCVEs(Key), "_")(1)
        epssPercentile = Split(uniqueCVEs(Key), "_")(2)
        newSheet.Cells(restartRow, 6).Value = cve
        newSheet.Cells(restartRow, 7).Value = epssScore
        newSheet.Cells(restartRow, 8).Value = epssPercentile
        restartRow = restartRow + 1
    Next Key
    
    restartRow = restartRow + 1
    newSheet.Cells(restartRow, 5).Value = "Vuln Version"
    newSheet.Cells(restartRow, 6).Value = "Complete Fix"
    
    For Each Key In uniqueVer.Keys
        restartRow = restartRow + 1
        newSheet.Cells(restartRow, 5).Value = Split(uniqueVer(Key), "_")(0)
        newSheet.Cells(restartRow, 6).Value = Split(uniqueVer(Key), "_")(1)
    Next Key
    
    newSheet.Columns("G:H").NumberFormat = "0.00%"
    Application.ScreenUpdating = True
    MsgBox "Unique rows and team-project pairs with CVE '" & searchText & "' extracted to new sheet."
    
End Sub

