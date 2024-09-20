Sub Fossa_CountAndSortTeamCritHigh()
    ' this is designed to run AFTER Fossa_AppID_extract script, in the same book
    Dim dataSheet As Worksheet
    Dim newSheet As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim dict As Object
    Dim teamName As Variant
    Dim projectUrl As String
    Dim occurrences As Integer
    Dim newRow As Long
    Dim teamsColIndex As Integer
    Dim severityColIndex As Integer
    Dim projectUrlColIndex As Integer
    Dim lastRow As Long
    Dim i As Long

    ' Create a dictionary to store team data
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Reference the data and newly added summary sheet
    Set dataSheet = ThisWorkbook.Sheets("Summary")
    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newSheet.Name = "Data Summary_CritHigh"
    
    ' Find the column indexes
    teamsColIndex = Application.Match("Teams", dataSheet.Rows(1), 0)
    severityColIndex = Application.Match("Severity", dataSheet.Rows(1), 0)
    projectUrlColIndex = Application.Match("ProjectUrl", dataSheet.Rows(1), 0)
    
    ' Check if the columns were found
    If IsError(teamsColIndex) Then
        MsgBox "Column 'Teams' not found."
        Exit Sub
    ElseIf IsError(severityColIndex) Then
        MsgBox "Column 'Severity' not found."
        Exit Sub
    ElseIf IsError(projectUrlColIndex) Then
        MsgBox "Column 'ProjectUrl' not found."
        Exit Sub
    End If
    
    ' Calculate the last row with data
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, teamsColIndex).End(xlUp).Row
    
    ' Loop through each cell in the Teams column
    For Each cell In dataSheet.Range(dataSheet.Cells(2, teamsColIndex), dataSheet.Cells(lastRow, teamsColIndex))
        If cell.Value <> "" Then
            Dim currentSeverity As String
            currentSeverity = dataSheet.Cells(cell.Row, severityColIndex).Value
            If currentSeverity = "critical" Or currentSeverity = "high" Then
                teamName = cell.Value
                projectUrl = dataSheet.Cells(cell.Row, projectUrlColIndex).Value
                If Not dict.Exists(teamName) Then
                    dict.Add teamName, Array(projectUrl, 1)
                Else
                    occurrences = dict(teamName)(1) + 1
                    dict(teamName) = Array(projectUrl, occurrences)
                End If
            End If
        End If
    Next cell
    
    ' Write the headers
    newSheet.Cells(1, 1).Value = "Team"
    newSheet.Cells(1, 2).Value = "Project URL"
    newSheet.Cells(1, 3).Value = "Occurrences"
    
    ' Output the dictionary to the new sheet
    i = 2
    For Each teamName In dict.Keys
        newSheet.Cells(i, 1).Value = teamName
        newSheet.Cells(i, 2).Value = dict(teamName)(0)
        newSheet.Cells(i, 3).Value = dict(teamName)(1)
        i = i + 1
    Next teamName
    

    ' Sort by occurrences in descending order
    newSheet.Sort.SortFields.Add2 key:=Range("C:C"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With newSheet.Sort
        .SetRange newSheet.Range("A:C")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    MsgBox "Data summary created successfully."
End Sub
