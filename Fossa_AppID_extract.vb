Sub Fossa_AppID_extract()
    Dim dataSheet As Worksheet
    Dim newSheet As Worksheet
    Dim identifiersDict As Object
    Dim identifiersArray As Variant
    Dim cell As Range
    Dim identifier As Variant
    Dim newRow As Long
    Dim teamColumn As Range
    Dim summaryRow As Long
    Dim headers As Variant
    Dim i As Long
    
    ' Define the array of identifiers
    identifiersArray = Array("1", "2", "3")

    ' Create a dictionary to store the identifiers
    Set identifiersDict = CreateObject("Scripting.Dictionary")
    For Each identifier In identifiersArray
        identifiersDict.Add identifier, True
    Next identifier

    ' This is designed to run off FOSSA export csv, with the tab vulnerablity 

    Set dataSheet = ThisWorkbook.Sheets("vulnerability")
        If dataSheet Is Nothing Then
        MsgBox "Required sheet 'vulnerability' not found in the workbook."
        Exit Sub
    End If

    Set newSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Name = "Summary"
    
    ' Set the headers for the Summary tab
    headers = Array("issueId", "title", "cvssVector", "cvss", "cve", "severity", "dependency", "packageLocator", _
                    "version", "details", "cwes", "remediation", "affectedVersionRanges", "patchedVersionRanges", _
                    "project", "projectUrl", "status", "depth", "usage", "teams", "references", "scannedAt", _
                    "analyzedAt", "epssScore", "epssPercentile", "exploitMaturity", "completeFix", "firstFoundAt")
    
    ' Write headers to the first row of Summary tab
    For i = LBound(headers) To UBound(headers)
        newSheet.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ' Find the column index of the "teams" column
    Set teamColumn = dataSheet.Rows(1).Find("teams", LookIn:=xlValues, lookat:=xlWhole)
    
    If teamColumn Is Nothing Then
        MsgBox "Column 'teams' not found in the 'vulnerability' tab."
        Exit Sub
    End If
    
    summaryRow = 2 ' Start from the second row after headers
    For Each cell In dataSheet.Columns(teamColumn.Column).Cells
        If CellContainsIdentifier(cell.Value, identifiersDict) Then
            newRow = summaryRow
            dataSheet.Rows(cell.Row).Copy Destination:=newSheet.Rows(newRow)
            summaryRow = summaryRow + 1
        End If
    Next cell
End Sub

Function CellContainsIdentifier(cellValue As String, identifiersDict As Object) As Boolean
    Dim identifier As Variant
    For Each identifier In identifiersDict.Keys
        If InStr(1, cellValue, identifier) > 0 Then
            CellContainsIdentifier = True
            Exit Function
        End If
    Next identifier
    CellContainsIdentifier = False
End Function
