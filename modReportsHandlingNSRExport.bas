Attribute VB_Name = "modReportsHandlingNSRExport"
Option Explicit

'Dev note: Well, we need another line wrap in NSR export. I have to check this in
'an upcoming release. The feature is stable and usable - But not perfect at the
'moment ;)

Public Function GenerateNSRReport(ByRef strPluginReportFileName As String, _
                                    ByRef strReportTargetFileName As String) As String
    Dim i As Integer
    Dim strPluginReportFileContent As String
    Dim strLinesArray() As String
    Dim strLinesArrayLineCount As Integer
    Dim strLinesFileNameArray() As String
    Dim strTargetDirectoryName As String
    Dim strNSRReportFileContent As String
    
    'Load the report from the file
    strPluginReportFileContent = LoadReportFromFile(strPluginReportFileName)
    
    strLinesArray = Split(strPluginReportFileContent, vbNewLine, , vbBinaryCompare)
    
    strLinesArrayLineCount = UBound(strLinesArray)
    
    'Prepare the plugin file names for list generation
    For i = 0 To strLinesArrayLineCount
        If LenB(strLinesArray(i)) Then
            strLinesFileNameArray = Split(strLinesArray(i), ";", , vbBinaryCompare)
            If strLinesFileNameArray(1) = 1 Then
                Call ParseATKPlugin(ReadPluginFromFile(strLinesFileNameArray(0), application_plugin_directory))
                strNSRReportFileContent = strNSRReportFileContent & GenerateNBEReportLine & vbNewLine
            End If
        End If
    Next i
    
    'Prepare the report directory for the target
    strTargetDirectoryName = application_report_directory & "\" & Target
    
    'Create the report directory if it does not exist
    If Not (Dir$(strTargetDirectoryName, 16) <> "") Then
        MkDir (strTargetDirectoryName)
    End If
    
    'Write the HTMLListcontent to a HTML file. The file name can note be chosen at
    'this time. Such a feature should be added in a further release.
    On Error Resume Next ' Needed if there are no write permissions
    Open strReportTargetFileName For Output As #1
        Print #1, strNSRReportFileContent
    Close
    
    'Open the report after generation if wanted
    If application_report_open_enable = True Then
        Call ShellExecute(frmReport.hwnd, "Open", "notepad.exe", strReportTargetFileName, "", 1)
    End If
End Function

Public Function GenerateNBEReportLine() As String
    GenerateNBEReportLine = Target & "|" & "unresolved (" & plugin_port & "/" & plugin_protocol & ")|" & _
        "ATK" & plugin_id & "|REPORT|" & bug_description & ";;" & _
        Replace$(Mid$(LoadResponseFromFile(application_response_directory & "\" & Target & "-" & plugin_filename & ".txt"), 1, 1024), vbCrLf, ";", , , vbBinaryCompare) & _
        ";;" & bug_solution & ";;" & _
        "Risk factor : " & bug_severity & ";;" & "CVE : " & source_cve & ";" & _
        "BID : " & source_securityfocus_bid & ";"
End Function

