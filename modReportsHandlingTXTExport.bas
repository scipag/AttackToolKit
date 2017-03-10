Attribute VB_Name = "modReportsHandlingTXTExport"
Option Explicit

Dim strTempReportPluginData As String

Public Sub WriteTXTReportToFile(ByRef strReportSourceFilename As String, _
                                ByRef strReportDestinationFileName As String)
    
    Dim strTargetDirectoryName As String
    Dim strReportContent As String
    
    'Prepare the report directory for the target
    strTargetDirectoryName = application_report_directory & "\" & Target
    
    'Create the report directory if it does not exist
    On Error Resume Next ' Needed if there are no write permissions
    If Not (Dir$(strTargetDirectoryName, 16) <> "") Then
        MkDir (strTargetDirectoryName)
    End If
    
    strReportContent = GenerateTXTReport(strReportSourceFilename)
    
    Open strReportDestinationFileName For Output As #1
        Print #1, strReportContent
    Close
    
    'Open the report after generation if wanted
    If application_report_open_enable = True Then
        Call ShellExecute(frmReport.hwnd, "Open", "notepad.exe", strReportDestinationFileName, "", 1)
    End If
End Sub

Public Function GenerateTXTReport(ByRef strPluginReportFileName As String) As String
    Dim i As Integer
    Dim strLinesArray() As String
    Dim strLinesArrayLineCount As Integer
    Dim strLinesFileNameArray() As String
    Dim strLinesFileNames As String
    
    'Reset the old report data
    strTempReportPluginData = vbNullString

    strLinesArray = Split(LoadReportFromFile(strPluginReportFileName), vbNewLine, , vbBinaryCompare)
    
    strLinesArrayLineCount = UBound(strLinesArray)
    
    'Prepare the plugin file names for list generation
    For i = 0 To strLinesArrayLineCount
        If LenB(strLinesArray(i)) Then
            strLinesFileNameArray = Split(strLinesArray(i), ";", , vbBinaryCompare)
            If strLinesFileNameArray(1) = 1 Then
                strLinesFileNames = strLinesFileNames & ";" & strLinesFileNameArray(0)
            End If
        End If
    Next i
    
    'Generate the list file
    GenerateTXTReport = GenerateTXTReportPluginsListFile(strLinesFileNames)
End Function

Private Function GenerateTXTReportPluginsListFile(ByRef strPluginsFileNamesList As String) As String
    Dim i As Integer
    Dim TXTListContent As String       'The content of the html list file
    Dim PluginsFileNamesList() As String
    Dim PluginsFileNamesCount As Integer
    
    PluginsFileNamesList = Split(strPluginsFileNamesList, ";", , vbBinaryCompare)
    PluginsFileNamesCount = UBound(PluginsFileNamesList)
    
    'Set the progress bar to zero
    frmMain.SetProgress 0
    
    'Prepare the HTML beginning (HTML header)
    frmMain.SetProgress 1
    TXTListContent = application_name & " - TXT Report for " & Target & vbNewLine & vbNewLine & _
        "Software: " & application_name & " (" & application_website_url & ")" & vbNewLine & _
        "Found vulnerabilities: " & PluginsFileNamesCount & vbNewLine & _
        "Date of report generation: " & GetTodaysDate("/") & vbNewLine & vbNewLine
    
    For i = 0 To PluginsFileNamesCount
        'Increase the progress bar. The On Error Resume Next prevents senseless
        'values that could lead to a programm error.
        On Error Resume Next
        
        'Everytime select the new plugin and do the check until finish
        If LenB(PluginsFileNamesList(i)) Then
            frmMain.SetProgress (100 / PluginsFileNamesCount) * i
            Call ParseATKPlugin(ReadPluginFromFile(PluginsFileNamesList(i), application_plugin_directory))
            
            'Generate the txt file
            'Call GenerateTXTReportPluginEntry(i)
            
            strTempReportPluginData = strTempReportPluginData & vbNewLine & _
                GenerateTXTReportPluginEntry(True, "     ", CStr(i)) & vbNewLine
        
            'Add the HTML row in the list
            TXTListContent = TXTListContent & _
                i & ". " & plugin_name & " (" & plugin_id & "), " & _
                plugin_protocol & "/" & plugin_port & ", " & _
                bug_severity & vbNewLine
        End If
    Next i
    
    'Print the main plugin data and footer
    TXTListContent = TXTListContent & vbNewLine & vbNewLine & vbNewLine & strTempReportPluginData
    
    GenerateTXTReportPluginsListFile = TXTListContent
    
    'Set the progress bar to 100
    frmMain.SetProgress 100
End Function

Public Function GenerateTXTReportPluginEntry(ByRef bolWordWrap As Boolean, _
                                            ByRef strSpaces As String, _
                                            Optional ByRef strVulnerabilityPosition As String) As String
    
    Dim i As Integer
    Dim ReportVulnerabilityStructureArray() As String
    Dim ReportVulnerabilityStructureArrayCount As Integer
    Dim txtPluginContent As String
    
    ReportVulnerabilityStructureArray = Split(report_structure, vbNewLine)
    ReportVulnerabilityStructureArrayCount = UBound(ReportVulnerabilityStructureArray)
    
    'Create the TXT plugin file html header
    If LenB(strVulnerabilityPosition) Then
        txtPluginContent = strVulnerabilityPosition & ". " & plugin_name & vbNewLine & vbNewLine
    End If
    
    'Open and read the report template file
    For i = 0 To ReportVulnerabilityStructureArrayCount
        'write the selected item
        If ReportVulnerabilityStructureArray(i) = "plugin_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin ID", plugin_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_name" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin name", plugin_name, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_filename" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin filename", plugin_filename, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_filesize" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin filesize", plugin_filesize & " bytes", bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_family" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin family", plugin_family, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_name" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin created name", plugin_created_name, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_email" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin created email", plugin_created_email, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_web" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin created web", plugin_created_web, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_company" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin created company", plugin_created_company, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_date" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin created date", plugin_created_date, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_name" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin updated name", plugin_updated_name, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_email" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin updated email", plugin_updated_email, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_web" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin updated web", plugin_updated_web, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_company" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin updated company", plugin_updated_company, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_date" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin updated date", plugin_updated_date, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_version" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin version", plugin_version, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_changelog" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin changelog", plugin_changelog, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_protocol" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin protocol", plugin_protocol, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_port" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin port", plugin_port, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_procedure_detection" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin procedure detection", plugin_procedure_detection, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_procedure_exploit" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin procedure exploit", plugin_procedure_exploit, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_detection_accuracy" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin detection accuracy", plugin_detection_accuracy, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_exploit_accuracy" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin exploit accuracy", plugin_exploit_accuracy, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_comment" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Plugin comment", plugin_comment, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_name" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug published name", bug_published_name, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_email" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug published email", bug_published_email, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_web" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug published web", bug_published_web, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_company" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug published company", bug_published_company, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_date" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug published date", bug_published_date, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_advisory" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug advisory", bug_advisory, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_produced_name" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug produced name", bug_produced_name, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_produced_email" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug produced email", bug_produced_email, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_produced_web" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug produced web", bug_produced_web, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_affected" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug affected", bug_affected, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_not_affected" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug not affected", bug_not_affected, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_false_positives" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug false positives", bug_false_positives, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_false_negatives" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug false negatives", bug_false_negatives, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_vulnerability_class" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug vulnerability class", bug_vulnerability_class, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_description" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug description", bug_description, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_response" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug response", vbNewLine & vbNewLine & _
                "     " & Replace(Mid$(LoadResponseFromFile(application_response_directory & "\" & Target & "-" & plugin_filename & ".txt"), 1, 1024), vbCrLf, vbCrLf & "     ", , , vbBinaryCompare), True, "     ")
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_solution" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug solution", bug_solution, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_fixing_time" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug fixing time", bug_fixing_time, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_exploit_availability" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug exploit availability", bug_exploit_availability, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_exploit_url" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug exploit url", bug_exploit_url, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_remote" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug remote", bug_remote, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_local" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug local", bug_local, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_severity" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug severity", bug_severity, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_popularity" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug popularity", bug_popularity, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_simplicity" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug simplicity", bug_simplicity, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_impact" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug impact", bug_impact, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_risk" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug risk", bug_risk, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_nessus_risk" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug Nessus risk", bug_nessus_risk, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_iss_scanner_rating" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug ISS Scanner rating", bug_iss_scanner_rating, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_netrecon_rating" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug Symantec NetRecon rating", bug_netrecon_rating, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_check_tool" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Bug check tools", bug_check_tool, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "source_cve" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source CVE", source_cve, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_certvu_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source CERT Vulnerability Note ID", source_certvu_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_cert_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source CERT ID", source_cert_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_uscertta_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source US-CERT ID", source_uscertta_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_securityfocus_bid" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source SecurityFocus BID", source_securityfocus_bid, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_osvdb_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source OSVDB ID", source_osvdb_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_secunia_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Secunia ID", source_secunia_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_securiteam_url" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source SecuriTeam URL", source_securiteam_url, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_securitytracker_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Security Tracker ID", source_securitytracker_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_scip_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source scipID", source_scip_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_tecchannel_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source tecchannel ID", source_tecchannel_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_heise_news" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Heise News", source_heise_news, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_heise_security" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Heise Security", source_heise_security, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_aerasec_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source AeraSecID", source_aerasec_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_nessus_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Nessus ID", source_nessus_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_issxforce_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source ISS X-Force ID", source_issxforce_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_snort_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Snort ID", source_snort_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_arachnids_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source ArachnIDS ID", source_arachnids_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_mssb_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Microsoft Security Bulletin ID", source_mssb_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_mskb_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Microsoft Knowledge Base ID", source_mskb_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_netbsdsa_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source NetBSD Security Advisory ID", source_netbsdsa_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_rhsa_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source RedHat Security Advisory ID", source_rhsa_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_ciac_id" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source CIAC ID", source_ciac_id, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_literature" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Literature", source_literature, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_misc" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Source Misc.", source_misc, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "application_name" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Name", application_name, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_website_url" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Website URL", application_website_url, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_configuration_filename" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Configuration Filename", application_configuration_filename, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_log_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Log Enable", CStr(application_log_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_speech_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Speech Enable", CStr(application_speech_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_suggestion_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Suggestion Enable", CStr(application_suggestion_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_vulnerability_found_alert_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Vulnerability Found Alert Enable", CStr(application_vulnerability_found_alert_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_vulnerability_not_found_alert_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Vulnerability Not Found Alert Enable", CStr(application_vulnerability_not_found_alert_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_attack_mode" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Vulnerability Attack Mode", application_attack_mode, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_attack_timeout" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Attack Timeout", CStr(application_attack_timeout) & " ms", bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_sleep_time_default" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Sleep Time", CStr(application_sleep_time_default) & " ms", bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_icmp_mapping_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application ICMP Mapping Enable", CStr(application_icmp_mapping_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_no_dos_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application No DoS Enable", CStr(application_no_dos_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_silent_checks_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Silen Checks Enable", CStr(application_silent_checks_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_help_url" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Help URL", application_help_url, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_log_directory" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Log Directory", application_log_directory, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_log_security_level" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Log Security Level", CStr(application_log_security_level), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_directory" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Plugin Directory", application_plugin_directory, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_download_url" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Plugin Download URL", application_plugin_download_url, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_external_editor" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Plugin External Editor", application_plugin_external_editor, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_report_directory" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Report Directory", application_report_directory, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_report_open_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Report Open Enable", CStr(application_report_open_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_response_directory" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Response Directory", application_response_directory, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_suggestion_directory" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Suggestion Directory", application_suggestion_directory, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_searchengine_url" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Searchengine URL", application_searchengine_url, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_icmp_mapping_ignore_enable" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application ICMP Mapping Ignore Enable", CStr(application_icmp_mapping_ignore_enable), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_count" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Application Plugin Count", HowManyLoadedPlugins, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "report_structure" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Structure", report_structure, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_filename" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Filename", report_filename, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_filesize" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Filesize", report_filesize & " bytes", bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_filecontent" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Filecontent", report_filecontent, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filename" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Template Filename", report_template_filename, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filepath" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Template Filepath", report_template_filepath, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filesize" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Template Filesize", report_template_filesize & " bytes", bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filecontent" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Report Template Filecontent", report_template_filecontent, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "<br>" Then
            txtPluginContent = txtPluginContent & vbNewLine
        
        ElseIf ReportVulnerabilityStructureArray(i) = "session_procedure_type" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Session procedure type", session_procedure_type, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "session_procedure_commands" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Session procedure commands", session_procedure_commands, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_target" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Scan Target", Target, bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_date" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Scan Date", GetTodaysDate("/"), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_time" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Scan Time", GetActualTime(":"), bolWordWrap, strSpaces)
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_mode" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("Scan Mode", application_attack_mode, bolWordWrap, strSpaces)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "system_username" Then
            txtPluginContent = txtPluginContent & CreateTextTableRow("System Username", system_username, bolWordWrap, strSpaces)
        End If
    Next i
        
    'Write the data back to the function
'    strTempReportPluginData = strTempReportPluginData & vbNewLine & _
'        txtPluginContent & vbNewLine
    GenerateTXTReportPluginEntry = txtPluginContent
End Function

Private Function CreateTextTableRow(ByRef RowName As String, _
                                    ByRef VariantContent As String, _
                                    ByRef bolWordWrap As Boolean, _
                                    ByRef strSpaces As String) As String
    
    If LenB(VariantContent) Then
        If bolWordWrap = True Then
            CreateTextTableRow = LineWrap(strSpaces & RowName & ": " & VariantContent) & vbNewLine
        Else
            CreateTextTableRow = RowName & ": " & VariantContent & vbNewLine
        End If
    End If
End Function

Public Function LineWrap(sString As String) As String
    Dim lPos As Long
    Dim iPosCounter As Long
    Dim lFinalLen As Long
    Dim lBeginPos As Long
    Dim lEndPos As Long
    Dim iWordLen As Long
    Dim iWordPos As Long
    Dim dWrapThresh As Integer
    Dim iInterval As Integer
    lFinalLen = Len(sString)
    
    iInterval = 70

    Do Until lPos >= lFinalLen

        If iPosCounter = iInterval Then 'ok, we hit the wrap point
            iPosCounter = 0 'Reset the interval counter
            'Get the beginning position of the curre
            '     nt word

            For lBeginPos = lPos To 0 Step -1
                If Mid$(sString, lBeginPos, 1) = " " Then Exit For
            Next lBeginPos

            'Get the ending position of the current
            '     word

            For lEndPos = lPos To lFinalLen
                If Mid$(sString, lEndPos, 1) = " " Then Exit For
            Next lEndPos

            'Get the length of the current word
            iWordLen = (lEndPos - 1) - (lBeginPos + 1)
            'Find out at which character we are loca
            '     ted in the word
            iWordPos = lPos - (lBeginPos + 1)
            'If we are over half way, then we move f
            '     orward, otherwise we move back
            dWrapThresh = iWordLen / 2
            If lEndPos > Len(sString) Then Exit Do

            If iWordPos >= dWrapThresh Then 'Wrap at End of word
                sString = Left$(sString, lEndPos) + vbNewLine + "     " + Right$(sString, lFinalLen - lEndPos)
            Else 'Wrap at beginning of word
                sString = Left$(sString, lBeginPos) + vbNewLine + "     " + Right$(sString, lFinalLen - lBeginPos)
            End If

            lFinalLen = Len(sString)
        End If

        iPosCounter = iPosCounter + 1
        If lPos > 0 Then If Mid$(sString, lPos, 2) = vbNewLine Then iPosCounter = 0 'Reset if new line already
        lPos = lPos + 1
    Loop

    LineWrap = sString
End Function
