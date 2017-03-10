Attribute VB_Name = "modReportsHandlingHTMLExport"
Option Explicit

Dim strTempReportPluginData As String

Public Function GenerateHTMLReport(ByRef strPluginReportFileName As String, _
                                ByRef strReportDestinationFileName As String) As String
    
    Dim i As Integer
    Dim TempString As String
    Dim strPluginReportFileContent As String
    Dim strLinesArray() As String
    Dim strLinesArrayLineCount As Integer
    Dim strLinesFileNameArray() As String
    Dim strLinesFileNames As String
    Dim strTargetDirectoryName As String
    
    'Reset the old report data
    strTempReportPluginData = vbNullString

    'Load the report from the file
    strPluginReportFileContent = LoadReportFromFile(strPluginReportFileName)
    
    strLinesArray = Split(strPluginReportFileContent, vbNewLine, , vbBinaryCompare)
    
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
    
    'Prepare the report directory for the target
    strTargetDirectoryName = application_report_directory & "\" & Target
    
    'Create the report directory if it does not exist
    If Not (Dir$(strTargetDirectoryName, 16) <> "") Then
        MkDir (strTargetDirectoryName)
    End If
    
    'Generate the list file
    TempString = GenerateHTMLReportPluginsListFile(strLinesFileNames)

    'Write the HTMLListcontent to a HTML file. The file name can note be chosen at
    'this time. Such a feature should be added in a further release.
    On Error Resume Next ' Needed if there are no write permissions
    Open strReportDestinationFileName For Output As #1
        Print #1, TempString
    Close
    
    'Open the report after generation if wanted
    If application_report_open_enable = True Then
        Call ShellExecute(frmMain.hwnd, "Open", strReportDestinationFileName, "", App.Path, 1)
    End If
End Function

Private Function GenerateHTMLReportPluginsListFile(ByRef strPluginsFileNamesList As String) As String
    
    Dim i As Integer
    Dim HTMLListTitle As String         'The title of the document
    Dim HTMLListContent As String       'The content of the html list file
    Dim PluginsFileNamesList() As String
    Dim PluginsFileNamesCount As Integer
    
    PluginsFileNamesList = Split(strPluginsFileNamesList, ";", , vbBinaryCompare)
    PluginsFileNamesCount = UBound(PluginsFileNamesList)
    
    'Set the progress bar to zero
    frmMain.SetProgress 0
    
    'Define the title of the html document
    HTMLListTitle = application_name & " - HTML Report for " & Target
    
    'Prepare the HTML beginning (HTML header)
    frmMain.SetProgress 1
    HTMLListContent = "<html>" & vbNewLine & _
        "<head>" & vbNewLine & _
        "<meta name=Author content=""Marc Ruef"">" & vbNewLine & _
        "<meta name=Generator content=""" & application_name & """>" & vbNewLine & _
        "<meta name=Description content=""ATK HTML Report"">" & vbNewLine & _
        "<meta name=KeyWords content=""ATK, Attack Tool Kit, Plugins, checks, list, report, reporting, html, Marc Ruef"">" & vbNewLine & _
        "<title>" & HTMLListTitle & "</title>" & vbNewLine & _
        "</head>" & vbNewLine & _
        "<body>" & vbNewLine & _
        "<font face=Verdana size=-1><b>" & HTMLListTitle & "</b>" & vbNewLine & _
        "<p>Software: <a href=" & application_website_url & " target=_TOP>" & application_name & "</a>" & vbNewLine & _
        "<br>Found vulnerabilities: " & PluginsFileNamesCount & "" & vbNewLine & _
        "<br>Date of report generation: " & Date & "</font>" & vbNewLine & _
        "<br>&nbsp;" & vbNewLine
    
    'Prepare the HTML table
    frmMain.SetProgress 2
    HTMLListContent = HTMLListContent & _
        "<a name=vulnerabilities><table border cellspacing=0 width=100%>" & vbNewLine & _
        "<tr align=left valign=top>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Name</font></font></b></td>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Port</font></font></b></td>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Family</font></font></b></td>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Class</font></font></b></td>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Severity</font></font></b></td>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>ID</font></font></b></td>" & vbNewLine & _
        "</tr>" & vbNewLine
    
    For i = 0 To PluginsFileNamesCount
        'Increase the progress bar. The On Error Resume Next prevents senseless
        'values that could lead to a programm error.
        On Error Resume Next
        
        'Everytime select the new plugin and do the check until finish
        If LenB(PluginsFileNamesList(i)) Then
            frmMain.SetProgress (100 / PluginsFileNamesCount) * i
            Call ParseATKPlugin(ReadPluginFromFile(PluginsFileNamesList(i), application_plugin_directory))
            
            'Generate the html file
            Call GenerateHTMLReportPluginEntry
        
            'Add the HTML row in the list
            HTMLListContent = HTMLListContent & _
            "<tr align=left valign=top>" & vbNewLine & _
            "<td align=left valign=top title=""" & CutTooLongString(bug_description, 128) & """><font face=Verdana><font size=-1><a href=""#" & plugin_id & """>" & plugin_name & "</a></font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & plugin_protocol & "/" & plugin_port & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & plugin_family & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & bug_vulnerability_class & "</font></font></td>" & vbNewLine & _
            "<td bgcolor=""#" & GetSeverityHTMLColor(bug_severity) & """><font face=Verdana><font size=-1>" & bug_severity & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & plugin_id & "</font></font></td>" & vbNewLine & _
            "</tr>" & vbNewLine
        End If
    Next i
    
    'Print the main plugin data and footer
    HTMLListContent = HTMLListContent & _
        "</table>" & vbNewLine & _
        "<br><hr><br>" & vbNewLine & _
        strTempReportPluginData & _
        "<font face=Verdana><font size=-3>This file was generated by the <a href=" & application_website_url & " target=_TOP>Attack Tool Kit (ATK)</a>, the open-sourced security scanner and exploiting framework.</font></font>" & vbNewLine & _
        "</body>" & vbNewLine & _
        "</html>" & vbNewLine
    
    GenerateHTMLReportPluginsListFile = HTMLListContent
    
    'Set the progress bar to 100
    frmMain.SetProgress 100
End Function

Private Sub GenerateHTMLReportPluginEntry()
    Dim i As Integer
    Dim ReportVulnerabilityStructureArray() As String
    Dim ReportVulnerabilityStructureArrayCount As Integer
    Dim HTMLPluginContent As String
    
    ReportVulnerabilityStructureArray = Split(report_structure, vbNewLine)
    ReportVulnerabilityStructureArrayCount = UBound(ReportVulnerabilityStructureArray)
    
    'Create the HTML plugin file html header
    HTMLPluginContent = "<a name=""" & plugin_id & """><font face=Verdana><font size=-1><b>" & plugin_name & "</b></font></font><br><br>" & vbNewLine & _
        "<table border=0 width=100%>" & vbNewLine
    'Open and read the report template file
    For i = 0 To ReportVulnerabilityStructureArrayCount
        'write the selected item
        If ReportVulnerabilityStructureArray(i) = "plugin_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin ID", plugin_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_name" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin name", plugin_name)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_filename" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin filename", plugin_filename)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_filesize" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin filesize", plugin_filesize & " bytes")
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_family" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin family", plugin_family)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_name" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin created name", plugin_created_name)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_email" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin created email", plugin_created_email, _
                "mailto:" & plugin_created_name & " <" & plugin_created_email & ">?subject=" & plugin_filename & "&" & _
                "body=Dear " & plugin_created_name & "%0D%0A%0D%0A" & _
                "I would like to ask you something about the plugin '" & plugin_filename & "' (ATK plugin ID " & _
                plugin_id & ") you have written at " & plugin_created_date & " for the Attack Tool Kit Project[1]." & _
                "%0D%0A%0D%0A" & "Kind regards" & "%0D%0A%0D%0A" & "[1] " & application_website_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_web" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin created web", plugin_created_web, plugin_created_web)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_company" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin created company", plugin_created_company)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_created_date" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin created date", plugin_created_date)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_name" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin updated name", plugin_updated_name)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_email" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin updated email", plugin_updated_email, _
            "mailto:" & plugin_updated_name & " <" & plugin_updated_email & ">?subject=" & plugin_filename & " " & plugin_version & "&" & _
            "body=Dear " & plugin_updated_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the plugin '" & plugin_filename & " " & plugin_version & "' (ATK plugin ID " & plugin_id & ") you have updated at " & _
            plugin_updated_date & " for the Attack Tool Kit Project[1]." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & application_website_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_web" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin updated web", plugin_updated_web, plugin_updated_web)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_company" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin updated company", plugin_updated_company)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_updated_date" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin updated date", plugin_updated_date)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_version" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin version", plugin_version)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_changelog" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin changelog", plugin_changelog)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_protocol" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin protocol", plugin_protocol)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_port" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin port", plugin_port)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_procedure_detection" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin procedure detection", plugin_procedure_detection)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_procedure_exploit" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin procedure exploit", plugin_procedure_exploit)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_detection_accuracy" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin detection accuracy", plugin_detection_accuracy)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_exploit_accuracy" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin exploit accuracy", plugin_exploit_accuracy)
        ElseIf ReportVulnerabilityStructureArray(i) = "plugin_comment" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Plugin comment", plugin_comment)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_name" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug published name", bug_published_name)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_email" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug published email", bug_published_email, "mailto:" & bug_published_name & " <" & bug_published_email & ">?subject=" & plugin_name & "&" & _
            "body=Dear " & bug_published_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the vulnerability '" & plugin_name & "'[1] that can also be tested/exploitet since " & plugin_created_date & " with the plugin " & plugin_id & _
            " of the Attack Tool Kit Project[2]." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & bug_advisory & "%0D%0A" & _
            "[2] " & application_website_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_web" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug published web", bug_published_web, bug_published_web)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_company" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug published company", bug_published_company)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_published_date" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug published date", bug_published_date)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_advisory" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug advisory", bug_advisory, bug_advisory)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_produced_name" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug produced name", bug_produced_name)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_produced_email" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug produced email", bug_produced_email, "mailto:" & bug_produced_name & " <" & bug_produced_email & ">?subject=" & plugin_name & "&" & _
            "body=Dear " & bug_produced_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the vulnerability '" & plugin_name & "'[1] that is affecting " & bug_affected & "." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & bug_advisory)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_produced_web" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug produced web", bug_produced_web, bug_produced_web)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_affected" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug affected", bug_affected)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_not_affected" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug not affected", bug_not_affected)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_false_positives" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug false positives", bug_false_positives)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_false_negatives" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug false negatives", bug_false_negatives)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_vulnerability_class" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug vulnerability class", bug_vulnerability_class)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_description" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug description", bug_description)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_response" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug response", Replace(Mid$(LoadResponseFromFile(application_response_directory & "\" & Target & "-" & plugin_filename & ".txt"), 1, 1024), vbCrLf, "<br>", , , vbBinaryCompare))
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_solution" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug solution", bug_solution)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_fixing_time" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug fixing time", bug_fixing_time)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_exploit_availability" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug exploit availability", bug_exploit_availability)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_exploit_url" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug exploit url", bug_exploit_url, bug_exploit_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_remote" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug remote", bug_remote)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_local" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug local", bug_local)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_severity" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug severity", bug_severity)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_popularity" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug popularity", bug_popularity)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_simplicity" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug simplicity", bug_simplicity)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_impact" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug impact", bug_impact)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_risk" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug risk", bug_risk)
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_nessus_risk" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug Nessus risk", bug_nessus_risk, "http://www.nessus.org")
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_iss_scanner_rating" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug ISS Scanner rating", bug_iss_scanner_rating, "http://www.iss.net")
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_netrecon_rating" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug Symantec NetRecon rating", bug_netrecon_rating, "http://www.symantec.com")
        ElseIf ReportVulnerabilityStructureArray(i) = "bug_check_tool" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Bug check tools", bug_check_tool, application_searchengine_url & bug_check_tool)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "source_cve" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source CVE", source_cve, "http://cve.mitre.org/cgi-bin/cvename.cgi?name=" & source_cve)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_certvu_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source CERT Vulnerability Note ID", source_certvu_id, "http://www.kb.cert.org/vuls/id/" & source_certvu_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_cert_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source CERT ID", source_cert_id, "http://www.cert.org/advisories/" & source_cert_id & ".html")
        ElseIf ReportVulnerabilityStructureArray(i) = "source_uscertta_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source US-CERT ID", source_uscertta_id, "http://www.us-cert.gov/cas/techalerts/" & source_uscertta_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_securityfocus_bid" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source SecurityFocus BID", source_securityfocus_bid, "http://www.securityfocus.com/bid/" & source_securityfocus_bid)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_osvdb_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source OSVDB ID", source_osvdb_id, "http://www.osvdb.org/" & source_osvdb_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_secunia_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Secunia ID", source_secunia_id, "http://www.secunia.com/advisories/" & source_secunia_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_securiteam_url" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source SecuriTeam URL", source_securiteam_url, source_securiteam_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_securitytracker_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Security Tracker ID", source_securitytracker_id, "http://www.securitytracker.com/id?" & source_securitytracker_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_scip_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source scipID", source_scip_id, "http://www.scip.ch/cgi-bin/smss/showadvf.pl?id=" & source_scip_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_tecchannel_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source tecchannel ID", source_tecchannel_id, "http://www.tecchannel.de/sicherheit/reports/" & source_tecchannel_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_heise_news" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Heise News", source_heise_news, "http://www.heise.de/newsticker/data/" & source_heise_news)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_heise_security" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Heise Security", source_heise_security, "http://www.heise.de/security/news/meldung/" & source_heise_security)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_aerasec_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source AeraSecID", source_aerasec_id, "http://www.aerasec.de/security/index.html?id=" & source_aerasec_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_nessus_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Nessus ID", source_nessus_id, "http://www.nessus.org/plugins/index.php?view=single&id=" & source_nessus_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_issxforce_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source ISS X-Force ID", source_issxforce_id, "http://xforce.iss.net/xforce/alerts/id/" & source_issxforce_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_snort_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Snort ID", source_snort_id, "http://www.snort.org/snort-db/sid.html?sid=" & source_snort_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_arachnids_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source ArachnIDS ID", source_arachnids_id, "http://www.whitehats.com/info/" & source_arachnids_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_mssb_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Microsoft Security Bulletin ID", source_mssb_id, "http://www.microsoft.com/technet/security/Bulletin/" & source_mssb_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_mskb_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Microsoft Knowledge Base ID", source_mskb_id, "http://support.microsoft.com/default.aspx?scid=kb;en-us;" & source_mskb_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_netbsdsa_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source NetBSD Security Advisory ID", source_netbsdsa_id, "ftp://ftp.netbsd.org/pub/NetBSD/security/advisories/" & source_netbsdsa_id & ".txt.asc")
        ElseIf ReportVulnerabilityStructureArray(i) = "source_rhsa_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source RedHat Security Advisory ID", source_rhsa_id, "https://www.redhat.com/security/" & source_rhsa_id)
        ElseIf ReportVulnerabilityStructureArray(i) = "source_ciac_id" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source CIAC ID", source_ciac_id, "http://www.ciac.org")
        ElseIf ReportVulnerabilityStructureArray(i) = "source_literature" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Literature", source_literature, "http://www.amazon.com/exec/obidos/tg/detail/-/" & GetISBNFromString(source_literature))
        ElseIf ReportVulnerabilityStructureArray(i) = "source_misc" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Source Misc.", source_misc, source_misc)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "application_attack_mode" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Attack Mode", application_attack_mode)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_attack_timeout" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Attack Timeout", CStr(application_attack_timeout) & " ms")
        ElseIf ReportVulnerabilityStructureArray(i) = "application_configuration_filename" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Configuration Filename", application_configuration_filename, "file://" & application_configuration_filename)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_help_url" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Help URL", application_help_url, application_help_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_icmp_mapping_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application ICMP Mapping Enable", CStr(application_icmp_mapping_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_icmp_mapping_ignore_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application ICMP Mapping Ignore Enable", CStr(application_icmp_mapping_ignore_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_log_directory" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Log Directory", application_log_directory, "file://" & application_log_directory)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_log_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Log Enable", CStr(application_log_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_log_security_level" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Log Security Level", CStr(application_log_security_level))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_name" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Name", application_name)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_no_dos_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application No DoS Enable", CStr(application_no_dos_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_count" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Plugin Count", HowManyLoadedPlugins)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_directory" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Plugin Directory", application_plugin_directory, "file://" & application_plugin_directory)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_download_url" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Plugin Download URL", application_plugin_download_url, application_plugin_download_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_plugin_external_editor" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Plugin External Editor", application_plugin_external_editor)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_report_directory" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Report Directory", application_report_directory, "file://" & application_report_directory)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_report_open_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Report Open Enable", CStr(application_report_open_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_response_directory" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Response Directory", application_response_directory, "file://" & application_response_directory)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_searchengine_url" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Searchengine URL", application_searchengine_url, application_searchengine_url)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_silent_checks_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Silent Checks Enable", CStr(application_silent_checks_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_sleep_time_default" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Sleep Time Default", CStr(application_sleep_time_default) & " ms")
        ElseIf ReportVulnerabilityStructureArray(i) = "application_speech_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Speech Enable", CStr(application_speech_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_suggestion_directory" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Suggestion Directory", application_suggestion_directory, "file://" & application_suggestion_directory)
        ElseIf ReportVulnerabilityStructureArray(i) = "application_suggestion_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Suggestion Enable", CStr(application_suggestion_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_vulnerability_found_alert_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Vulnerability Found Alert Enable", CStr(application_vulnerability_found_alert_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_vulnerability_not_found_alert_enable" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Vulnerability Not Found Alert Enable", CStr(application_vulnerability_not_found_alert_enable))
        ElseIf ReportVulnerabilityStructureArray(i) = "application_website_url" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Application Website URL", application_website_url, application_website_url)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "report_filecontent" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Filecontent", report_filecontent)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_filename" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Filename", report_filename)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_filesize" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Filesize", report_filesize & " bytes")
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filename" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Template Filename", report_template_filename)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filepath" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Template Filepath", report_template_filepath)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filesize" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Template Filesize", report_template_filesize & " bytes")
        ElseIf ReportVulnerabilityStructureArray(i) = "report_template_filecontent" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Template Filecontent", report_template_filecontent)
        ElseIf ReportVulnerabilityStructureArray(i) = "report_structure" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Report Structure", report_structure)

        ElseIf ReportVulnerabilityStructureArray(i) = "system_username" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("System Username", system_username)

        ElseIf ReportVulnerabilityStructureArray(i) = "scan_target" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Scan Target", Target, "http://" & Target & ":" & plugin_port)
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_date" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Scan Date", GetTodaysDate("/"))
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_time" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Scan Time", GetActualTime(":"))
        ElseIf ReportVulnerabilityStructureArray(i) = "scan_mode" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Scan Mode", application_attack_mode)
        
        ElseIf ReportVulnerabilityStructureArray(i) = "session_procedure_type" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Session procedure type", session_procedure_type)
        ElseIf ReportVulnerabilityStructureArray(i) = "session_procedure_commands" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow("Session procedure commands", session_procedure_commands)
    
        ElseIf ReportVulnerabilityStructureArray(i) = "<br>" Then
            HTMLPluginContent = HTMLPluginContent & CreateHTMLTableRow(" ", " ")
        End If
    Next i
        
    'Write the data back to the function
    strTempReportPluginData = strTempReportPluginData & vbNewLine & _
        HTMLPluginContent & vbNewLine & _
        "</table><br>" & vbNewLine & _
        "<font face=Verdana><font size=-2><div align=right><a href=#vulnerabilities>back to the list of vulnerabilities</a></div></font></font><hr><br>" & vbNewLine
End Sub
