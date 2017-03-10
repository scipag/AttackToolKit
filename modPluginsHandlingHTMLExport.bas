Attribute VB_Name = "modPluginsHandlingHTMLExport"
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2004-12-27                                                           *
' * - Added a html sanitizing routine to prevent html errors (e.g. xss attacks).     *
' * Version 4.0 2004-12-05                                                           *
' * - Made the first preparations for the new reporting functionality in 4.0.        *
' * Version 3.0 2004-11-13                                                           *
' * - Changed the misc source link to not using the search engine.                   *
' * Version 3.0 2004-11-07                                                           *
' * - Shortened the title information for the plugin comments.                       *
' * Version 3.0 2004-11-03                                                           *
' * - Added the html title tag for description and comments in the main html file.   *
' * - Fixed the bug with the missing space in the title link.                        *
' * Version 2.1 2004-09-09                                                           *
' * - Improved the mailto tags to provide mail templates.                            *
' * Version 2.1 2004-09-08                                                           *
' * - Corrected a dedicated and small bug during export for URLs.                    *
' * - Added URL linking for bug_published_web and bug_published_email.               *
' * - Optimized the OSVDB.org linking (using short URLs now).                        *
' ************************************************************************************

Public Function CreateHTMLTableRow(ByRef RowName As String, _
                                    ByRef VariantContent As String, _
                                    Optional ByRef LinkURL As String) As String

    'Check if there is a content in the variant. If yes, create a row
    If LenB(VariantContent) Then
        'Check if a href link is needed
        If LenB(LinkURL) = 0 Then
            CreateHTMLTableRow = "<tr align=left valign=top><td width=200 nowrap><font face=Verdana size=-1>" & RowName & _
            "</font></td><td><font face=Verdana size=-1>" & _
            StringToHTMLEncoding(VariantContent) & "</font></td></tr>" & vbNewLine
        Else
            CreateHTMLTableRow = "<tr align=left valign=top><td width=200 nowrap><font face=Verdana size=-1>" & RowName & _
            "</font></td><td><font face=Verdana size=-1><a href=""" & LinkURL & """ target=_TOP>" & _
            StringToHTMLEncoding(VariantContent) & "</a></font></td></tr>" & vbNewLine
        End If
    End If
End Function

Public Sub ExportPluginsToHTMLFile()
    Dim i As Integer                    'This i is used for the counter
    Dim PluginsListFileNames As String
    Dim AvailablePlugins As Integer
    Dim PluginHTMListContent As String

    'Write log entry
    WriteLogEntry "Exporting loaded plugins list to html file ...", 6

    'Count the available plugins
    AvailablePlugins = frmMain.filATKPlugins.ListCount - 1
    
    'Prepare the plugin file names for list generation
    For i = 0 To AvailablePlugins
        frmMain.filATKPlugins.ListIndex = i
        PluginsListFileNames = PluginsListFileNames & ";" & frmMain.filATKPlugins.Filename
    Next i
    
    'Generate the list file
    PluginHTMListContent = GeneratePluginsListHTMLFile(PluginsListFileNames, application_plugin_directory)

    'Write the HTMLListcontent to a HTML file. The file name can note be chosen at
    'this time. Such a feature should be added in a further release.
    On Error Resume Next ' Needed if there are no write permissions
    Open application_plugin_directory & "\pluginslist.html" For Output As #1
        Print #1, PluginHTMListContent
    Close
    
    'Open the exported file in the default web browser.
    Call ShellExecute(frmMain.hwnd, "Open", application_plugin_directory & "\pluginslist.html", "", App.Path, 1)
End Sub

Public Function GeneratePluginsListHTMLFile(ByRef strPluginsFileNamesList As String, _
                                        ByRef strTargetDirectoryName As String) As String
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
    HTMLListTitle = application_name & " - Exported list of loaded plugins " & GetTodaysDate("/")
    
    'Prepare the HTML beginning (HTML header)
    frmMain.SetProgress 1
    HTMLListContent = "<html>" & vbNewLine & _
        "<head>" & vbNewLine & _
        "<meta name=Author content=""Marc Ruef"">" & vbNewLine & _
        "<meta name=Generator content=""" & application_name & """>" & vbNewLine & _
        "<meta name=Description content=""Exported list of ATK plugins"">" & vbNewLine & _
        "<meta name=KeyWords content=""ATK, Attack Tool Kit, Plugins, checks, list, Marc Ruef"">" & vbNewLine & _
        "<title>" & HTMLListTitle & "</title>" & vbNewLine & _
        "</head>" & vbNewLine & _
        "<body>" & vbNewLine & _
        "<font face=Verdana size=-1><b>" & HTMLListTitle & "</b>" & vbNewLine & _
        "<p>Software: <a href=" & application_website_url & " target=_TOP>" & application_name & "</a>" & vbNewLine & _
        "<br>Loaded Plugins: " & PluginsFileNamesCount & "" & vbNewLine & _
        "<br>Date of export: " & Date & "</font>" & vbNewLine & _
        "<br>&nbsp;" & vbNewLine
    
    'Prepare the HTML table
    frmMain.SetProgress 2
    HTMLListContent = HTMLListContent & _
        "<table border cellspacing=0 width=100%>" & vbNewLine & _
        "<tr align=left valign=top>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Name</font></font></b></td>" & vbNewLine & _
        "<td><b><font face=Verdana><font size=-1>Version</font></font></b></td>" & vbNewLine & _
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
            Call WritePluginHTMLFileToFile(strTargetDirectoryName, plugin_filename & ".html")
        
            'Add the HTML row in the list
            HTMLListContent = HTMLListContent & _
            "<tr align=left valign=top>" & vbNewLine & _
            "<td align=left valign=top title=""" & CutTooLongString(bug_description, 92) & """><font face=Verdana><font size=-1><a href=""" & plugin_filename & ".html"">" & plugin_name & "</a></font></font></td>" & vbNewLine & _
            "<td title=""" & CutTooLongString(GetLatestChange, 92) & """><font face=Verdana><font size=-1>" & plugin_version & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & plugin_protocol & "/" & plugin_port & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & plugin_family & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & bug_vulnerability_class & "</font></font></td>" & vbNewLine & _
            "<td bgcolor=""#" & GetSeverityHTMLColor(bug_severity) & """><font face=Verdana><font size=-1>" & bug_severity & "</font></font></td>" & vbNewLine & _
            "<td><font face=Verdana><font size=-1>" & plugin_id & "</font></font></td>" & vbNewLine & _
            "</tr>" & vbNewLine
        End If
    Next i
    
    'Print the footer
    HTMLListContent = HTMLListContent & _
        "</table>" & vbNewLine & _
        "<font face=Verdana><font size=-3><br>This file was generated by the <a href=" & application_website_url & " target=_TOP>Attack Tool Kit (ATK)</a>, the open-sourced security scanner and exploiting framework.</font></font>" & vbNewLine & _
        "</body>" & vbNewLine & _
        "</html>" & vbNewLine
    
    GeneratePluginsListHTMLFile = HTMLListContent
    
    'Set the progress bar to 100
    frmMain.SetProgress 100
End Function

Public Function GeneratePluginHTMLFile() As String
    Dim HTMLPluginContent As String
    
    'Create the HTML plugin file html header
    HTMLPluginContent = "<html>" & vbNewLine & _
        "<head>" & vbNewLine & _
        "<meta name=Author content=""Marc Ruef"">" & vbNewLine & _
        "<meta name=Generator content=""" & application_name & """>" & vbNewLine & _
        "<meta name=Description content=""Exported plugin of ATK"">" & vbNewLine & _
        "<meta name=KeyWords content=""ATK, Attack Tool Kit, Plugins, checks, list"">" & vbNewLine & _
        "<title>" & application_name & " - " & plugin_filename & " " & plugin_version & "</title>" & vbNewLine & _
        "</head>" & vbNewLine & _
        "<body>" & vbNewLine & _
        "<font face=Verdana><font size=-1><b>" & plugin_name & " " & plugin_version & "</b></font></font>" & vbNewLine & _
        "<br>&nbsp;" & vbNewLine & _
        "<center><table border=0 width=100%>" & vbNewLine
    
    'Add the HTML plugin file plugin data part 1
    HTMLPluginContent = HTMLPluginContent & _
        CreateHTMLTableRow("Plugin ID", plugin_id) & _
        CreateHTMLTableRow("Plugin name", plugin_name) & _
        CreateHTMLTableRow("Plugin filename", plugin_filename, plugin_filename) & _
        CreateHTMLTableRow("Plugin filesize", plugin_filesize & " bytes") & _
        CreateHTMLTableRow("Plugin family", plugin_family) & _
        CreateHTMLTableRow("Plugin created name", plugin_created_name) & _
        CreateHTMLTableRow("Plugin created email", plugin_created_email, _
            "mailto:" & plugin_created_name & " <" & plugin_created_email & ">?subject=" & plugin_filename & "&" & _
            "body=Dear " & plugin_created_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the plugin '" & plugin_filename & "' (ATK plugin ID " & plugin_id & ") you have written at " & _
            plugin_created_date & " for the Attack Tool Kit Project[1]." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & application_website_url) & _
        CreateHTMLTableRow("Plugin created web", plugin_created_web, plugin_created_web) & _
        CreateHTMLTableRow("Plugin created company", plugin_created_company) & _
        CreateHTMLTableRow("Plugin created date", plugin_created_date)
    
    'Add the HTML plugin file plugin data part 2
    HTMLPluginContent = HTMLPluginContent & _
        CreateHTMLTableRow("Plugin updated name", plugin_updated_name) & _
        CreateHTMLTableRow("Plugin updated email", plugin_updated_email, _
            "mailto:" & plugin_updated_name & " <" & plugin_updated_email & ">?subject=" & plugin_filename & " " & plugin_version & "&" & _
            "body=Dear " & plugin_updated_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the plugin '" & plugin_filename & " " & plugin_version & "' (ATK plugin ID " & plugin_id & ") you have updated at " & _
            plugin_updated_date & " for the Attack Tool Kit Project[1]." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & application_website_url) & _
        CreateHTMLTableRow("Plugin updated web", plugin_updated_web, plugin_updated_web) & _
        CreateHTMLTableRow("Plugin updated company", plugin_updated_company) & _
        CreateHTMLTableRow("Plugin updated date", plugin_updated_date) & _
        CreateHTMLTableRow("Plugin version", plugin_version) & _
        CreateHTMLTableRow("Plugin changelog", plugin_changelog) & _
        CreateHTMLTableRow("Plugin protocol", plugin_protocol) & _
        CreateHTMLTableRow("Plugin port", plugin_port) & _
        CreateHTMLTableRow("Plugin procedure detection", plugin_procedure_detection) & _
        CreateHTMLTableRow("Plugin procedure exploit", plugin_procedure_exploit) & _
        CreateHTMLTableRow("Plugin detection accuracy", plugin_detection_accuracy) & _
        CreateHTMLTableRow("Plugin exploit accuracy", plugin_exploit_accuracy) & _
        CreateHTMLTableRow("Plugin comment", plugin_comment)
        
    'Add the HTML plugin file bug data part 1
    HTMLPluginContent = HTMLPluginContent & _
        CreateHTMLTableRow("Bug published name", bug_published_name) & _
        CreateHTMLTableRow("Bug published email", bug_published_email, "mailto:" & bug_published_name & " <" & bug_published_email & ">?subject=" & plugin_name & "&" & _
            "body=Dear " & bug_published_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the vulnerability '" & plugin_name & "'[1] that can also be tested/exploitet since " & plugin_created_date & " with the plugin " & plugin_id & _
            " of the Attack Tool Kit Project[2]." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & bug_advisory & "%0D%0A" & _
            "[2] " & application_website_url) & _
        CreateHTMLTableRow("Bug published web", bug_published_web, bug_published_web) & _
        CreateHTMLTableRow("Bug published company", bug_published_company) & _
        CreateHTMLTableRow("Bug published date", bug_published_date) & _
        CreateHTMLTableRow("Bug advisory", bug_advisory, bug_advisory) & _
        CreateHTMLTableRow("Bug produced name", bug_produced_name) & _
        CreateHTMLTableRow("Bug produced email", bug_produced_email, "mailto:" & bug_produced_name & " <" & bug_produced_email & ">?subject=" & plugin_name & "&" & _
            "body=Dear " & bug_produced_name & "%0D%0A%0D%0A" & _
            "I would like to ask you something about the vulnerability '" & plugin_name & "'[1] that is affecting " & bug_affected & "." & "%0D%0A%0D%0A" & _
            "Kind regards" & "%0D%0A%0D%0A" & _
            "[1] " & bug_advisory) & _
        CreateHTMLTableRow("Bug produced web", bug_produced_web, bug_produced_web) & _
        CreateHTMLTableRow("Bug affected", bug_affected) & _
        CreateHTMLTableRow("Bug not affected", bug_not_affected) & _
        CreateHTMLTableRow("Bug vulnerability class", bug_vulnerability_class) & _
        CreateHTMLTableRow("Bug false positives", bug_false_positives) & _
        CreateHTMLTableRow("Bug false negatives", bug_false_negatives)
        
    'Add the HTML plugin file bug data part 2
    HTMLPluginContent = HTMLPluginContent & _
        CreateHTMLTableRow("Bug description", bug_description) & _
        CreateHTMLTableRow("Bug solution", bug_solution) & _
        CreateHTMLTableRow("Bug fixing time", bug_fixing_time) & _
        CreateHTMLTableRow("Bug exploit availability", bug_exploit_availability) & _
        CreateHTMLTableRow("Bug exploit url", bug_exploit_url, bug_exploit_url) & _
        CreateHTMLTableRow("Bug remote", bug_remote) & _
        CreateHTMLTableRow("Bug local", bug_local) & _
        CreateHTMLTableRow("Bug severity", bug_severity) & _
        CreateHTMLTableRow("Bug popularity", bug_popularity) & _
        CreateHTMLTableRow("Bug simplicity", bug_simplicity) & _
        CreateHTMLTableRow("Bug impact", bug_impact) & _
        CreateHTMLTableRow("Bug risk", bug_risk) & _
        CreateHTMLTableRow("Bug Nessus risk", bug_nessus_risk, "http://www.nessus.org") & _
        CreateHTMLTableRow("Bug ISS Scanner rating", bug_iss_scanner_rating, "http://www.iss.net") & _
        CreateHTMLTableRow("Bug Symantec NetRecon rating", bug_netrecon_rating, "http://www.symantec.com") & _
        CreateHTMLTableRow("Bug check tools", bug_check_tool, application_searchengine_url & bug_check_tool)
    
    'Add the HTML plugin file source data part 1
    HTMLPluginContent = HTMLPluginContent & _
        CreateHTMLTableRow("Source CVE", source_cve, "http://cve.mitre.org/cgi-bin/cvename.cgi?name=" & source_cve) & _
        CreateHTMLTableRow("Source CERT Vulnerability Note ID", source_certvu_id, "http://www.kb.cert.org/vuls/id/" & source_certvu_id) & _
        CreateHTMLTableRow("Source CERT ID", source_cert_id, "http://www.cert.org/advisories/" & source_cert_id & ".html") & _
        CreateHTMLTableRow("Source US-CERT ID", source_uscertta_id, "http://www.us-cert.gov/cas/techalerts/" & source_uscertta_id) & _
        CreateHTMLTableRow("Source SecurityFocus BID", source_securityfocus_bid, "http://www.securityfocus.com/bid/" & source_securityfocus_bid) & _
        CreateHTMLTableRow("Source OSVDB ID", source_osvdb_id, "http://www.osvdb.org/" & source_osvdb_id) & _
        CreateHTMLTableRow("Source Secunia ID", source_secunia_id, "http://www.secunia.com/advisories/" & source_secunia_id) & _
        CreateHTMLTableRow("Source SecuriTeam URL", source_securiteam_url, source_securiteam_url) & _
        CreateHTMLTableRow("Source Security Tracker ID", source_securitytracker_id, "http://www.securitytracker.com/id?" & source_securitytracker_id) & _
        CreateHTMLTableRow("Source scipID", source_scip_id, "http://www.scip.ch/cgi-bin/smss/showadvf.pl?id=" & source_scip_id) & _
        CreateHTMLTableRow("Source tecchannel ID", source_tecchannel_id, "http://www.tecchannel.de/sicherheit/reports/" & source_tecchannel_id) & _
        CreateHTMLTableRow("Source Heise News", source_heise_news, "http://www.heise.de/newsticker/data/" & source_heise_news) & _
        CreateHTMLTableRow("Source Heise Security", source_heise_security, "http://www.heise.de/security/news/meldung/" & source_heise_security) & _
        CreateHTMLTableRow("Source AeraSecID", source_aerasec_id, "http://www.aerasec.de/security/index.html?id=" & source_aerasec_id)
    
    'Add the HTML plugin file source data part 2
    HTMLPluginContent = HTMLPluginContent & _
        CreateHTMLTableRow("Source Nessus ID", source_nessus_id, "http://www.nessus.org/plugins/index.php?view=single&id=" & source_nessus_id) & _
        CreateHTMLTableRow("Source ISS X-Force ID", source_issxforce_id, "http://xforce.iss.net/xforce/alerts/id/" & source_issxforce_id) & _
        CreateHTMLTableRow("Source Snort ID", source_snort_id, "http://www.snort.org/snort-db/sid.html?sid=" & source_snort_id) & _
        CreateHTMLTableRow("Source ArachnIDS ID", source_arachnids_id, "http://www.whitehats.com/info/" & source_arachnids_id) & _
        CreateHTMLTableRow("Source Microsoft Security Bulletin ID", source_mssb_id, "http://www.microsoft.com/technet/security/Bulletin/" & source_mssb_id) & _
        CreateHTMLTableRow("Source Microsoft Knowledge Base ID", source_mskb_id, "http://support.microsoft.com/default.aspx?scid=kb;en-us;" & source_mskb_id) & _
        CreateHTMLTableRow("Source NetBSD Security Advisory ID", source_netbsdsa_id, "ftp://ftp.netbsd.org/pub/NetBSD/security/advisories/" & source_netbsdsa_id & ".txt.asc") & _
        CreateHTMLTableRow("Source RedHat Security Advisory ID", source_rhsa_id, "https://www.redhat.com/security/" & source_rhsa_id) & _
        CreateHTMLTableRow("Source CIAC ID", source_ciac_id, "http://www.ciac.org") & _
        CreateHTMLTableRow("Source Literature", source_literature, "http://www.amazon.com/exec/obidos/tg/detail/-/" & GetISBNFromString(source_literature)) & _
        CreateHTMLTableRow("Source Misc.", source_misc, source_misc)
        
    'Print the footer
    HTMLPluginContent = HTMLPluginContent & _
        "</table>" & vbNewLine & _
        "<div align=left><font face=Verdana><font size=-3><br>This file was generated by the <a href=" & application_website_url & " target=_TOP>Attack Tool Kit (ATK)</a>, the open-sourced security scanner and exploiting framework.</font></font></div>" & vbNewLine & _
        "</body>" & vbNewLine & _
        "</html>" & vbNewLine
    
    'Write the data back to the function
    GeneratePluginHTMLFile = HTMLPluginContent
End Function

Public Sub WritePluginHTMLFileToFile(ByRef strTargetDirectoryName As String, ByRef strTargetFileName As String)
    If Not (Dir$(strTargetDirectoryName, 16) <> "") Then
        MkDir (strTargetDirectoryName)
    End If
    
    'Write the HTMLPluginContent to a HTML file.
    On Error Resume Next ' Needed if there are no write permissions
    Open strTargetDirectoryName & "\" & strTargetFileName For Output As #1
        Print #1, GeneratePluginHTMLFile
    Close
End Sub

Public Function GetLatestChange() As String
    Dim strCommentsArray() As String
    
    If InStrB(1, plugin_changelog, ". ", vbBinaryCompare) Then
        
        strCommentsArray = Split(plugin_changelog, ". ", , vbBinaryCompare)
        
        GetLatestChange = strCommentsArray(UBound(strCommentsArray))
    Else
        GetLatestChange = plugin_changelog
    End If
End Function

Public Function CutTooLongString(ByRef strString As String, ByVal intLength As Integer) As String
    If Len(strString) > intLength Then
        CutTooLongString = Mid$(strString, 1, intLength) & "..."
    Else
        CutTooLongString = strString
    End If
End Function

Public Function GetSeverityHTMLColor(ByVal strSeverity As String) As String
    strSeverity = LCase$(strSeverity)
    
    If LenB(strSeverity) Then
        If InStrB(1, strSeverity, "high", vbBinaryCompare) Then
            GetSeverityHTMLColor = "FFF5F5"
        ElseIf InStrB(1, strSeverity, "critical", vbBinaryCompare) Then
            GetSeverityHTMLColor = "FFF5F5"
        ElseIf InStrB(1, strSeverity, "medium", vbBinaryCompare) Then
            GetSeverityHTMLColor = "FFFFF5"
        ElseIf InStrB(1, strSeverity, "low", vbBinaryCompare) Then
            GetSeverityHTMLColor = "F5FFF5"
        Else
            GetSeverityHTMLColor = "FFFFFF"
        End If
    End If
End Function

Public Function StringToHTMLEncoding(ByVal strString As String) As String
    'Dev note: Keep the order. It is important to replace the <br> at the end!
    
    '"
    If InStrB(1, strString, """", vbBinaryCompare) Then
        strString = Replace$(strString, """", "&quot;", , , vbTextCompare)
    End If
    
    '<
    If InStrB(1, strString, "<", vbBinaryCompare) Then
        strString = Replace$(strString, "<", "&lt;", , , vbTextCompare)
    End If
    
    '>
    If InStrB(1, strString, ">", vbBinaryCompare) Then
        strString = Replace$(strString, ">", "&gt;", , , vbTextCompare)
    End If
    
    'Afterwards re-replace the break tags. This is needed to keep the pre-formatting
    'of the http responses.
    If InStrB(1, strString, "&lt;br&gt;", vbBinaryCompare) Then
        strString = Replace$(strString, "&lt;br&gt;", "<br>", , , vbTextCompare)
    End If
    
    StringToHTMLEncoding = strString
End Function
