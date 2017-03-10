Attribute VB_Name = "modPluginsHandlingATKExport"
Option Explicit

Public Function GeneratePluginLine(ByRef strVariantName As String, ByRef strVariantContent As String) As String
    GeneratePluginLine = "<" & strVariantName & ">" & strVariantContent & "</" & strVariantName & ">" & vbNewLine
End Function

Public Sub WritePluginToFile(ByRef Filename As String)
    Dim PluginContent As String
    Dim PluginContentArray() As String
    Dim PluginContentItemCount As Integer
    Dim i As Integer
    
    'Prepare the comment
    If LenB(plugin_comment) = 0 Then
        plugin_comment = "This plugin was written with the ATK Attack Editor."
    End If
    
    'Prepare the exploit URL if a SecurityFocus exploit may be given
    If LenB(bug_exploit_url) = 0 Then
        If LenB(source_securityfocus_bid) Then
            bug_exploit_url = "http://www.securityfocus.com/bid/" & source_securityfocus_bid & "/exploit/"
        End If
    End If
    
    'Prepare the misc source
    If LenB(source_misc) = 0 Then
        source_misc = "http://www.computec.ch"
    End If
    
    'Prepare the literature
    If LenB(source_literature) = 0 Then
        source_literature = "Hacking Intern - Angriffe, Strategien, Abwehr, " & _
        "Marc Ruef, Marko Rogge, Uwe Velten and Wolfram Gieseke, " & _
        "November 1, 2002, Data Becker, Düsseldorf, ISBN 381582284X"
    End If
    
    'Collect the whole data
    PluginContent = GeneratePluginLine("plugin_id", plugin_id)
    PluginContent = PluginContent & GeneratePluginLine("plugin_name", plugin_name)
    PluginContent = PluginContent & GeneratePluginLine("plugin_family", plugin_family)
    PluginContent = PluginContent & GeneratePluginLine("plugin_created_date", plugin_created_date)
    PluginContent = PluginContent & GeneratePluginLine("plugin_created_name", plugin_created_name)
    PluginContent = PluginContent & GeneratePluginLine("plugin_created_email", plugin_created_email)
    PluginContent = PluginContent & GeneratePluginLine("plugin_created_web", plugin_created_web)
    PluginContent = PluginContent & GeneratePluginLine("plugin_created_company", plugin_created_company)
    PluginContent = PluginContent & GeneratePluginLine("plugin_updated_name", plugin_updated_name)
    PluginContent = PluginContent & GeneratePluginLine("plugin_updated_email", plugin_updated_email)
    PluginContent = PluginContent & GeneratePluginLine("plugin_updated_web", plugin_updated_web)
    PluginContent = PluginContent & GeneratePluginLine("plugin_updated_company", plugin_updated_company)
    PluginContent = PluginContent & GeneratePluginLine("plugin_updated_date", plugin_updated_date)
    PluginContent = PluginContent & GeneratePluginLine("plugin_version", plugin_version)
    PluginContent = PluginContent & GeneratePluginLine("plugin_changelog", plugin_changelog)
    PluginContent = PluginContent & GeneratePluginLine("plugin_protocol", plugin_protocol)
    PluginContent = PluginContent & GeneratePluginLine("plugin_port", plugin_port)
    PluginContent = PluginContent & GeneratePluginLine("plugin_procedure_detection", plugin_procedure_detection)
    PluginContent = PluginContent & GeneratePluginLine("plugin_procedure_exploit", plugin_procedure_exploit)
    PluginContent = PluginContent & GeneratePluginLine("plugin_detection_accuracy", plugin_detection_accuracy)
    PluginContent = PluginContent & GeneratePluginLine("plugin_exploit_accuracy", plugin_exploit_accuracy)
    PluginContent = PluginContent & GeneratePluginLine("plugin_comment", plugin_comment)
    
    PluginContent = PluginContent & GeneratePluginLine("bug_published_name", bug_published_name)
    PluginContent = PluginContent & GeneratePluginLine("bug_published_email", bug_published_email)
    PluginContent = PluginContent & GeneratePluginLine("bug_published_web", bug_published_web)
    PluginContent = PluginContent & GeneratePluginLine("bug_published_company", bug_published_company)
    PluginContent = PluginContent & GeneratePluginLine("bug_published_date", bug_published_date)
    PluginContent = PluginContent & GeneratePluginLine("bug_advisory", bug_advisory)
    PluginContent = PluginContent & GeneratePluginLine("bug_produced_name", bug_produced_name)
    PluginContent = PluginContent & GeneratePluginLine("bug_produced_email", bug_produced_email)
    PluginContent = PluginContent & GeneratePluginLine("bug_produced_web", bug_produced_web)
    PluginContent = PluginContent & GeneratePluginLine("bug_affected", bug_affected)
    PluginContent = PluginContent & GeneratePluginLine("bug_not_affected", bug_not_affected)
    PluginContent = PluginContent & GeneratePluginLine("bug_false_positives", bug_false_positives)
    PluginContent = PluginContent & GeneratePluginLine("bug_false_negatives", bug_false_negatives)
    PluginContent = PluginContent & GeneratePluginLine("bug_vulnerability_class", bug_vulnerability_class)
    PluginContent = PluginContent & GeneratePluginLine("bug_description", bug_description)
    PluginContent = PluginContent & GeneratePluginLine("bug_solution", bug_solution)
    PluginContent = PluginContent & GeneratePluginLine("bug_fixing_time", bug_fixing_time)
    PluginContent = PluginContent & GeneratePluginLine("bug_exploit_availability", bug_exploit_availability)
    PluginContent = PluginContent & GeneratePluginLine("bug_exploit_url", bug_exploit_url)
    PluginContent = PluginContent & GeneratePluginLine("bug_remote", bug_remote)
    PluginContent = PluginContent & GeneratePluginLine("bug_local", bug_local)
    PluginContent = PluginContent & GeneratePluginLine("bug_severity", bug_severity)
    PluginContent = PluginContent & GeneratePluginLine("bug_popularity", bug_popularity)
    PluginContent = PluginContent & GeneratePluginLine("bug_simplicity", bug_simplicity)
    PluginContent = PluginContent & GeneratePluginLine("bug_impact", bug_impact)
    PluginContent = PluginContent & GeneratePluginLine("bug_risk", bug_risk)
    PluginContent = PluginContent & GeneratePluginLine("bug_nessus_risk", bug_nessus_risk)
    PluginContent = PluginContent & GeneratePluginLine("bug_iss_scanner_rating", bug_iss_scanner_rating)
    PluginContent = PluginContent & GeneratePluginLine("bug_netrecon_rating", bug_netrecon_rating)
    PluginContent = PluginContent & GeneratePluginLine("bug_check_tool", bug_check_tool)
    
    PluginContent = PluginContent & GeneratePluginLine("source_cve", source_cve)
    PluginContent = PluginContent & GeneratePluginLine("source_certvu_id", source_certvu_id)
    PluginContent = PluginContent & GeneratePluginLine("source_cert_id", source_cert_id)
    PluginContent = PluginContent & GeneratePluginLine("source_uscertta_id", source_uscertta_id)
    PluginContent = PluginContent & GeneratePluginLine("source_securityfocus_bid", source_securityfocus_bid)
    PluginContent = PluginContent & GeneratePluginLine("source_osvdb_id", source_osvdb_id)
    PluginContent = PluginContent & GeneratePluginLine("source_secunia_id", source_secunia_id)
    PluginContent = PluginContent & GeneratePluginLine("source_securiteam_url", source_securiteam_url)
    PluginContent = PluginContent & GeneratePluginLine("source_securitytracker_id", source_securitytracker_id)
    PluginContent = PluginContent & GeneratePluginLine("source_scip_id", source_scip_id)
    PluginContent = PluginContent & GeneratePluginLine("source_tecchannel_id", source_tecchannel_id)
    PluginContent = PluginContent & GeneratePluginLine("source_heise_news", source_heise_news)
    PluginContent = PluginContent & GeneratePluginLine("source_heise_security", source_heise_security)
    PluginContent = PluginContent & GeneratePluginLine("source_aerasec_id", source_aerasec_id)
    PluginContent = PluginContent & GeneratePluginLine("source_nessus_id", source_nessus_id)
    PluginContent = PluginContent & GeneratePluginLine("source_issxforce_id", source_issxforce_id)
    PluginContent = PluginContent & GeneratePluginLine("source_snort_id", source_snort_id)
    PluginContent = PluginContent & GeneratePluginLine("source_arachnids_id", source_arachnids_id)
    PluginContent = PluginContent & GeneratePluginLine("source_mssb_id", source_mssb_id)
    PluginContent = PluginContent & GeneratePluginLine("source_mskb_id", source_mskb_id)
    PluginContent = PluginContent & GeneratePluginLine("source_netbsdsa_id", source_netbsdsa_id)
    PluginContent = PluginContent & GeneratePluginLine("source_rhsa_id", source_rhsa_id)
    PluginContent = PluginContent & GeneratePluginLine("source_ciac_id", source_ciac_id)
    PluginContent = PluginContent & GeneratePluginLine("source_literature", source_literature)
    PluginContent = PluginContent & GeneratePluginLine("source_misc", source_misc)
        
    'Kill all useless lines to save space and increase performance
    PluginContentArray = Split(PluginContent, vbNewLine, , vbBinaryCompare)
    PluginContentItemCount = UBound(PluginContentArray)
    PluginContent = vbNullString
    
    For i = 0 To PluginContentItemCount
        If InStrB(4, PluginContentArray(i), "></", vbBinaryCompare) = 0 Then
                PluginContent = PluginContent & PluginContentArray(i) & vbNewLine
        End If
    Next i
        
    'Write the collected data into the file; the plugin name will be the file name
    On Error Resume Next
    Open Filename & ".plugin" For Output As 1
        Print #1, PluginContent
    Close
End Sub

Public Sub GenerateActualATKPluginsList()
    Dim strFileContent As String
    Dim i As Integer
    Dim intLoadedPlugins As Integer

    'Set the progress bar to zero
    frmMain.SetProgress 0

    'Count the loaded plugins
    intLoadedPlugins = frmMain.filATKPlugins.ListCount - 1

    For i = 0 To intLoadedPlugins
        'Increase the progress bar. The On Error Resume Next prevents senseless
        'values that could lead to a programm error.
        On Error Resume Next
        frmMain.SetProgress (100 / intLoadedPlugins) * i
        
        'Everytime select the new plugin and do the check until finish
        'Set lsvPlugins.SelectedItem = lsvPlugins.ListItems(i)
        frmMain.filATKPlugins.ListIndex = i

        strFileContent = strFileContent & _
            plugin_id & ";" & frmMain.filATKPlugins.Filename & ";" & plugin_version & ";" & plugin_updated_date & ";" & plugin_filesize & vbNewLine
    
    Next i
    
    On Error Resume Next ' Needed if there are no write permissions
    Open application_plugin_directory & "\pluginslist.txt" For Output As #1
        Print #1, strFileContent
    Close
    
    'Set the progress bar to 100
    frmMain.SetProgress 100
End Sub

