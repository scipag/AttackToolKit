Attribute VB_Name = "modPluginsHandlingATKRead"
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2004-12-28                                                           *
' * - Renamed ParseATKPluginTag to ParseAMLTag. AML is the new name for the whole    *
' *   XML based plugin and suggestions structure.                                    *
' * Version 4.0 2004-12-08                                                           *
' * - Improved the speed of the XML tag parsing; especially if small plugins are     *
' *   loaded.                                                                        *
' * - Improved the possibility of parsing the XML tags case insensitive (liberate).  *
' * Version 3.0 2004-11-13                                                           *
' * - Introduced the plugin_changelog, bug_false_positives and _negatives fields.    *
' * Version 3.0 2004-11-01                                                           *
' * - Replaced all useless functions with normal subs.                               *
' * Version 3.0 2004-10-01                                                           *
' * - Added the session variants.                                                    *
' * Version 2.1 2004-09-08                                                           *
' * - Additional filling of the fields source_misc and source_literature if empty.   *
' * Version 2.0 2004-08-24                                                           *
' * - Changed Len to LenB checking during ATK plugin parsing. This increases the     *
' *   speed of the procedure very much.                                              *
' * - Increased the speed of some array handling during the writing of a plugin.     *
' * Version 2.0 2004-08-16                                                           *
' * - Corrected an error with the closing tag bug_not_affected during writing.       *
' ************************************************************************************

Public Function ParseAMLTag(ByRef strTag As String, _
                                    ByRef strPluginContent As String) As String
                                    
    Dim TempArray() As String       'A temporary array for the splitting and parsing

    'Check for the presence of the beginning tag
    If InStrB(1, strPluginContent, "<" & strTag & ">", vbBinaryCompare) Then
        'Split the beginning tag
        TempArray() = Split(strPluginContent, "<" & strTag & ">", , vbBinaryCompare)
        'Check for the presence of the ending tag
        If InStrB(1, TempArray(1), "</" & strTag & ">", vbBinaryCompare) Then
            'Split the ending tag
            TempArray = Split(TempArray(1), "</" & strTag & ">", , vbBinaryCompare)
            'Check for the length of a result
            If LenB(TempArray(0)) Then
                'Write the result back
                ParseAMLTag = TempArray(0)
            End If
        End If
    End If
End Function

Public Sub ParseATKPlugin(ByRef strATKPluginContent As String)
    'Clear the values from the last plugin to prevent misunderstandings
    Call ClearAllPluginVariables    'Plugin variables itself
    'Call ClearAllResponseVariables  'Plugin last response
    
    'Parse the different fields/tags
    plugin_id = ParseAMLTag("plugin_id", strATKPluginContent)
    plugin_name = ParseAMLTag("plugin_name", strATKPluginContent)
    plugin_family = ParseAMLTag("plugin_family", strATKPluginContent)
    plugin_created_name = ParseAMLTag("plugin_created_name", strATKPluginContent)
    plugin_created_email = ParseAMLTag("plugin_created_email", strATKPluginContent)
    plugin_created_web = ParseAMLTag("plugin_created_web", strATKPluginContent)
    plugin_created_company = ParseAMLTag("plugin_created_company", strATKPluginContent)
    plugin_created_date = ParseAMLTag("plugin_created_date", strATKPluginContent)
    plugin_updated_name = ParseAMLTag("plugin_updated_name", strATKPluginContent)
    plugin_updated_email = ParseAMLTag("plugin_updated_email", strATKPluginContent)
    plugin_updated_web = ParseAMLTag("plugin_updated_web", strATKPluginContent)
    plugin_updated_company = ParseAMLTag("plugin_updated_company", strATKPluginContent)
    plugin_updated_date = ParseAMLTag("plugin_updated_date", strATKPluginContent)
    plugin_version = ParseAMLTag("plugin_version", strATKPluginContent)
    plugin_changelog = ParseAMLTag("plugin_changelog", strATKPluginContent)
    plugin_protocol = ParseAMLTag("plugin_protocol", strATKPluginContent)
    plugin_port = ParseAMLTag("plugin_port", strATKPluginContent)
    plugin_procedure_detection = ParseAMLTag("plugin_procedure_detection", strATKPluginContent)
    plugin_procedure_exploit = ParseAMLTag("plugin_procedure_exploit", strATKPluginContent)
    plugin_detection_accuracy = ParseAMLTag("plugin_detection_accuracy", strATKPluginContent)
    plugin_exploit_accuracy = ParseAMLTag("plugin_exploit_accuracy", strATKPluginContent)
    plugin_comment = ParseAMLTag("plugin_comment", strATKPluginContent)
    
    bug_published_name = ParseAMLTag("bug_published_name", strATKPluginContent)
    bug_published_email = ParseAMLTag("bug_published_email", strATKPluginContent)
    bug_published_web = ParseAMLTag("bug_published_web", strATKPluginContent)
    bug_published_company = ParseAMLTag("bug_published_company", strATKPluginContent)
    bug_published_date = ParseAMLTag("bug_published_date", strATKPluginContent)
    bug_advisory = ParseAMLTag("bug_advisory", strATKPluginContent)
    bug_produced_name = ParseAMLTag("bug_produced_name", strATKPluginContent)
    bug_produced_email = ParseAMLTag("bug_produced_email", strATKPluginContent)
    bug_produced_web = ParseAMLTag("bug_produced_web", strATKPluginContent)
    bug_affected = ParseAMLTag("bug_affected", strATKPluginContent)
    bug_not_affected = ParseAMLTag("bug_not_affected", strATKPluginContent)
    bug_false_positives = ParseAMLTag("bug_false_positives", strATKPluginContent)
    bug_false_negatives = ParseAMLTag("bug_false_negatives", strATKPluginContent)
    bug_local = ParseAMLTag("bug_local", strATKPluginContent)
    bug_remote = ParseAMLTag("bug_remote", strATKPluginContent)
    bug_vulnerability_class = ParseAMLTag("bug_vulnerability_class", strATKPluginContent)
    bug_description = ParseAMLTag("bug_description", strATKPluginContent)
    bug_solution = ParseAMLTag("bug_solution", strATKPluginContent)
    bug_fixing_time = ParseAMLTag("bug_fixing_time", strATKPluginContent)
    bug_exploit_availability = ParseAMLTag("bug_exploit_availability", strATKPluginContent)
    bug_exploit_url = ParseAMLTag("bug_exploit_url", strATKPluginContent)
    bug_severity = ParseAMLTag("bug_severity", strATKPluginContent)
    bug_popularity = ParseAMLTag("bug_popularity", strATKPluginContent)
    bug_simplicity = ParseAMLTag("bug_simplicity", strATKPluginContent)
    bug_impact = ParseAMLTag("bug_impact", strATKPluginContent)
    bug_risk = ParseAMLTag("bug_risk", strATKPluginContent)
    bug_nessus_risk = ParseAMLTag("bug_nessus_risk", strATKPluginContent)
    bug_iss_scanner_rating = ParseAMLTag("bug_iss_scanner_rating", strATKPluginContent)
    bug_netrecon_rating = ParseAMLTag("bug_netrecon_rating", strATKPluginContent)
    bug_check_tool = ParseAMLTag("bug_check_tool", strATKPluginContent)
    
    source_cve = ParseAMLTag("source_cve", strATKPluginContent)
    source_certvu_id = ParseAMLTag("source_certvu_id", strATKPluginContent)
    source_cert_id = ParseAMLTag("source_cert_id", strATKPluginContent)
    source_uscertta_id = ParseAMLTag("source_uscertta_id", strATKPluginContent)
    source_securityfocus_bid = ParseAMLTag("source_securityfocus_bid", strATKPluginContent)
    source_osvdb_id = ParseAMLTag("source_osvdb_id", strATKPluginContent)
    source_secunia_id = ParseAMLTag("source_secunia_id", strATKPluginContent)
    source_securiteam_url = ParseAMLTag("source_securiteam_url", strATKPluginContent)
    source_securitytracker_id = ParseAMLTag("source_securitytracker_id", strATKPluginContent)
    source_scip_id = ParseAMLTag("source_scip_id", strATKPluginContent)
    source_tecchannel_id = ParseAMLTag("source_tecchannel_id", strATKPluginContent)
    source_heise_news = ParseAMLTag("source_heise_news", strATKPluginContent)
    source_heise_security = ParseAMLTag("source_heise_security", strATKPluginContent)
    source_aerasec_id = ParseAMLTag("source_aerasec_id", strATKPluginContent)
    source_nessus_id = ParseAMLTag("source_nessus_id", strATKPluginContent)
    source_issxforce_id = ParseAMLTag("source_issxforce_id", strATKPluginContent)
    source_snort_id = ParseAMLTag("source_snort_id", strATKPluginContent)
    source_arachnids_id = ParseAMLTag("source_arachnids_id", strATKPluginContent)
    source_mssb_id = ParseAMLTag("source_mssb_id", strATKPluginContent)
    source_mskb_id = ParseAMLTag("source_mskb_id", strATKPluginContent)
    source_netbsdsa_id = ParseAMLTag("source_netbsdsa_id", strATKPluginContent)
    source_rhsa_id = ParseAMLTag("source_rhsa_id", strATKPluginContent)
    source_ciac_id = ParseAMLTag("source_ciac_id", strATKPluginContent)
    source_literature = ParseAMLTag("source_literature", strATKPluginContent)
    source_misc = ParseAMLTag("source_misc", strATKPluginContent)
    
    Call SetPluginSessionProcedure
    Call ActivatePluginMenuModes
End Sub
