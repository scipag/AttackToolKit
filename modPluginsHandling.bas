Attribute VB_Name = "modPluginsHandling"
' This module reads the plugin, performs the parsing and writes the result
' in the global variables.

Option Explicit

                                            'In this column is a short description
                                            'of the meaning of the variables. You
                                            'find more information about them on the
                                            'project web site or the readme.

Public plugin_filename As String            'The filename of the plugin. This value is not
                                            'saved in the plugin file.
Public plugin_filepath As String            'The path of the plugin file.
Public plugin_filesize As String            'The filesize of the plugin. Also not a saved
                                            'value in the plugin file.
Public plugin_filecontent As String         'The whole content of the plugin file

Public plugin_id As String                  'Unique ID of the ATK plugin
Public plugin_name As String                'Plugin name of the ATK plugin
Public plugin_family As String              'Plugin family of the ATK plugin
Public plugin_created_name As String        'Name of the person who created the plugin
Public plugin_created_email As String       'Email of the person who created the plugin
Public plugin_created_web As String         'Web site of the person who created the plugin
Public plugin_created_company As String     'Companyname of the person who created the plugin
Public plugin_created_date As String        'When was the ATK plugin created
Public plugin_updated_name As String        'Name of the person who updated the ATK plugin
Public plugin_updated_email As String       'Email of the person who updated the ATK plugin
Public plugin_updated_web As String         'Web site of the person who updated the ATK plugin
Public plugin_updated_company As String     'Company name of the person who updated the ATK plugin
Public plugin_updated_date As String        'When was the ATK plugin updated the last time
Public plugin_version As String             'Which version of the plugin is it
Public plugin_changelog As String           'The changelog of the plugin
Public plugin_protocol As String            'Which protocol does the plugin use
Public plugin_port As String                'Which port does the ATK plugin use
Public plugin_procedure_detection As String 'The request procedure for detection
Public plugin_procedure_exploit As String   'The request procedure for exploit
Public plugin_detection_accuracy As String  'The accuracy for detection
Public plugin_exploit_accuracy As String    'The accuracy for exploiting
Public plugin_comment As String             'Some words about the ATK plugin

Public bug_published_name As String         'Who published the bug first
Public bug_published_email As String        'Who published the bug first
Public bug_published_web As String          'Who published the bug first
Public bug_published_company As String      'Who published the bug first
Public bug_published_date As String         'Who published the bug first
Public bug_produced_name As String          'Who produced the product
Public bug_produced_email As String         'Who produced the product
Public bug_produced_web As String           'Who produced the product
Public bug_advisory As String               'What is the name and URL of the advisory
Public bug_affected As String               'Which systems and solutions are affected
Public bug_not_affected As String           'Which systems and solutions are not affected
Public bug_false_positives As String        'Known false-positives
Public bug_false_negatives As String        'Known false-negatives
Public bug_vulnerability_class As String    'The class of the vulnerability
Public bug_local As String                  'A local vulnerability
Public bug_remote As String                 'A remote vulnerability
Public bug_description As String            'The description of the vulnerability
Public bug_solution As String               'The solution(s) for the vulnerability
Public bug_fixing_time As String            'The time needed to fix the bug (e.g. hours)
Public bug_exploit_availability As String   'The existence of an exploit
Public bug_exploit_url As String            'The URL of the exploit
Public bug_severity As String               'The severity of the vulnerability
Public bug_popularity As String             'The popularity of the vulnerability (1 to 10)
Public bug_simplicity As String             'The simplicity of the bug (1 to 10)
Public bug_impact As String                 'The impact level of the bug (1 to 10)
Public bug_risk As String                   'The rist of the bug (1 to 10)
Public bug_nessus_risk As String            'The risk level of the bug by Nessus
Public bug_iss_scanner_rating As String     'The risk level by ISS Scanners
Public bug_netrecon_rating As String        'The risk level by Symantec NetRecon
Public bug_check_tool As String             'List of tools that are able to check the bug

Public source_cve As String                 'The unique CAN or CVE number of the vulnerability
Public source_certvu_id As String           'The unique CERT Vulnerability ID
Public source_cert_id As String             'The unique CERT ID
Public source_uscertta_id As String         'The unique US-CERT Technical Advisory ID
Public source_securityfocus_bid As String   'The unique SecurityFocus/Bugtraq ID
Public source_osvdb_id As String            'The unique Open Source Vulnerability Data Base ID
Public source_secunia_id As String          'The unique Secunia ID of the vulnerability
Public source_securiteam_url As String      'The SecuriTeam.com URL
Public source_securitytracker_id As String  'The SecurityTracker ID
Public source_scip_id As String             'The unique scipID of the vulnerability
Public source_tecchannel_id As String       'The unique tecchannel ID
Public source_heise_news As String          'The unique Heise News
Public source_heise_security As String      'The unique Heise Security
Public source_aerasec_id As String          'The unique AeraSec ID
Public source_nessus_id As String           'The unique Nessus ID of the vulnerability
Public source_issxforce_id As String        'The unique ISS X-Force ID
Public source_snort_id As String            'The unique Snort ID of the vulnerability
Public source_arachnids_id As String        'The unique ArachnIDS ID
Public source_mssb_id As String             'The unique Microsoft Security Bulletin ID
Public source_mskb_id As String             'The unique Microsoft Knowledge-Base Article ID
Public source_netbsdsa_id As String         'The unique NetBSD Security Advisory ID
Public source_rhsa_id As String             'The unique Red Hat Security Advisory ID
Public source_ciac_id As String             'The unique CIAC ID
Public source_literature As String          'List of books about the flaw
Public source_misc As String                'List of other sources (e.g. TV shows)

Public session_procedure_type As String     'Type of the procedure (detection or exploit)
Public session_procedure_commands As String 'The commands of the defined procedure
Public session_triggers As String           'The triggers of the check and type
Public session_trigger_match As String      'The matching trigger

' *******************************************************************
' * Reset all plugin variables. This is usually done before new     *
' * data is read. Old "garbage" is prevented and the software don't *
' * need so much ressources during runtime.                         *
' *******************************************************************

Public Sub ClearAllPluginVariables()
    If frmMain.lblVulnerabilityState.BackColor <> &HE0E0E0 Then
        Dim strAlertingText As String
        
        strAlertingText = "The vulnerability was not tested. " & _
            "Please run the selected plugin to determine the existence of the flaw."
        
        'Message if the vulnerability was found
        frmMain.lblVulnerabilityState.Caption = strAlertingText
        frmMain.lblVulnerabilityState.BackColor = &HE0E0E0
    End If
    
    plugin_id = vbNullString
    plugin_name = vbNullString
    plugin_family = vbNullString
    plugin_created_name = vbNullString
    plugin_created_email = vbNullString
    plugin_created_web = vbNullString
    plugin_created_company = vbNullString
    plugin_created_date = vbNullString
    plugin_updated_name = vbNullString
    plugin_updated_email = vbNullString
    plugin_updated_web = vbNullString
    plugin_updated_company = vbNullString
    plugin_updated_date = vbNullString
    plugin_version = vbNullString
    plugin_changelog = vbNullString
    plugin_protocol = vbNullString
    plugin_port = vbNullString
    plugin_procedure_detection = vbNullString
    plugin_procedure_exploit = vbNullString
    plugin_detection_accuracy = vbNullString
    plugin_exploit_accuracy = vbNullString
    plugin_comment = vbNullString
    
    bug_published_name = vbNullString
    bug_published_email = vbNullString
    bug_published_web = vbNullString
    bug_published_company = vbNullString
    bug_published_date = vbNullString
    bug_produced_name = vbNullString
    bug_produced_email = vbNullString
    bug_produced_web = vbNullString
    bug_advisory = vbNullString
    bug_affected = vbNullString
    bug_not_affected = vbNullString
    bug_false_positives = vbNullString
    bug_false_negatives = vbNullString
    bug_vulnerability_class = vbNullString
    bug_local = vbNullString
    bug_remote = vbNullString
    bug_description = vbNullString
    bug_solution = vbNullString
    bug_fixing_time = vbNullString
    bug_exploit_availability = vbNullString
    bug_exploit_url = vbNullString
    bug_severity = vbNullString
    bug_popularity = vbNullString
    bug_simplicity = vbNullString
    bug_impact = vbNullString
    bug_risk = vbNullString
    bug_nessus_risk = vbNullString
    bug_iss_scanner_rating = vbNullString
    bug_netrecon_rating = vbNullString
    bug_check_tool = vbNullString
    
    source_cve = vbNullString
    source_certvu_id = vbNullString
    source_cert_id = vbNullString
    source_uscertta_id = vbNullString
    source_securityfocus_bid = vbNullString
    source_osvdb_id = vbNullString
    source_secunia_id = vbNullString
    source_securiteam_url = vbNullString
    source_securitytracker_id = vbNullString
    source_scip_id = vbNullString
    source_tecchannel_id = vbNullString
    source_heise_news = vbNullString
    source_heise_security = vbNullString
    source_aerasec_id = vbNullString
    source_nessus_id = vbNullString
    source_issxforce_id = vbNullString
    source_snort_id = vbNullString
    source_arachnids_id = vbNullString
    source_mssb_id = vbNullString
    source_mskb_id = vbNullString
    source_netbsdsa_id = vbNullString
    source_rhsa_id = vbNullString
    source_ciac_id = vbNullString
    source_literature = vbNullString
    source_misc = vbNullString
    
    session_procedure_type = vbNullString
    session_procedure_commands = vbNullString
    session_triggers = vbNullString
    session_trigger_match = vbNullString
End Sub

Public Function ReadPluginFromFile(ByRef strFileName As String, _
                                    ByRef strFilePath As String) As String
                                    
    Dim strPluginFullFileName As String 'The full path and name of the plugin file

    'This is just a workaround because the Open dialog can't split file name and path
    If InStrB(1, strFileName, "\", vbBinaryCompare) Then
        strPluginFullFileName = strFileName
    Else
        strPluginFullFileName = strFilePath & "\" & strFileName
    End If

    'Check the existence of the file
    On Error Resume Next
    If LenB(Dir$(strPluginFullFileName)) Then
        'Flush the old plugin content before loading new data
        plugin_filecontent = vbNullString
        
        plugin_filepath = strFilePath
        plugin_filename = strFileName
        
        'Open and read the plugin file
        Open strPluginFullFileName For Input As 1
            plugin_filecontent = Input(LOF(1), #1)
        Close
        ReadPluginFromFile = plugin_filecontent
    
        'Set the plugin silesize
        plugin_filesize = Len(plugin_filecontent)
    Else
        Call errPluginDoesNotExist(strFilePath & "\" & strFileName)
    End If
End Function

Public Function HowManyLoadedPlugins() As Integer
    If frmMain.filATKPlugins.ListCount Then
        If frmMain.tvwPlugins.Nodes.Count Then
            On Error Resume Next 'Workaround to prevent kill if node is not available.
            HowManyLoadedPlugins = frmMain.tvwPlugins.Nodes("ATK ID").Children
        End If
    Else
        HowManyLoadedPlugins = 0
    End If

    If frmMain.filNASLPlugins.ListCount Then
        If frmMain.tvwPlugins.Nodes.Count Then
            On Error Resume Next 'Workaround to prevent kill if node is not available.
            HowManyLoadedPlugins = HowManyLoadedPlugins + frmMain.tvwPlugins.Nodes("NASL ID").Children
        End If
    Else
        HowManyLoadedPlugins = HowManyLoadedPlugins + 0
    End If
End Function

' ******************************************************************
' * This function extracts the possible ISBN number from a string. *
' * It works but there is one really nasty limitation:             *
' * 1. The ISBN numbers have to be written without the suggested   *
' *    delimiters as like spaces or dashes.                        *
' * This limitation should be fixed in an upcoming release.        *
' ******************************************************************

Public Function GetISBNFromString(TextString As String) As String
    Dim WordArray() As String
    Dim PossibleISBNNumber As String
    Dim i As Integer
    Dim j As Integer
    
    WordArray = Split(TextString, " ")
    
    For i = 0 To UBound(WordArray)
        'Reset the possible ISBN number for the next text block
        PossibleISBNNumber = vbNullString
        
        For j = 1 To Len(WordArray(i))
            If Len(PossibleISBNNumber) < 12 Then
                If Mid$(WordArray(i), j, 1) Like "[0-9]" Then
                    PossibleISBNNumber = PossibleISBNNumber & Mid$(WordArray(i), j, 1)
                ElseIf InStrB(j, WordArray(i), "X", vbBinaryCompare) Then
                    PossibleISBNNumber = PossibleISBNNumber & Mid$(WordArray(i), j, 1)
                End If
                    
                If PossibleISBNNumber Like "#########?" Then
                    GetISBNFromString = PossibleISBNNumber
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Public Sub SetPluginSessionProcedure()
    If LenB(plugin_procedure_detection) Then
        session_procedure_type = "detection"
        session_procedure_commands = plugin_procedure_detection
    ElseIf LenB(plugin_procedure_exploit) Then
        session_procedure_type = "exploit"
        session_procedure_commands = plugin_procedure_exploit
    Else
        session_procedure_type = vbNullString
        session_procedure_commands = vbNullString
    End If
End Sub

Public Sub ActivatePluginMenuModes()
    If LenB(plugin_procedure_detection) Then
        frmMain.mnuPluginsRunDetectionItem.Enabled = True
    Else
        frmMain.mnuPluginsRunDetectionItem.Enabled = False
    End If

    If LenB(plugin_procedure_exploit) Then
        frmMain.mnuPluginsRunExploitItem.Enabled = True
    Else
        frmMain.mnuPluginsRunExploitItem.Enabled = False
    End If
End Sub
