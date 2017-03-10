Attribute VB_Name = "modConfigHandling"
Option Explicit

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2004-12-12                                                           *
' * - Really introduced the application_response_directory.                          *
' * Version 4.0 2004-12-07                                                           *
' * - Added the procedures for handling the external plugin editor.                  *
' * Version 3.0 2004-10-11                                                           *
' * - Added a default value if it was not defined if logs should be activated.       *
' * - Added the whole procedures for handling the new logging security levels.       *
' ************************************************************************************

'For getting the username
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Change this "constant" on every new release to write the right software name
'and version.
Public Const application_name As String = "Attack Tool Kit 4.1"
Public Const application_website_url As String = "http://www.computec.ch/projekte/atk/"

Public application_attack_mode As String
Public application_attack_timeout As Long
Public application_help_url As String
Public application_icmp_mapping_enable As Boolean
Public application_icmp_mapping_ignore_enable As Boolean
Public application_log_directory As String
Public application_log_enable As Boolean
Public application_log_security_level As Integer
Public application_no_dos_enable As Boolean
Public application_plugin_directory As String
Public application_plugin_download_url As String
Public application_plugin_external_editor As String
Public application_report_directory As String
Public application_report_open_enable As Boolean
Public application_response_directory As String
Public application_searchengine_url As String
Public application_silent_checks_enable As Boolean
Public application_sleep_time_default As Long
Public application_speech_enable As Boolean
Public application_suggestion_directory As String
Public application_suggestion_enable As Boolean
Public application_vulnerability_found_alert_enable As Boolean
Public application_vulnerability_not_found_alert_enable As Boolean
'Public ReportTemplateDirectory As String

Public Target As String

Public application_configuration_filename As String
Public system_username As String

Public Sub LoadConfigFromFile(Optional ByRef strConfigurationFileName As String)
    Dim intFreeFile As Integer
    Dim TempString As String
    
    'This boolean values indicate that a value could be found. We need this state
    'to find missing or wrong input and correct them. This list is alphabetically until 1.1
    Dim application_log_enableV As Boolean
    Dim application_speech_enableV As Boolean
    Dim application_suggestion_enableV As Boolean
    Dim application_vulnerability_found_alert_enableV As Boolean
    Dim application_vulnerability_not_found_alert_enableV As Boolean
    Dim application_attack_modeV As Boolean
    Dim application_attack_timeoutV As Boolean
    Dim application_sleep_time_defaultV As Boolean
    Dim application_icmp_mapping_enableV As Boolean
    Dim application_no_dos_enableV As Boolean
    Dim application_silent_checks_enableV As Boolean
    Dim application_help_urlV As Boolean
    Dim application_log_directoryV As Boolean
    Dim application_log_security_levelV As Boolean
    Dim application_plugin_directoryV As Boolean
    Dim application_plugin_download_urlV As Boolean
    Dim application_plugin_external_editorV As Boolean
    Dim application_response_directoryV As Boolean
    Dim application_report_directoryV As Boolean
    Dim application_report_open_enableV As Boolean
    Dim application_icmp_mapping_ignore_enableV As Boolean
    Dim application_searchengine_urlV As Boolean
    Dim application_suggestion_directoryV As Boolean
    Dim TargetV As Boolean
        
    If LenB(strConfigurationFileName) Then
        application_configuration_filename = strConfigurationFileName
    Else
        application_configuration_filename = App.Path & "\configs\default.config"
    End If
    
    'WORKAROUND!
    application_response_directory = App.Path & "\responses\"
        
    'Check the existence of the config file
    If (Dir$(application_configuration_filename, 16) <> "") Then
        'Open and read the plugin file
        intFreeFile = FreeFile
        Open application_configuration_filename For Input As #intFreeFile
            Do While Not EOF(intFreeFile)
                Line Input #intFreeFile, TempString
                
                If Mid$(TempString, 1, 1) <> "#" Then
                    If InStrB(1, TempString, "=", vbBinaryCompare) Then
                        If Mid$(TempString, 1, 29) = "application_plugin_directory=" Then
                            application_plugin_directory = Mid$(TempString, 30, Len(TempString))
                            If LenB(application_plugin_directory) Then
                                application_plugin_directoryV = True
                            End If
                        ElseIf Mid$(TempString, 1, 32) = "application_plugin_download_url=" Then
                            application_plugin_download_url = Mid$(TempString, 33, Len(TempString))
                            If LenB(application_plugin_download_url) Then
                                application_plugin_download_urlV = True
                            End If
                        ElseIf Mid$(TempString, 1, 35) = "application_plugin_external_editor=" Then
                            application_plugin_external_editor = Mid$(TempString, 36, Len(TempString))
                            If LenB(application_plugin_external_editor) Then
                                application_plugin_external_editorV = True
                            End If
                        ElseIf Mid$(TempString, 1, 21) = "application_help_url=" Then
                            application_help_url = Mid$(TempString, 22, Len(TempString))
                            If LenB(application_help_url) Then
                                application_help_urlV = True
                            End If
                        ElseIf Mid$(TempString, 1, 26) = "application_speech_enable=" Then
                            application_speech_enableV = True
                            If Mid$(TempString, 27, Len(TempString)) = 1 Then
                                application_speech_enable = True
                            Else
                                application_speech_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 30) = "application_suggestion_enable=" Then
                            application_suggestion_enableV = True
                            If Mid$(TempString, 31, Len(TempString)) = 1 Then
                                application_suggestion_enable = True
                            Else
                                application_suggestion_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 45) = "application_vulnerability_found_alert_enable=" Then
                            application_vulnerability_found_alert_enableV = True
                            If Mid$(TempString, 46, Len(TempString)) = 1 Then
                                application_vulnerability_found_alert_enable = True
                            Else
                                application_vulnerability_found_alert_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 49) = "application_vulnerability_not_found_alert_enable=" Then
                            application_vulnerability_not_found_alert_enableV = True
                            If Mid$(TempString, 50, Len(TempString)) = 1 Then
                                application_vulnerability_not_found_alert_enable = True
                            Else
                                application_vulnerability_not_found_alert_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 33) = "application_suggestion_directory=" Then
                            application_suggestion_directory = Mid$(TempString, 34, Len(TempString))
                            If LenB(application_suggestion_directory) Then
                                application_suggestion_directoryV = True
                            End If
                        ElseIf Mid$(TempString, 1, 31) = "application_response_directory=" Then
                            application_response_directoryV = True
                            application_response_directory = Mid$(TempString, 32, Len(TempString))
                            'Load another directory if it does not exists.
                            If Not (Dir$(application_response_directory, 16) <> "") Then
                                application_response_directory = App.Path
                            End If
                        ElseIf Mid$(TempString, 1, 29) = "application_report_directory=" Then
                            application_report_directoryV = True
                            application_report_directory = Mid$(TempString, 30, Len(TempString))
                            'Load another directory if it does not exists.
                            If Not (Dir$(application_report_directory, 16) <> "") Then
                                application_report_directory = App.Path
                            End If
                        ElseIf Mid$(TempString, 1, 31) = "application_report_open_enable=" Then
                            application_report_open_enableV = True
                            If Mid$(TempString, 32, Len(TempString)) = 1 Then
                                application_report_open_enable = True
                            Else
                                application_report_open_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 27) = "application_attack_timeout=" Then
                            application_attack_timeoutV = True
                            application_attack_timeout = Mid$(TempString, 28, Len(TempString))
                        ElseIf Mid$(TempString, 1, 31) = "application_sleep_time_default=" Then
                            application_sleep_time_defaultV = True
                            application_sleep_time_default = Mid$(TempString, 32, Len(TempString))
                        ElseIf Mid$(TempString, 1, 24) = "application_attack_mode=" Then
                            application_attack_modeV = True
                            application_attack_mode = Mid$(TempString, 25, Len(TempString))
                        ElseIf Mid$(TempString, 1, 33) = "application_silent_checks_enable=" Then
                            application_silent_checks_enableV = True
                            If Mid$(TempString, 34, Len(TempString)) = 1 Then
                                application_silent_checks_enable = True
                            Else
                                application_silent_checks_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 26) = "application_no_dos_enable=" Then
                            application_no_dos_enableV = True
                            If Mid$(TempString, 27, Len(TempString)) = 1 Then
                                application_no_dos_enable = True
                            Else
                                application_no_dos_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 32) = "application_icmp_mapping_enable=" Then
                            application_icmp_mapping_enableV = True
                            If Mid$(TempString, 33, Len(TempString)) = 1 Then
                                application_icmp_mapping_enable = True
                            Else
                                application_icmp_mapping_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 39) = "application_icmp_mapping_ignore_enable=" Then
                            application_icmp_mapping_ignore_enableV = True
                            If Mid$(TempString, 40, Len(TempString)) = 1 Then
                                application_icmp_mapping_ignore_enable = True
                            Else
                                application_icmp_mapping_ignore_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 7) = "Target=" Then
                            Target = Mid$(TempString, 8, Len(TempString))
                            If LenB(Target) Then
                                TargetV = True
                            End If
                        ElseIf Mid$(TempString, 1, 23) = "application_log_enable=" Then
                            application_log_enableV = True
                            If Mid$(TempString, 24, Len(TempString)) = 1 Then
                                application_log_enable = True
                            Else
                                application_log_enable = False
                            End If
                        ElseIf Mid$(TempString, 1, 26) = "application_log_directory=" Then
                            application_log_directory = Mid$(TempString, 27, Len(TempString))
                            If LenB(application_log_directory) Then
                                application_log_directoryV = True
                            End If
                        ElseIf Mid$(TempString, 1, 31) = "application_log_security_level=" Then
                            application_log_security_levelV = True
                            If Mid$(TempString, 32, Len(TempString)) = 0 Then
                                application_log_security_level = 0
                            ElseIf Mid$(TempString, 32, Len(TempString)) = 1 Then
                                application_log_security_level = 1
                            ElseIf Mid$(TempString, 32, Len(TempString)) = 2 Then
                                application_log_security_level = 2
                            ElseIf Mid$(TempString, 32, Len(TempString)) = 3 Then
                                application_log_security_level = 3
                            ElseIf Mid$(TempString, 32, Len(TempString)) = 4 Then
                                application_log_security_level = 4
                            ElseIf Mid$(TempString, 32, Len(TempString)) = 5 Then
                                application_log_security_level = 5
                            ElseIf Mid$(TempString, 32, Len(TempString)) = 6 Then
                                application_log_security_level = 6
                            Else
                                application_log_security_level = 7
                            End If
                        ElseIf Mid$(TempString, 1, 29) = "application_searchengine_url=" Then
                            application_searchengine_url = Mid$(TempString, 30, Len(TempString))
                            If LenB(application_searchengine_url) Then
                                application_searchengine_urlV = True
                            End If
                        End If
                    End If
                End If
            Loop
        Close
    End If

    'Define default values if there is no config or no useful value in the config.
    'This is done to prevent false or missing input that would cause to an
    'undefined programm state.
    If application_plugin_directoryV = False Then
        application_plugin_directory = App.Path & "\plugins"
    End If
    
'    If application_plugin_external_editorV = False Then
'        application_plugin_external_editor = "notepad.exe"
'    End If
    
    If application_plugin_download_urlV = False Then
        application_plugin_download_url = application_website_url & "plugins/pluginslist/"
    End If
    
    If application_help_urlV = False Then
        application_help_url = application_website_url & "documentation/help/"
    End If
    
    If application_suggestion_enableV = False Then
        application_suggestion_enable = True
    End If
        
    If application_speech_enableV = False Then
        application_speech_enable = False
    End If
            
    If application_vulnerability_found_alert_enableV = False Then
        application_vulnerability_found_alert_enable = False
    End If
        
    If application_vulnerability_not_found_alert_enableV = False Then
        application_vulnerability_not_found_alert_enable = False
    End If
    
    If application_log_enableV = False Then
        application_log_enable = True
    End If
        
    If application_suggestion_directoryV = False Then
        application_suggestion_directory = App.Path & "\suggestions"
    End If
        
    If application_response_directoryV = False Then
        application_response_directory = App.Path & "\responses"
    End If
    
    If application_report_directoryV = False Then
        application_report_directory = App.Path & "\reports"
    End If
    
    If application_report_open_enableV = False Then
        application_report_open_enable = True
    End If
    
    If application_log_directoryV = False Then
        application_log_directory = App.Path & "\logs"
    End If
        
    If application_log_security_levelV = False Then
        application_log_security_level = 5
    End If
        
    If application_attack_timeoutV = False Then
        application_attack_timeout = 30000
    End If
        
    If application_sleep_time_defaultV = False Then
        application_sleep_time_default = 3000
    End If
        
    If application_attack_modeV = False Then
        application_attack_mode = "SingleCheck"
    End If
        
    If application_silent_checks_enableV = False Then
        application_silent_checks_enable = True
    End If
        
    If application_no_dos_enableV = False Then
        application_no_dos_enable = False
    End If
        
    If application_icmp_mapping_enableV = False Then
        application_icmp_mapping_enable = True
    End If
        
    If application_icmp_mapping_ignore_enableV = False Then
        application_icmp_mapping_ignore_enable = False
    End If
        
    If TargetV = False Then
        Target = "127.0.0.1"
    End If
    
    If application_searchengine_urlV = False Then
        application_searchengine_url = "http://www.google.com/search?q="
    End If
    
    'Change frame title so the user can see the next target
    frmMain.Caption = application_name & " - " & Target
End Sub

Public Sub WriteConfigurationToFile(ByRef strConfigurationFileName As String)
    Dim intFreeFile As Integer
    Dim ConfigContent As String
    
    application_configuration_filename = strConfigurationFileName
    
    'Write the config file header
    ConfigContent = "#" & vbNewLine & _
        "# " & application_name & " configuration file" & vbNewLine & _
        "# " & vbNewLine & _
        "#   Date       " & GetTodaysDate("/") & vbNewLine & _
        "#   Time       " & GetActualTime(":") & vbNewLine & _
        "#   File name  " & application_configuration_filename & vbNewLine & _
        "#   System     " & frmMain.wskTCPWinsock.Item(0).LocalIP & vbNewLine & _
        "#   User name  " & system_username & vbNewLine & _
        "#" & vbNewLine

    'Write a disclaimer
    ConfigContent = ConfigContent & _
        "# Disclaimer: This config file is generated automatically by the software" & vbNewLine & _
        "# itself during runtime. Please do not manually edit these values unless you" & vbNewLine & _
        "# do know what you're doing." & vbNewLine & _
        "#" & vbNewLine & _
        "# All values are shortly described. The left side specifies the variant were" & vbNewLine & _
        "# the data is saved and the right side defines the dynamicly saved value." & vbNewLine & _
        "# As it is used in most higher programming languages (e.g. Microsoft Visual" & vbNewLine & _
        "# Basic or ANSI C). The sharp sign can be used to uncomment a line. In this" & vbNewLine & _
        "# case the ATK uses the default value which is usually recommended." & vbNewLine & _
        "#" & vbNewLine & _
        "# See the online help, documentation and the official project web site" & vbNewLine & _
        "# http://www.computec.ch/projekte/atk/ for more details." & vbNewLine & _
        "#" & vbNewLine & vbNewLine
    
    'Write if logging should be done
    ConfigContent = ConfigContent & _
        "# The activate logs is a boolean variant were the activation for the logging" & vbNewLine & _
        "# feature is saved. The logging mechanism is used to do further analysis of" & vbNewLine & _
        "# scanning or debugging of the software. Activated logs may slow down the" & vbNewLine & _
        "# software a little bit. Activation of the logs is recommended. The value 0" & vbNewLine & _
        "# deactivates and 1 activates the logging. Activated logging with is the" & vbNewLine & _
        "# default value." & vbNewLine & _
        "#" & vbNewLine
    If application_log_enable = True Then
        ConfigContent = ConfigContent & "application_log_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_log_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write if speech output should be done
    ConfigContent = ConfigContent & _
        "# The activate speech is a boolean variant were the support for voice output" & vbNewLine & _
        "# is saved. This feature support the output of the application. Using spoken" & vbNewLine & _
        "# output slows down the software very much. The value 1 activates the speech" & vbNewLine & _
        "# and 0 deactivates it. The default value is 0 for deactivating the speech" & vbNewLine & _
        "# feature." & vbNewLine & _
        "#" & vbNewLine
    If application_speech_enable = True Then
        ConfigContent = ConfigContent & "application_speech_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_speech_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write the suggestions mode
    ConfigContent = ConfigContent & _
        "# The activate suggestion is a boolean value were the support for suggestions" & vbNewLine & _
        "# is saved. The value 1 stands for active and the opposit value 0 stands for" & vbNewLine & _
        "# not active. Activating the suggestions may slow down dedicated scans a" & vbNewLine & _
        "# little bit. But the suggestions are recommended for users who wants to be" & vbNewLine & _
        "# guided thru a penetration test." & vbNewLine & _
        "#" & vbNewLine
    If application_suggestion_enable = True Then
        ConfigContent = ConfigContent & "application_suggestion_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_suggestion_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write if alerting if the bug is found should be done
    ConfigContent = ConfigContent & _
        "# The alerting vuln found is a boolean variant were the activation for" & vbNewLine & _
        "# messages if a vulnerability has been found is saved. This informs the user" & vbNewLine & _
        "# by a big message and may be useful in very long plugin testing attempts." & vbNewLine & _
        "# The value 1 activates the notification and 0 deactivates it. The default" & vbNewLine & _
        "# value is 0 for deactivated notification." & vbNewLine & _
        "#" & vbNewLine
    If application_vulnerability_found_alert_enable = True Then
        ConfigContent = ConfigContent & "application_vulnerability_found_alert_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_vulnerability_found_alert_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write if alerting if the bug is found should be done
    ConfigContent = ConfigContent & _
        "# The alerting vuln not found is a boolean variant were the activation for" & vbNewLine & _
        "# messages if a vulnerability has not been found is saved. This informs the" & vbNewLine & _
        "# user by a big message and may be useful in very long plugin testing" & vbNewLine & _
        "# attempts. The value 1 activates the notification and 0 deactivates it. The" & vbNewLine & _
        "# default value is 0 for deactivated notification." & vbNewLine & _
        "#" & vbNewLine
    If application_vulnerability_not_found_alert_enable = True Then
        ConfigContent = ConfigContent & "application_vulnerability_not_found_alert_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_vulnerability_not_found_alert_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write the attack mode
    ConfigContent = ConfigContent & _
        "# The attack mode is a string were the attack mode is saved. The possible" & vbNewLine & _
        "# values are SingleCheck and FullAudit. The first one is used to run" & vbNewLine & _
        "# singular plugins in a penetration test. The second one is used to run all" & vbNewLine & _
        "# loaded plugins in a security audit. The ATK was written to verify potential" & vbNewLine & _
        "# flaws and exploit verified vulnerabilities. So It is recommended to run in" & vbNewLine & _
        "# SingleCheck mode. The enumeration before exploiting with the ATK should be" & vbNewLine & _
        "# done with a vulnerability scanner as like Nessus. This because such are" & vbNewLine & _
        "# faster in security auditing than the ATK. The ATK is more an exploiting" & vbNewLine & _
        "# framework as like MetaSploit Framework or raccess." & vbNewLine & _
        "#" & vbNewLine
    If application_attack_mode = "SingleCheck" Then
        ConfigContent = ConfigContent & "application_attack_mode=SingleCheck" & vbNewLine & vbNewLine
    ElseIf application_attack_mode = "FullAudit" Then
        ConfigContent = ConfigContent & "application_attack_mode=FullAudit" & vbNewLine & vbNewLine
    End If
    
    'Write the attack timeout
    ConfigContent = ConfigContent & _
        "# The attack timeout is an integer value were the default value for timeouts" & vbNewLine & _
        "# during attacks is saved. The attack abords if it takes longer than this" & vbNewLine & _
        "# timeout value. This value is saved in milliseconds. This means 10000 stands" & vbNewLine & _
        "# for 10 seconds. Keep in mind that Microsoft Windows is not a real-time" & vbNewLine & _
        "# operating system as like QNX is. This is why very well defined values are" & vbNewLine & _
        "# not used as exactly as it may be wanted. Too short timeouts prevent a" & vbNewLine & _
        "# plugin to be successful and accurate. The recommended value is 30000 ms" & vbNewLine & _
        "# (this is 30 seconds timeout per attack)." & vbNewLine & _
        "#" & vbNewLine & _
        "application_attack_timeout=" & application_attack_timeout & vbNewLine & vbNewLine
    
    'Write the default sleep value
    ConfigContent = ConfigContent & _
        "# The default sleep value is an integer were the default sleep time for the" & vbNewLine & _
        "# sleep command is saved. This is used to let the application or a plugin" & vbNewLine & _
        "# wait a defined time value. This value is saved in milliseconds. This means" & vbNewLine & _
        "# 1000 stands for 1 second. Keep in mind that Microsoft Windows is not a" & vbNewLine & _
        "# real-time operating system as like QNX is. This is why very well defined" & vbNewLine & _
        "# values are not used as exactly as it may be wanted. Too short sleep values" & vbNewLine & _
        "# prevent a plugin to be successful and accurate. Too long sleep values will" & vbNewLine & _
        "# take a check longer to finish. The recommended value is 3000 ms (this is 3" & vbNewLine & _
        "# seconds timeout per attack)." & vbNewLine & _
        "#" & vbNewLine & _
        "application_sleep_time_default=" & application_sleep_time_default & vbNewLine & vbNewLine
    
    'Write if ICMP mapping should be done
    ConfigContent = ConfigContent & _
        "# The do icmp mapping is a boolean variant were the support for icmp/ping" & vbNewLine & _
        "# mapping is saved. Icmp mapping is used to determine the existence and" & vbNewLine & _
        "# reachability of a target before scanning. This may prevent attack attempts" & vbNewLine & _
        "# to non existing nor non reachable hosts. Such pre-verification may save" & vbNewLine & _
        "# time for further analysis. The mapping feature makes the software a bit" & vbNewLine & _
        "# slower and uses a bit more network ressources. But the feature is" & vbNewLine & _
        "# recommended to be more accurate. The value 1 activates the mapping" & vbNewLine & _
        "# feature and 0 deactivates it. The default value is 1 for activated icmp" & vbNewLine & _
        "# mapping." & vbNewLine & _
        "#" & vbNewLine
    If application_icmp_mapping_enable = True Then
        ConfigContent = ConfigContent & "application_icmp_mapping_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_icmp_mapping_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write of denial of service checks should be done
    ConfigContent = ConfigContent & _
        "# The do no DoS checks is a boolean variant were the support for destructive" & vbNewLine & _
        "# and dangerous denial of service attacks is saved. You should deactivate the" & vbNewLine & _
        "# feature if the testing is done in a live environment that should not be" & vbNewLine & _
        "# harmed. The activation of DoS is recommended if the accuracy of a" & vbNewLine & _
        "# penetration test is very important. The value 1 activates the DoS save" & vbNewLine & _
        "# feature and 0 deactivates it. The default value is 0 for activated denial" & vbNewLine & _
        "# of service checks." & vbNewLine & _
        "#" & vbNewLine
    If application_no_dos_enable = True Then
        ConfigContent = ConfigContent & "application_no_dos_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_no_dos_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write if silent checks should be done
    ConfigContent = ConfigContent & _
        "# The silent check is a boolean variant were the support for silent checks is" & vbNewLine & _
        "# saved. Silent checks are working like the KB save feature in Nessus. The" & vbNewLine & _
        "# gathered data is used to verify other potential vulnerabilities without" & vbNewLine & _
        "# touching the target anymore. This makes the access attempts much faster and" & vbNewLine & _
        "# harder to detect by the target network. The silent check mode makes the" & vbNewLine & _
        "# software a bit slower. But the feature is recommended in all circumstances" & vbNewLine & _
        "# when very fast verification of potential vulnerabilities is required. The" & vbNewLine & _
        "# value 1 activates the silent check feature and 0 deactivates it. The" & vbNewLine & _
        "# default value is 1 for activated silent checks." & vbNewLine & _
        "#" & vbNewLine
    If application_silent_checks_enable = True Then
        ConfigContent = ConfigContent & "application_silent_checks_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_silent_checks_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write the help URL
    ConfigContent = ConfigContent & _
        "# The help url is a string were the default url for access to the application" & vbNewLine & _
        "# online help is saved. This online help repository provides the information" & vbNewLine & _
        "# to handle the software. You are able to provide your own online help" & vbNewLine & _
        "# repository by putting an equivalent html based help on an accessible web" & vbNewLine & _
        "# server and specifying the online help repository url here. It is" & vbNewLine & _
        "# recommended to use the official online help repository at" & vbNewLine & _
        "# http://www.computec.ch/projekte/atk/documentation/help/" & vbNewLine & _
        "#" & vbNewLine & _
        "application_help_url=" & application_help_url & vbNewLine & vbNewLine
    
    'Write the Logs directory
    ConfigContent = ConfigContent & _
        "# The logs directory is as string were the default path name of the log files" & vbNewLine & _
        "# is saved. This data is relevant for the application to write and load the" & vbNewLine & _
        "# logging data. The default value is \logs" & vbNewLine & _
        "#" & vbNewLine & _
        "application_log_directory=" & application_log_directory & vbNewLine & vbNewLine
    
    'Write the logging security level
    ConfigContent = ConfigContent & _
        "# The log security level is an integer variant were the logging level is" & vbNewLine & _
        "# specified. This is the same as like the security level of syslog. Possible" & vbNewLine & _
        "# integer values range from 0 to 7. 7 are very important messages and 7 are" & vbNewLine & _
        "# for debugging only. As more messages are logged, as more ressources are" & vbNewLine & _
        "# used. Very verbose logging may slow down the software a bit. It is" & vbNewLine & _
        "# recommended to set the log level at 5 to get the most import messages." & vbNewLine & _
        "#" & vbNewLine
    If application_log_security_level = 0 Then
        ConfigContent = ConfigContent & "application_log_security_level=0" & vbNewLine & vbNewLine
    ElseIf application_log_security_level = 1 Then
        ConfigContent = ConfigContent & "application_log_security_level=1" & vbNewLine & vbNewLine
    ElseIf application_log_security_level = 2 Then
        ConfigContent = ConfigContent & "application_log_security_level=2" & vbNewLine & vbNewLine
    ElseIf application_log_security_level = 3 Then
        ConfigContent = ConfigContent & "application_log_security_level=3" & vbNewLine & vbNewLine
    ElseIf application_log_security_level = 4 Then
        ConfigContent = ConfigContent & "application_log_security_level=4" & vbNewLine & vbNewLine
    ElseIf application_log_security_level = 5 Then
        ConfigContent = ConfigContent & "application_log_security_level=5" & vbNewLine & vbNewLine
    ElseIf application_log_security_level = 6 Then
        ConfigContent = ConfigContent & "application_log_security_level=6" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_log_security_level=7" & vbNewLine & vbNewLine
    End If
    
    'Write the plugin directory
    ConfigContent = ConfigContent & _
        "# The plugin directory is as string were the default path name of the plugins" & vbNewLine & _
        "# is saved This data is relevant for the application to load the checks and" & vbNewLine & _
        "# to run the access attempts. The default value is \plugins" & vbNewLine & _
        "#" & vbNewLine & _
        "application_plugin_directory=" & application_plugin_directory & vbNewLine & vbNewLine
    
    'Write the plugin download URL
    ConfigContent = ConfigContent & _
        "# The plugin download url is a string were the default url for access to the" & vbNewLine & _
        "# plugin repository is saved. This data is used to fetch the latest plugins" & vbNewLine & _
        "# and install them into the plugin directory. You are able to provide your" & vbNewLine & _
        "# own plugin repository server by putting your exported plugin list on an" & vbNewLine & _
        "# accessible web server and specifying the online plugin repository url here" & vbNewLine & _
        "# It is recommended to use the official plugin repository at" & vbNewLine & _
        "# http://www.computec.ch/projekte/atk/plugins/pluginslist/" & vbNewLine & _
        "#" & vbNewLine & _
        "application_plugin_download_url=" & application_plugin_download_url & vbNewLine & vbNewLine
    
    'Write the plugins external editor
    ConfigContent = ConfigContent & _
        "# The plugin are written as simple ASCII XML text files. This makes it easy" & vbNewLine & _
        "# to edit them with any text editor. The plugins external editor is a string" & vbNewLine & _
        "# where the default application for external plugin editing is saved. It is" & vbNewLine & _
        "# usually possible to define 'notepad.exe' as external editor." & vbNewLine & _
        "#" & vbNewLine & _
        "application_plugin_external_editor=" & application_plugin_external_editor & vbNewLine & vbNewLine
    
    'Write the reportsdirectory
    ConfigContent = ConfigContent & _
        "# The reports directory is a string were the default directory for all" & vbNewLine & _
        "# the reporting is saved. The default directory is \reports" & vbNewLine & _
        "#" & vbNewLine & _
        "application_report_directory=" & application_report_directory & vbNewLine & vbNewLine
    
    'Write the reports open after export
    ConfigContent = ConfigContent & _
        "# It is possible to open a report after save, generation or export. So" & vbNewLine & _
        "# you are able to check the correcteness of the stored data. This" & vbNewLine & _
        "# possibility is saved as a boolean value where 1 stands for activated" & vbNewLine & _
        "# and 0 for de-activated automated opening. The default value is 1 for" & vbNewLine & _
        "# automated opening of stored report data." & vbNewLine & _
        "#" & vbNewLine
        If application_report_open_enable = True Then
            ConfigContent = ConfigContent & "application_report_open_enable=1" & vbNewLine & vbNewLine
        Else
            ConfigContent = ConfigContent & "application_report_open_enable=0" & vbNewLine & vbNewLine
        End If
    
    'Write the responsesdirectory
    ConfigContent = ConfigContent & _
        "# The response directory is a string were the default directory for all" & vbNewLine & _
        "# attack responses is saved. The default directory is \responses" & vbNewLine & _
        "#" & vbNewLine & _
        "ResponsesDirectory=" & application_response_directory & vbNewLine & vbNewLine
    
    'Write if scan should be done if ICMP mapping fails
    ConfigContent = ConfigContent & _
        "# The scan if icmp mapping fails is a boolean variant were the support for" & vbNewLine & _
        "# scans if icmp mapping has been failed is saved. Icmp mapping is used to" & vbNewLine & _
        "# determine the existence and reachability of a target before scanning." & vbNewLine & _
        "# This may prevent attack attempts to non existing nor non reachable hosts." & vbNewLine & _
        "# Under some circumstances the target system does not not is able to react" & vbNewLine & _
        "# on icmp requests so the scan would fail. In this case this function should" & vbNewLine & _
        "# be activated to run a check if icmp mapping has been failed. The feature is" & vbNewLine & _
        "# recommended to be more accurate. The value 1 activates the overriding" & vbNewLine & _
        "# feature and 0 deactivates it. The default value is 0 for stopping scanning" & vbNewLine & _
        "# if icmp mapping fails." & vbNewLine & _
        "#" & vbNewLine
    If application_icmp_mapping_ignore_enable = True Then
        ConfigContent = ConfigContent & "application_icmp_mapping_ignore_enable=1" & vbNewLine & vbNewLine
    Else
        ConfigContent = ConfigContent & "application_icmp_mapping_ignore_enable=0" & vbNewLine & vbNewLine
    End If
    
    'Write the search engine URL
    ConfigContent = ConfigContent & _
        "# The search engine url is a string were the default query url for web" & vbNewLine & _
        "# searches. These are used for further ivestigation on the world wide web" & vbNewLine & _
        "# (e.g. looking for exploits). The software is providing a few well known" & vbNewLine & _
        "# search engine query urls. You are able to define your favorite search" & vbNewLine & _
        "# engine. It is important that the specified search engine allows queries as" & vbNewLine & _
        "# usual HTTP GET requests. In this case you are able to see your query string" & vbNewLine & _
        "# in the URL. Most web searches allow this method." & vbNewLine & _
        "#" & vbNewLine & _
        "application_searchengine_url=" & application_searchengine_url & vbNewLine & vbNewLine
    
    'Write the application_suggestion_directory
    ConfigContent = ConfigContent & _
        "# The suggestions directory is a string were the default directory for all" & vbNewLine & _
        "# the suggestions is saved. This suggestions repository helps new users to" & vbNewLine & _
        "# define the next steps after running an audit attempt. The default" & vbNewLine & _
        "# directory is \suggestions" & vbNewLine & _
        "#" & vbNewLine & _
        "application_suggestion_directory=" & application_suggestion_directory & vbNewLine & vbNewLine
    
    'Write the Target
    ConfigContent = ConfigContent & _
        "# The target is a string were the target for the checking is specified. In" & vbNewLine & _
        "# here host names and ip addresses may be defined. Do not scan ressources" & vbNewLine & _
        "# without premission of the owner of the administrator. The default value is" & vbNewLine & _
        "# the loopback ip address 127.0.0.1 to allow check of the own localhost." & vbNewLine & _
        "#" & vbNewLine & _
        "Target=" & Target & vbNewLine & vbNewLine
    
    'Write the config in the config gile
    On Error Resume Next ' Needed if there are no write permissions
    intFreeFile = FreeFile
    Open strConfigurationFileName For Output As #intFreeFile
        Print #intFreeFile, ConfigContent
    Close
    
    'Show the frame title with configuration file name
    frmConfiguration.Caption = "Configuration - " & application_configuration_filename
    
    'Change frame title so the user can see the next target
    frmMain.Caption = application_name & " - " & Target
End Sub

Public Sub LoadUserName()
    Dim strTemp As String
    
    strTemp = String(255, 0)
    GetUserName strTemp, 255
    system_username = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
End Sub
