Attribute VB_Name = "modErrorHandling"
Option Explicit

Public Sub errPluginsDirectoryEmpty()
    WriteLogEntry "In " & application_plugin_directory & " no plugins could be found.", 5
    
    If MsgBox("No plugins could be loaded because the default plugin directory" & vbNewLine & _
        application_plugin_directory & vbNewLine & _
        "is empty! No predefined checks are possible at the moment." & vbNewLine & _
        "Please check the plugins directory configuration." & vbNewLine & vbNewLine & _
        "Would you like to start the AutoUpdate to download the latest ATK plugins?", _
        vbYesNo + vbInformation, "Attack Tool Kit load plugins error") = vbYes Then
        
        'Open the AutoUpdate
        Call frmMain.mnuPluginsDownloadTheLatestPluginsItem_Click
    Else
        WriteLogEntry "Opening AutoUpdate to get the latest plugins has been manually aborded.", 4
    End If
End Sub

Public Sub errPluginsDirectoryNotExist()
    'Error message if the plugin directory does not exists
    WriteLogEntry "The plugin directory " & application_plugin_directory & " does not exists.", 3
    
    If MsgBox("No plugins could be loaded because the default plugin directory" & vbNewLine & _
        application_plugin_directory & vbNewLine & _
        "does not exists! No predefined checks are possible at the moment." & vbNewLine & _
        "Please check the plugins directory configuration." & vbNewLine & vbNewLine & _
        "Would you like to create the plugins directory " & vbNewLine & _
        application_plugin_directory & "?", vbYesNo + vbInformation, "Attack Tool Kit load plugins error") = vbYes Then
        
        'Make the plugin directory
        On Error Resume Next
        MkDir (application_plugin_directory)
        WriteLogEntry "Plugins directory " & application_plugin_directory & " created.", 6
    Else
        WriteLogEntry "Creating the plugin directory " & application_plugin_directory & _
            " has been manually aborded.", 3
    End If
End Sub

Public Sub errLogDirectoryNotExist()
    'Developer note: We cannot use the logging feature in this procedure because we would
    'get a nasty recursive routine without an exit.
    
    If application_log_directory_enable = False Then
        If MsgBox("No file logging could be done because the default logs directory" & vbNewLine & _
            application_log_directory & vbNewLine & _
            "does not exists! No additionall debugging was possible until now." & vbNewLine & vbNewLine & _
            "Would you like to create the logs directory " & vbNewLine & _
            application_log_directory & "?", vbYesNo + vbInformation, "Attack Tool Kit precheck logs warning") = vbYes Then
        
            'Make the logs directory
            On Error Resume Next 'Skip the mkdir command if there are no write permissions
            MkDir (application_log_directory)
            WriteLogEntry "Logs directory " & application_log_directory & " created.", 6
        Else
            'Set the value that no log directory is wished. All further error messages
            'in this field will be ignored and not shown.
            application_log_directory_enable = True
        End If
    End If
End Sub

Public Sub errLogDirectoryEmpty()
    WriteLogEntry "In " & application_log_directory & " no log files could be found.", 4
    
    If MsgBox("No log files could be found because the default log directory" & vbNewLine & _
        application_log_directory & vbNewLine & _
        "is empty! No further application analysis possible at the moment." & vbNewLine & _
        "Please check the log directory configuration." & vbNewLine & vbNewLine & _
        "Would you like to load a specific log file to start a log analysis?", _
        vbYesNo + vbInformation, "Attack Tool Kit load log error") = vbYes Then
        
        'Open the AutoUpdate
        Call frmLog.mnuFileOpenItem_Click
    Else
        WriteLogEntry "Opening a specific log file has been manually aborded.", 4
    End If
End Sub

Public Sub errSuggestionsDirectoryNotExist()
    'Error message if the plugin directory does not exists
    WriteLogEntry "The suggestions directory " & application_suggestion_directory & " does not exist.", 3
    
    If MsgBox("No suggestions could be loaded because the default suggestions directory" & vbNewLine & _
        application_suggestion_directory & vbNewLine & _
        "does not exists! No additionall suggestions are possible at the moment." & vbNewLine & _
        "Please check the suggestions directory configuration." & vbNewLine & vbNewLine & _
        "Would you like to create the suggestions directory " & vbNewLine & _
        application_suggestion_directory & "?", vbYesNo + vbInformation, "Attack Tool Kit suggestions error") = vbYes Then
            
        'Make the suggestions directory
        On Error Resume Next 'Skip the mkdir command if there are no write permissions
        MkDir (application_suggestion_directory)
        WriteLogEntry "Suggestions directory " & application_suggestion_directory & " created.", 6
    Else
        WriteLogEntry "Creating the suggestions directory " & application_suggestion_directory & _
            " has been manually aborded.", 4
    End If
End Sub

'Public Sub errReportDirectoryNotExist()
'    'Error message if the plugin directory does not exists
'    WriteLogEntry "The reports directory " & application_report_directory & " does not exist.", 3
'
'    If MsgBox("No reports could be cached because the default reports directory" & vbNewline & _
'        application_report_directory & vbNewline & _
'        "does not exists! No further analysis was possible until now." & vbNewline & vbNewline & _
'        "Would you like to create the suggestions directory " & vbNewline & _
'        application_report_directory & "?", vbYesNo + vbInformation, "Attack Tool Kit report warning") = vbYes Then
'
'        'Make the suggestions directory
'        On Error Resume Next 'Skip the mkdir command if there are no write permissions
'        MkDir (application_report_directory)
'    End If
'End Sub

Public Sub errPluginDoesNotExist(ByRef strPluginFileName As String)
    WriteLogEntry "The plugin " & strPluginFileName & " does not exist anymore.", 2
    
    If MsgBox("The specified plugin " & strPluginFileName & vbNewLine & _
        "does not exist anymore. It may be deleted since the last access. You are not able to use the plugin at the moment." & vbNewLine & vbNewLine & _
        "Please check the plugins directory configuration or run the AutoUpdate to download the latest plugins." & vbNewLine & vbNewLine & _
        "Would you like to start the AutoUpdate to re-initialize your local ATK plugins repository?", _
        vbYesNo + vbInformation, "Attack Tool Kit load plugin error") = vbYes Then
        
        'Open the AutoUpdate
        Call frmMain.mnuPluginsDownloadTheLatestPluginsItem_Click
    Else
        WriteLogEntry "Opening AutoUpdate to get the latest plugins as been manually aborded.", 4
    End If
End Sub

Public Sub errPluginDataMissing(ByRef strMissingDataName As String, ByRef strPluginFileName As String, ByRef intPluginID As String)
    'Write a log entry about the error
    WriteLogEntry "Important attack data " & strMissingDataName & " is missing. Check aborded.", 1
    
    'Show the error message
    If MsgBox("Important attack data " & strMissingDataName & " is missing." & vbNewLine & vbNewLine & _
        "You will not be able to run the plugin " & intPluginID & vbNewLine & _
        " (" & strPluginFileName & ") correctly." & vbNewLine & vbNewLine & _
        "Would you like to open the Attack Editor to check the error manually?", _
        vbYesNo + vbInformation, "Attack Tool Kit plugin data error") = vbYes Then
    
        'Show the attack editor to eliminate the check error
        frmAttackEditor.Visible = True
    Else
        WriteLogEntry "Opening the Attack Editor to check the missing data manually has been manually aborded.", 3
    End If
End Sub

Public Sub errTargetWrongSpecification()
    'Error message if the has been specified in a wrong way
    WriteLogEntry "The target has been specified in a wrong way.", 4
    
    MsgBox "You have specified the target in a wrong way that is not supported by this version" & vbNewLine & _
        "of the Attack Tool Kit (ATK)." & vbNewLine & vbNewLine & _
        "You can specify host names (e.g. www.computec.ch) or IP addresses (e.g." & vbNewLine & _
        "192.168.0.1) only. Your input has been re-written to prevent run-time errors." & vbNewLine & vbNewLine & _
        "Please check the new target definition to get the wanted match for your" & vbNewLine & _
        "attack.", vbOKOnly + vbInformation, "Attack Tool Kit target error"
End Sub

Public Sub errPluginExternalEditorMissing()
    'Write a log entry about the error
    WriteLogEntry "Could not open the plugin " & plugin_filename & " with the selected external editor " & application_plugin_external_editor, 4
    
    If MsgBox("It was not possible to open the selected plugin" & vbNewLine & _
        application_plugin_directory & "\" & plugin_filename & vbNewLine & _
        "with the external editor " & application_plugin_external_editor & "." & vbNewLine & vbNewLine & _
        "Please check the in the configuration specified external editor for plugins." & vbNewLine & vbNewLine & _
        "Or would you like to open the default editor notepad.exe for editing the plugin?", _
        vbYesNo + vbInformation, "Attack Tool Kit external plugin error") = vbYes Then

        Call ShellExecute(frmMain.hwnd, "Open", "notepad.exe", plugin_filename, application_plugin_directory, 1)
    End If
End Sub

'Private Sub errPluginReadError(ByRef Filename As String, ByRef Position As String)
'    'Write error message if something went wrong during parsing of the plugin file
'    MsgBox ("Could not find the data field '" & Position & "'" & vbNewLine & _
'        "in the plugin '" & Filename & "'." & vbNewLine & vbNewLine & _
'        "The plugin seems to be broken and can't be used." & vbNewLine & _
'        "Please check this manually."), _
'        vbInformation, "Attack Tool Kit Plugin parsing error"
'End Sub

