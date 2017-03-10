Attribute VB_Name = "modLogHandling"
Option Explicit

Public application_log_directory_enable As Boolean  'Saves if the user want to have a log directory or not

Public Sub WriteLogEntry(ByRef strMessageText As String, ByRef intSecurityLevel As Integer)
    'Keep the user up to date in the statusbar
    frmMain.StatusBar.Panels(1).Text = strMessageText
    
    If application_log_enable = True Then
        If application_log_security_level >= intSecurityLevel Then
            On Error Resume Next    'Needed bevause I can't detect read-only files at the moment.
            
            'Check if the log directory exists and prepare for the writint
            If Not (Dir$(application_log_directory, 16) <> "") Then
                Call errLogDirectoryNotExist
            End If
        
            'And write the new entry in the log file
            Open application_log_directory & "\log-" & GetTodaysDate(".") & ".log" For Append As #1
                'Write the log entry in the log file
                Print #1, GetTodaysDate("/") & ";" & GetActualTime(":") & ";" & strMessageText & ";" & intSecurityLevel
            Close
        End If
        
        'Add the data in real-time in the log frame if it is loaded
        If IsFormVisible("frmLog") = True Then
            Dim List As ListItem        'Needed for the listview handling
            
            Set List = frmLog.lsvLog.ListItems.Add(, , GetTodaysDate("/"))
                List.SubItems(1) = GetActualTime(":")
                List.SubItems(2) = strMessageText
                List.SubItems(3) = intSecurityLevel
            
            'Set the right column width
            LVColumnWidth frmLog.lsvLog
        End If
    End If
End Sub
