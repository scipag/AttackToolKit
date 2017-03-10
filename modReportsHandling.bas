Attribute VB_Name = "modReportsHandling"
Option Explicit

' ************************************************************************************
' * Frame Description                                                                *
' *                                                                                  *
' * In this frame the user is able to configure the report structure.                *
' ************************************************************************************

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 4.0 2004-12-27                                                           *
' * - Splitted the report handling and report templates handling into two modules.   *
' * - Added different string variables to handle report attributes.                  *
' * Version 4.0 2004-12-05                                                           *
' * - Made the first preparations for the new reporting functionality in 4.0.        *
' * Version 2.0 2004-08-24                                                           *
' * - A nasty bug was fixed. The last entry was not computed and missing.            *
' ************************************************************************************

Public report_filecontent As String             'The content of the report file
Public report_filename As String                'The filename of the report file
'Public report_filepath As String                'The filepath of the report file
Public report_filesize As String                'The filesize of the report file

Public report_structure As String               'The loaded report structure

Public Sub WritePluginNameToReportFile(ByRef InputString As String)
    Dim strTargetDirectoryName As String
    
    strTargetDirectoryName = application_report_directory & "\" & Target
    
    If Not (Dir$(strTargetDirectoryName, 16) <> "") Then
        MkDir (strTargetDirectoryName)
    End If
    
    'Write the collected data into the file; the plugin name will be the file name
    Open strTargetDirectoryName & "\" & Target & ".report" For Append As 1
        Print #1, InputString
    Close
End Sub

Public Function LoadReportFromFile(ByRef strReportFileName As String) As String
    Dim intFreeFile As Integer
    
    If (Dir$(strReportFileName, 16) <> "") Then
        'Open and read the report file
        intFreeFile = FreeFile
        Open strReportFileName For Input As #intFreeFile
            LoadReportFromFile = Input(LOF(intFreeFile), #intFreeFile)
        Close
    End If
End Function
