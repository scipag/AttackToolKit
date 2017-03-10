Attribute VB_Name = "modReportTemplatesHandling"
Option Explicit

Public report_template_filecontent As String    'The content of the report template
Public report_template_filename As String       'The filename of the report template
Public report_template_filepath As String       'The filepath of the report template
Public report_template_filesize As String       'The filesize of the report template

Public Sub LoadDefaultReportStructure()
    'Load the default data
    report_structure = _
    "plugin_id" & vbNewLine & _
    "plugin_name" & vbNewLine & _
    "plugin_protocol" & vbNewLine & _
    "plugin_port" & vbNewLine & _
    "bug_severity" & vbNewLine & _
    "bug_advisory" & vbNewLine & _
    "bug_affected" & vbNewLine & _
    "bug_not_affected" & vbNewLine & _
    "bug_vulnerability_class" & vbNewLine & _
    "bug_exploit_url" & vbNewLine
    
    report_structure = report_structure & _
    "<br>" & vbNewLine & _
    "bug_description" & vbNewLine & _
    "<br>" & vbNewLine & _
    "bug_response" & vbNewLine & _
    "<br>" & vbNewLine & _
    "bug_solution" & vbNewLine & _
    "<br>" & vbNewLine & _
    "source_cve" & vbNewLine & _
    "source_securityfocus_bid" & vbNewLine & _
    "source_osvdb_id" & vbNewLine & _
    "source_nessus_id"
    
    'Generate the attribute data of the default report
    report_template_filecontent = report_structure
    report_template_filename = App.EXEName & " (internal default)"
    report_template_filepath = App.Path
    report_template_filesize = Len(report_template_filecontent)
End Sub

Public Sub PrepareReportStructure()
    Dim i As Integer
    Dim VulnerabilityListCount As Integer

    'Delete the old report template content
    report_structure = vbNullString

    VulnerabilityListCount = frmReportConfiguration.lstVulnerabilityReport.ListCount - 1

    For i = 0 To VulnerabilityListCount
        report_structure = report_structure & _
            frmReportConfiguration.lstVulnerabilityReport.List(i) & vbNewLine
    Next i
End Sub

Public Function LoadReportTemplateFromFile(ByRef strFileName As String, _
                                            ByRef strFilePath As String) As String
                                            
    Dim strTemplateFullFileName As String 'The full path and name of the plugin file

    'This is just a workaround because the Open dialog can't split file name and path
    If InStrB(1, strFileName, "\", vbBinaryCompare) Then
        strTemplateFullFileName = strFileName
    Else
        strTemplateFullFileName = strFilePath & "\" & strFileName
    End If

    'Check the existence of the file
    On Error Resume Next
    If (Dir$(strTemplateFullFileName, 16) <> "") Then
        'Flush the old plugin content before loading new data
        report_template_filecontent = vbNullString
        
        'Open and read the plugin file
        Open strTemplateFullFileName For Input As 1
            report_template_filecontent = Input(LOF(1), #1)
        Close
        
        report_template_filesize = Len(report_template_filecontent)
        report_template_filepath = strFilePath
        report_template_filename = strFileName
        
        LoadReportTemplateFromFile = report_template_filecontent
    End If
End Function

Public Sub WriteReportTemplateToFile(ByRef strTemplateFileName As String)
    'Write the collected data into the file
    On Error Resume Next
    Open strTemplateFileName For Output As 1
        Print #1, report_structure
    Close
    
    report_template_filename = strTemplateFileName
    report_template_filesize = Len(report_structure)
    report_template_filecontent = report_structure
End Sub
