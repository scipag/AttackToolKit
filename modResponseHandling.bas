Attribute VB_Name = "modResponseHandling"
Option Explicit

Public LastResponse As String
Public LastResponseTime As String

Public Sub ClearAllResponseVariables()
    LastResponseTime = vbNullString
    LastResponse = vbNullString
End Sub

Public Sub WriteLastResponseToFile()
    If Not (Dir$(application_response_directory, 16) <> "") Then
        On Error Resume Next    'Prevent errors if the device is write protected
        MkDir (application_response_directory)
    End If
    
    Open application_response_directory & "\" & Target & "-" & plugin_filename & ".txt" For Output As #1
        On Error Resume Next    'Prevent errors if the device is write protected
        Print #1, LastResponse
    Close
End Sub

Public Sub LoadLatestResponse()
    If IsFormVisible("frmAttackResponse") = True Then
        frmAttackResponse.PrepareTabs
        
        frmAttackResponse.lblHost.Caption = Target
        frmAttackResponse.lblPort.Caption = plugin_port
        frmAttackResponse.lblTime.Caption = LastResponseTime
        frmAttackResponse.txtLastResponse.Text = LastResponse
        frmAttackResponse.lblLength.Caption = Len(LastResponse) & _
            " bytes"
    End If
End Sub

Public Function LoadResponseFromFile(ByRef strResponseFileName As String) As String
    'Check the existence of the file
    On Error Resume Next
    If (Dir$(strResponseFileName, 16) <> "") Then
        'Open and read the plugin file
        Open strResponseFileName For Input As 1
            LoadResponseFromFile = Input(LOF(1), #1)
        Close
'    Else
'        Call errPluginDoesNotExist(strFilePath & "\" & strFileName)
    End If
End Function
