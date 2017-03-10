Attribute VB_Name = "modBrowserCall"
Option Explicit

'Declare the function for the browser call
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
    lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Sub OpenProjectWebsite()
    'Load the project web site
    Call ShellExecute(frmMain.hwnd, "Open", application_website_url, "", App.Path, 1)
End Sub

Public Sub OpenOnlineHelp(Optional ByRef strSubDirectory As String)
    Dim strFullOnlineHelpURL As String
    
    strFullOnlineHelpURL = application_help_url & strSubDirectory
    
    If LenB(strFullOnlineHelpURL) = 0 Then
        strFullOnlineHelpURL = application_website_url
    End If
    
    'Load the online help
    WriteLogEntry "Opening the online help URL " & strFullOnlineHelpURL & " ...", 6
    Call ShellExecute(frmMain.hwnd, "Open", strFullOnlineHelpURL, "", App.Path, 1)
End Sub

Public Sub OpenOnlineSearch(Optional ByRef strSearchString As String)
    Dim strFullSearchURL As String
    
    strFullSearchURL = application_searchengine_url & strSearchString
    
    If LenB(strFullSearchURL) = 0 Then
        strFullSearchURL = "http://www.google.com"
    End If
    
    'Load the online search
    WriteLogEntry "Opening the search engine URL " & strFullSearchURL & " ...", 6
    Call ShellExecute(frmMain.hwnd, "Open", strFullSearchURL, "", App.Path, 1)
End Sub

Public Sub OpenSelectedTextIfItIsURL(ByRef strSelectedText As String)
    If LenB(strSelectedText) Then
        If Mid$(strSelectedText, 1, 7) = "http://" Then
            Call ShellExecute(frmMain.hwnd, "Open", strSelectedText, "", App.Path, 1)
        End If
    End If
End Sub
