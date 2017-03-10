Attribute VB_Name = "modSuggestionHandling"
Option Explicit

Public suggestion_name As String
Public suggestion_trigger As String
Public suggestion_description As String
Public suggestion_todo As String

Public Sub ReadSuggestionFromFile(Filename As String)
    Dim SuggestionContent As String 'The plugin content itself
    
    'Open and read the plugin file
    On Error Resume Next
    Open App.Path & "\suggestions\" & Filename For Input As 1
        SuggestionContent = Input(LOF(1), #1)
    Close
        
    suggestion_name = ParseAMLTag("name", SuggestionContent)
    suggestion_trigger = ParseAMLTag("trigger", SuggestionContent)
    suggestion_description = ParseAMLTag("description", SuggestionContent)
    suggestion_todo = ParseAMLTag("suggestion", SuggestionContent)
End Sub
