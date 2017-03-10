Attribute VB_Name = "modLocalization"
Option Explicit

Public Language As String * 2

Public Sub LoadEnglishLocalizationfrmMain()

    'Change the language of the menu
    frmMain.mnuFile.Caption = "&File"
        frmMain.mnuExitItem.Caption = "&Exit"
    frmMain.mnuTools.Caption = "&Tools"
        frmMain.mnuConfigurationItem.Caption = "&Configuration"
    frmMain.mnuAddons.Caption = "&Addons"
        frmMain.mnuNslookupItem.Caption = "&Nslookup"
        frmMain.mnuICMPPingItem.Caption = "&ICMP ping"
        frmMain.mnuPortscannerItem.Caption = "&Portscanner"
    frmMain.mnuPlugins.Caption = "&Plugins"
        frmMain.mnuReloadPluginsItem.Caption = "&Reload Plugins"
        frmMain.mnuAttackGeneratorItem.Caption = "Attack Generator"
        frmMain.mnuAttackEditorItem.Caption = "&Attack Editor"
        frmMain.mnuDownloadLatestPluginsItem.Caption = "&Download latest Plugins"
    frmMain.mnuDebugging.Caption = "&Debugging"
        frmMain.mnuAttackResponseItem.Caption = "&Attack Response"
    frmMain.mnuReporting.Caption = "&Reporting"
        frmMain.mnuReportConfigurationItem.Caption = "Report &Configuration"
        frmMain.mnuReportItem.Caption = "&Report"
    frmMain.mnuHelp.Caption = "Help"
        frmMain.mnuProjectWebSiteItem.Caption = "Project &web site"
        frmMain.mnuAboutItem.Caption = "&About"

    frmMain.fraPlugins.Caption = "Plugins"
    
    frmMain.fraPluginOverview.Caption = "Plugin Overview"
    frmMain.lblIDName.Caption = "ID"
    frmMain.lblPortName.Caption = "Port"
    frmMain.lblFamilyName.Caption = "Family"
    frmMain.lblClassName.Caption = "Class"
    frmMain.lblSeverityName.Caption = "Severity"
    frmMain.lblDescriptionName.Caption = "Description"

End Sub

Public Sub LoadEnglishLocalizationfrmAttackEditor()
    frmAttackEditor.Caption = "Attack Editor"
    
    frmAttackEditor.fraAttackData.Caption = "Attack Data"
    
    frmAttackEditor.lblIDName.Caption = "ID"
    frmAttackEditor.lblPluginNameName.Caption = "Name"
    frmAttackEditor.lblProtocolName.Caption = "Protocol"
    frmAttackEditor.lblPortName.Caption = "Port"
    frmAttackEditor.lblPortNote.Caption = "(Note: 0 to 65535)"
    frmAttackEditor.lblRequestName.Caption = "Request"
    frmAttackEditor.lblTriggerName.Caption = "Trigger"
    
    frmAttackEditor.lblNoteName.Caption = "Note"
    frmAttackEditor.lblAttackEditorNote.Caption = "This changes will only be saved until ATK is closed, the plugin is reloaded or a new one is selected."
    
    frmAttackEditor.cmdClose.Caption = "&Close"
End Sub

Public Sub LoadEnglishLocalizationfrmAttackResponse()
    frmAttackResponse.Caption = "Attack Response"
    
    frmAttackResponse.fraLastResponse.Caption = "Last Response"
    frmAttackResponse.lblHostName.Caption = "Host"
    frmAttackResponse.lblPortName.Caption = "Port"
    frmAttackResponse.lblTimeName.Caption = "Time"
    frmAttackResponse.lblLengthName.Caption = "Length"
    
    frmAttackResponse.lblPositionName.Caption = "Position"
    
    frmAttackResponse.cmdSelectTrigger.Caption = "Select &Trigger"
    frmAttackResponse.cmdClose.Caption = "&Close"
End Sub

Public Sub LoadEnglishLocalizationfrmReport()
    frmReport.Caption = "Report"
    
    frmReport.fraActualReport = "Actual Report"

    frmReport.cmdSave.Caption = "&Save"
    frmReport.cmdClose.Caption = "&Close"
End Sub

Public Sub LoadGermanLocalizationfrmMain()

    'Change the language of the menu
    frmMain.mnuFile.Caption = "&Datei"
        frmMain.mnuExitItem.Caption = "&Beenden"
    frmMain.mnuTools.Caption = "&Tools"
        frmMain.mnuConfigurationItem.Caption = "&Konfiguration"
    frmMain.mnuAddons.Caption = "&Addons"
        frmMain.mnuNslookupItem.Caption = "&Nslookup"
        frmMain.mnuICMPPingItem.Caption = "&ICMP ping"
        frmMain.mnuPortscannerItem.Caption = "&Portscanner"
    frmMain.mnuPlugins.Caption = "&Plugins"
        frmMain.mnuReloadPluginsItem.Caption = "&Plugins neu laden"
        frmMain.mnuAttackGeneratorItem.Caption = "Attack Generator"
        frmMain.mnuAttackEditorItem.Caption = "&Attack Editor"
        frmMain.mnuDownloadLatestPluginsItem.Caption = "Die neuesten Plugin &herunterladen"
    frmMain.mnuDebugging.Caption = "&Debugging"
        frmMain.mnuAttackResponseItem.Caption = "&Angriffs Rückantwort"
    frmMain.mnuReporting.Caption = "&Reporting"
        frmMain.mnuReportConfigurationItem.Caption = "Report &Konfiguration"
        frmMain.mnuReportItem.Caption = "&Report"
    frmMain.mnuHelp.Caption = "Hilfe"
        frmMain.mnuProjectWebSiteItem.Caption = "Projekt &Webseite"
        frmMain.mnuAboutItem.Caption = "&About"

    frmMain.fraPlugins.Caption = "Plugins"
    
    frmMain.fraPluginOverview.Caption = "Plugin Übersicht"
    frmMain.lblIDName.Caption = "ID"
    frmMain.lblPortName.Caption = "Port"
    frmMain.lblFamilyName.Caption = "Familie"
    frmMain.lblClassName.Caption = "Klasse"
    frmMain.lblSeverityName.Caption = "Schweregrad"
    frmMain.lblDescriptionName.Caption = "Beschreibung"

End Sub

Public Sub LoadGermanLocalizationfrmAttackEditor()
    frmAttackEditor.Caption = "Attack Editor"
    
    frmAttackEditor.fraAttackData.Caption = "Angriffsdaten"
    
    frmAttackEditor.lblIDName.Caption = "ID"
    frmAttackEditor.lblPluginNameName.Caption = "Name"
    frmAttackEditor.lblProtocolName.Caption = "Protokoll"
    frmAttackEditor.lblPortName.Caption = "Port"
    frmAttackEditor.lblPortNote.Caption = "(Bereich 0 bis 65535)"
    frmAttackEditor.lblRequestName.Caption = "Anfragen"
    frmAttackEditor.lblTriggerName.Caption = "Trigger"
    
    frmAttackEditor.lblNoteName.Caption = "Hinweis"
    frmAttackEditor.lblAttackEditorNote.Caption = "This changes will only be saved until ATK is closed, the plugin is reloaded or a new one is selected."
    frmAttackEditor.lblAttackEditorNote.Caption = "Diese Änderungen werden nur gespeichert, bis das ATK geschlossen, das Plugin neu oder ein anderes geladen wird."
    
    frmAttackEditor.cmdClose.Caption = "&Schliessen"
End Sub

Public Sub LoadGermanLocalizationfrmAttackResponse()
    frmAttackResponse.Caption = "Angriffs Antwort"
    
    frmAttackResponse.fraLastResponse.Caption = "Letzte Reaktion"
    frmAttackResponse.lblHostName.Caption = "Host"
    frmAttackResponse.lblPortName.Caption = "Port"
    frmAttackResponse.lblTimeName.Caption = "Zeit"
    frmAttackResponse.lblLengthName.Caption = "Länge"
    
    frmAttackResponse.lblPositionName.Caption = "Position"
    
    frmAttackResponse.cmdSelectTrigger.Caption = "&Trigger"
    frmAttackResponse.cmdClose.Caption = "&Schliessen"
End Sub

Public Sub LoadGermanLocalizationfrmReport()
    frmReport.Caption = "Report"
    
    frmReport.fraActualReport = "Aktueller Report"

    frmReport.cmdSave.Caption = "S&peichern"
    frmReport.cmdClose.Caption = "&Schliessen"
End Sub

Public Sub LoadEnglishLocalizationfrmNslookup()
    frmNslookup.fraHost.Caption = "Host"
    
    frmNslookup.fraResult.Caption = "Result"
    frmNslookup.lblIPAddressName.Caption = "IP address"
    frmNslookup.lblHostNameName.Caption = "Host name"
    
    'Not done yet!
    
End Sub
