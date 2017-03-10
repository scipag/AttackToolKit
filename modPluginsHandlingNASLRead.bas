Attribute VB_Name = "modPluginsHandlingNASLRead"
Option Explicit

Public Sub ParseNASLPlugin(ByRef strNASLPluginContent As String)
    Dim TempArray() As String   'A temporary array for the splitting and parsing
    
    'Replace the problematic whitespaces in the plugin content
    strNASLPluginContent = Replace$(strNASLPluginContent, vbLf, vbNewLine, , , vbBinaryCompare)
    strNASLPluginContent = Replace$(strNASLPluginContent, vbCr, vbNewLine, , , vbBinaryCompare)
    strNASLPluginContent = Replace$(strNASLPluginContent, vbNewLine & vbNewLine, vbNewLine, , , vbBinaryCompare)
    strNASLPluginContent = Replace$(strNASLPluginContent, vbTab, "", , , vbBinaryCompare)

    'Clear the values from the last plugin to prevent misunderstandings
    Call ClearAllPluginVariables

    On Error Resume Next

    'Get the data fields and write them into the public variables
    TempArray = Split(strNASLPluginContent, "script_id(")
    TempArray = Split(TempArray(1), ");")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        plugin_id = "N" & Val(TempArray(0))
        source_nessus_id = Val(TempArray(0))
    End If

    TempArray = Split(strNASLPluginContent, "script_name(english:" & ChrW$(34))
    TempArray = Split(TempArray(1), ChrW$(34) & ");")
    If LenB(TempArray(0)) = LenB(strNASLPluginContent) Then
        TempArray = Split(strNASLPluginContent, "name[" & ChrW$(34) & "english" & ChrW$(34) & "] = " & ChrW$(34))
        TempArray = Split(TempArray(1), ChrW$(34) & ";")
        If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
            plugin_name = TempArray(0)
        Else
            plugin_name = plugin_filename
        End If
    Else
        plugin_name = TempArray(0)
    End If

    'Plugin version
    TempArray = Split(strNASLPluginContent, "(" & ChrW$(34) & "$Revision: ")
    TempArray = Split(TempArray(1), " $" & ChrW$(34) & ")")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        plugin_version = TempArray(0)
    Else
        TempArray = Split(Replace(strNASLPluginContent, " ", vbNullString, , , vbBinaryCompare), "script_version(" & ChrW$(34))
        TempArray = Split(TempArray(1), ChrW$(34) & ");")
        If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
            plugin_version = TempArray(0)
        End If
    End If

    'Description
    TempArray = Split(strNASLPluginContent, "desc[" & ChrW$(34) & "english" & ChrW$(34) & "] = " & ChrW$(34))
    TempArray = Split(TempArray(1), "Solution")
    If LenB(TempArray(0)) = LenB(strNASLPluginContent) Then
        TempArray = Split(strNASLPluginContent, "edesc= " & ChrW$(34) & vbNewLine)
        TempArray = Split(TempArray(1), vbNewLine & "Solution")
        
        If LenB(TempArray(0)) = LenB(strNASLPluginContent) Then
            TempArray = Split(strNASLPluginContent, "script_description(english:string(" & ChrW$(34))
            TempArray = Split(TempArray(1), "Risk")
        End If
    End If
    bug_description = Replace$(TempArray(0), vbNewLine, " ")
    bug_description = Replace$(bug_description, ChrW$(10), " ")
    bug_description = Trim$(bug_description)

    'Solution
    TempArray = Split(strNASLPluginContent, "Solution")
    TempArray = Split(TempArray(1), "Risk")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        TempArray = Split(TempArray(0), ": ")
        TempArray = Split(TempArray(1), ";")
        bug_solution = Trim(Replace(TempArray(0), vbNewLine, " "))
    End If

    'The risk
    TempArray = Split(strNASLPluginContent, "actor")
    TempArray = Split(TempArray(1), ChrW$(34) & ";")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        If InStrB(1, LCase$(TempArray(0)), "low", vbBinaryCompare) Then
            bug_severity = "Low"
        ElseIf InStrB(1, LCase$(TempArray(0)), "medium", vbBinaryCompare) Then
            bug_severity = "Medium"
        ElseIf InStrB(1, LCase$(TempArray(0)), "high", vbBinaryCompare) Then
            bug_severity = "High"
        ElseIf InStrB(1, LCase$(TempArray(0)), "critical", vbBinaryCompare) Then
            bug_severity = "Critical"
        Else
            Dim j As Integer
            
            For j = 1 To LenB(TempArray(0))
                If Mid$(TempArray(0), j, 1) Like "[A-Za-z]" Then
                    bug_severity = bug_severity & Mid$(TempArray(0), j, 1)
                ElseIf j > 3 Then
                    Exit For
                End If
            Next j
        End If
        bug_nessus_risk = bug_severity
    End If

    'Plugin family
    TempArray = Split(strNASLPluginContent, "family[" & ChrW$(34) & "english" & ChrW$(34) & "] = " & ChrW$(34))
    TempArray = Split(TempArray(1), ChrW$(34) & ";")
    If LenB(TempArray(0)) = LenB(strNASLPluginContent) Then
        TempArray = Split(TempArray(0), "script_family(english:" & ChrW$(34))
        TempArray = Split(TempArray(1), ChrW$(34) & ");")
        If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
            plugin_family = TempArray(0)
        End If
    Else
        plugin_family = TempArray(0)
    End If

    'Vulnerability Class
    If InStrB(1, LCase$(plugin_name), "buffer overflow", vbBinaryCompare) Then
        bug_vulnerability_class = "Buffer Overflow"
    ElseIf InStrB(1, LCase$(plugin_name), "bufferoverflow", vbBinaryCompare) Then
        bug_vulnerability_class = "Buffer Overflow"
    ElseIf InStrB(1, LCase$(plugin_name), "configuration", vbBinaryCompare) Then
        bug_vulnerability_class = "Configuration"
    ElseIf InStrB(1, LCase$(plugin_name), "cross site scripting", vbBinaryCompare) Then
        bug_vulnerability_class = "Cross Site Scripting"
    ElseIf InStrB(1, LCase$(plugin_name), "css", vbBinaryCompare) Then
        bug_vulnerability_class = "Cross Site Scripting"
    ElseIf InStrB(1, LCase$(plugin_name), "xss", vbBinaryCompare) Then
        bug_vulnerability_class = "Cross Site Scripting"
    ElseIf InStrB(1, LCase$(plugin_name), "html injection", vbBinaryCompare) Then
        bug_vulnerability_class = "Cross Site Scripting"
    ElseIf InStrB(1, LCase$(plugin_name), "cross domain scripting", vbBinaryCompare) Then
        bug_vulnerability_class = "Cross Domain Scripting"
    ElseIf InStrB(1, LCase$(plugin_name), "denial of service", vbBinaryCompare) Then
        bug_vulnerability_class = "Denial Of Service"
    ElseIf InStrB(1, LCase$(plugin_name), "evasion", vbBinaryCompare) Then
        bug_vulnerability_class = "Evasion"
    ElseIf InStrB(1, LCase$(plugin_name), "circumvent", vbBinaryCompare) Then
        bug_vulnerability_class = "Evasion"
    ElseIf InStrB(1, LCase$(plugin_name), "format string", vbBinaryCompare) Then
        bug_vulnerability_class = "Format String"
    ElseIf InStrB(1, LCase$(plugin_name), "sql injection", vbBinaryCompare) Then
        bug_vulnerability_class = "SQL Injection"
    ElseIf InStrB(1, LCase$(plugin_name), "symlink", vbBinaryCompare) Then
        bug_vulnerability_class = "Symlink"
    ElseIf InStrB(1, LCase$(plugin_name), "authentication", vbBinaryCompare) Then
        bug_vulnerability_class = "Weak Authentication"
    ElseIf InStrB(1, LCase$(plugin_name), "encryption", vbBinaryCompare) Then
        bug_vulnerability_class = "Weak Encryption"
    Else
        bug_vulnerability_class = "Unknown"
    End If

    'Port
    TempArray = Split(strNASLPluginContent, "require_ports(", , vbBinaryCompare)
    TempArray = Split(TempArray(1), ");", , vbBinaryCompare)
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        TempArray = Split(TempArray(0), ", ", , vbBinaryCompare)
        TempArray = Split(TempArray(1), ");", , vbBinaryCompare)
        If InStr(1, TempArray(0), "/", vbBinaryCompare) Then
            TempArray = Split(strNASLPluginContent, "require_ports(", , vbBinaryCompare)
            TempArray = Split(TempArray(1), ", ", , vbBinaryCompare)
            plugin_port = Val(TempArray(0))
        Else
            plugin_port = Val(TempArray(0))
        End If
    Else
        TempArray = Split(strNASLPluginContent, "ort(default:", , vbBinaryCompare)
        TempArray = Split(TempArray(1), ");", , vbBinaryCompare)
        If LenB(TempArray(0)) = LenB(strNASLPluginContent) Then
            TempArray = Split(TempArray(0), ",")
            plugin_port = Replace(Val(TempArray(0)), "," Or vbNewLine, "", , , vbBinaryCompare)
        Else
            plugin_port = "80"
        End If
    End If
    
    'Usual open tcp socket
    If InStr(1, strNASLPluginContent, "open_sock_tcp", vbBinaryCompare) Or _
        InStr(1, strNASLPluginContent, "http_get", vbBinaryCompare) Then
        plugin_protocol = "tcp"
    ElseIf InStr(1, strNASLPluginContent, "open_sock_udp", vbBinaryCompare) Then
        plugin_protocol = "udp"
    ElseIf InStr(1, strNASLPluginContent, "forge_icmp_packet", vbBinaryCompare) Then
        plugin_protocol = "icmp"
    Else
        plugin_protocol = "unknown"
    End If
    
    'Plugin request
    TempArray = Split(strNASLPluginContent, "http_get(item:" & ChrW$(34))
    TempArray = Split(TempArray(1), ChrW$(34) & ", port:")
    If LenB(TempArray(0)) = LenB(strNASLPluginContent) Then
        plugin_procedure_detection = "open|sleep|close|pattern_exists"
    Else
        plugin_procedure_detection = "open|send " & TempArray(0) & " HTTP/1.0\n\n|sleep|close|pattern_exists HTTP/#.# ### *"
    End If

    'Copyright information
    TempArray = Split(strNASLPluginContent, "script_copyright(english:" & ChrW$(34))
    TempArray = Split(TempArray(1), ChrW$(34) & ");")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        plugin_comment = TempArray(0)
    Else
        plugin_comment = "This script may be copyrighted by the Nessus project or Tenable Network Security."
    End If

    'bug_published_by
    TempArray = Split(strNASLPluginContent, "From: ")
    TempArray = Split(TempArray(1), vbNewLine)
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        bug_published_name = Replace(TempArray(0), ChrW$(34), "")
    Else
        TempArray = Split(strNASLPluginContent, "Ref: ")
        TempArray = Split(TempArray(1), vbNewLine)
        If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
            bug_published_name = Replace(TempArray(0), ChrW$(34), "")
        End If
    End If
    
    'Plugin author
    TempArray = Split(strNASLPluginContent, "Author: ")
    TempArray = Split(TempArray(1), vbNewLine)
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        plugin_created_name = Replace(TempArray(0), ChrW$(34), "")
    End If
    
    'Pattern matching
    TempArray = Split(strNASLPluginContent, "if(" & ChrW$(34))
    TempArray = Split(TempArray(1), ChrW$(34) & " >< res")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        plugin_procedure_detection = plugin_procedure_detection & " " & Replace(TempArray(0), "^", vbNullString, , , vbBinaryCompare) & "*"
        plugin_detection_accuracy = "70"
    Else
        TempArray = Split(strNASLPluginContent, "egrep(pattern:")
        TempArray = Split(TempArray(1), ", string:")
        If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
            plugin_procedure_detection = plugin_procedure_detection & " " & Replace(TempArray(0), "^", vbNullString, , , vbBinaryCompare) & "*"
            plugin_detection_accuracy = "70"
        Else
            plugin_detection_accuracy = "20"
        End If
    End If

    'CVE ID
    TempArray = Split(strNASLPluginContent, "script_cve_id(")
    TempArray = Split(TempArray(1), ");")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        source_cve = Replace(TempArray(0), ChrW$(34), "")
    End If

    'Bugtraq ID
    TempArray = Split(strNASLPluginContent, "script_bugtraq_id(")
    TempArray = Split(TempArray(1), ");")
    If LenB(TempArray(0)) <> LenB(strNASLPluginContent) Then
        source_securityfocus_bid = Val(Replace(TempArray(0), ChrW$(34), ""))
    End If
    
    'bug_checking_tool
    bug_check_tool = "Nessus can check this flaw with the plugin " & source_nessus_id & " (" & plugin_name & ")."

    bug_exploit_availability = "Maybe"

    'bug_remote
    bug_remote = "Yes"
    bug_local = "Maybe"
    
    source_literature = "Hacking Exposed: Network Security Secrets & Solutions, " & _
        "Stuart McClure, Joel Scambray and George Kurtz, " & _
        "February 25, 2003, 4th Edition, McGraw-Hill Osborne Media, ISBN 0072227427"
    source_misc = "http://www.computec.ch"
End Sub

