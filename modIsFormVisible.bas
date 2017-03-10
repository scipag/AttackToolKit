Attribute VB_Name = "modIsFormVisible"
Option Explicit

Public Function IsFormVisible(ByRef FormName As String) As Boolean
    Dim frm As Form
    
    'Initialize the status as form is not visible
    IsFormVisible = False
    
    For Each frm In Forms
        If frm.Name = FormName Then
            IsFormVisible = True
            Exit For
        End If
    Next frm
End Function
