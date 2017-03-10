Attribute VB_Name = "modListViewColumns"
Option Explicit
 
Private Const LVM_SETCOLUMNWIDTH As Integer = &H1000 + 30
Private Const LVSCW_AUTOSIZE As Integer = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Integer = -2

'Set the list view columns on the needed width
Public Sub LVColumnWidth(oListView As MSComctlLib.ListView, _
    Optional AccountForHeaders As Boolean = False)
    
    Dim col As Long
    Dim LParm As Long
      
    On Error GoTo error
    If AccountForHeaders Then
        LParm = LVSCW_AUTOSIZE_USEHEADER
    Else
        LParm = LVSCW_AUTOSIZE
    End If
      
    For col = 0 To oListView.ColumnHeaders.Count - 1
        SendMessage oListView.hwnd, LVM_SETCOLUMNWIDTH, _
            col, ByVal LParm
    Next col
error:
End Sub

Public Sub ListViewColumnReorder(ByRef lsvListViewName As ListView, ByRef ColumnHeader As MSComctlLib.ColumnHeader)
    WriteLogEntry "Reorder the columns in the selected listview.", 6
    
    If lsvListViewName.SortKey = ColumnHeader.Index - 1 Then
        If lsvListViewName.SortOrder = lvwAscending Then
            lsvListViewName.SortKey = ColumnHeader.Index - 1
            lsvListViewName.SortOrder = lvwDescending
        Else
            lsvListViewName.SortKey = ColumnHeader.Index - 1
            lsvListViewName.SortOrder = lvwAscending
        End If
    Else
        lsvListViewName.SortKey = ColumnHeader.Index - 1
        lsvListViewName.SortOrder = lvwAscending
    End If
End Sub
