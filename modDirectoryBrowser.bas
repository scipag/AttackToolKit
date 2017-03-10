Attribute VB_Name = "modDirectoryBrowser"
Option Explicit

Private Type BROWSEINFO
  hwndOwner       As Long
  pIDLRoot        As Long
  pszDisplayName  As Long
  lpszTitle       As String
  ulFlags         As Long
  lpfnCallback    As Long
  lParam          As Long
  iImage          As Long
End Type
    
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)

Private Const MAX_PATH  As Integer = 260
Private Const BIF_RETURNONLYFSDIRS As String = &H1&

Public Function BrowseForFolder(Optional Parent As Variant, _
                                Optional Title As Variant) As String
                                
  Dim tBI         As BROWSEINFO
  Dim lhWndParent As Long
  Dim lngPIDL     As Long
  Dim strPath     As String
  
  If IsMissing(Title) Then Title = "Please choose a directory"
  If IsMissing(Parent) = False Then lhWndParent = Parent.hwnd
  
  With tBI
    .hwndOwner = lhWndParent
    .lpszTitle = Title
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  
  lngPIDL = SHBrowseForFolder(tBI)
  
  If (lngPIDL <> 0) Then
    strPath = Space$(MAX_PATH)
    SHGetPathFromIDList lngPIDL, strPath
    
    strPath = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
    
    CoTaskMemFree lngPIDL
  Else
    strPath = ""
  End If
  
  BrowseForFolder = strPath
End Function
