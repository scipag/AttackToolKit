Attribute VB_Name = "modComboboxComplete"
Option Explicit

'Function to complete a combobox writing
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) _
    As Long

Public Const CB_FINDSTRING As Long = &H14C

Public Sub ComboAutoComplete(ByRef SourceCtl As VB.ComboBox, _
    ByRef KeyAscii As Integer, ByRef LeftOffPos As Long)
    
    Dim iStart As Long
    Dim sSearchKey As String
    
    With SourceCtl
        'If text entered so far matches item(s) in the list, use autocomplete
        Select Case ChrW$(KeyAscii)
          Case vbBack
            'Let backspace characters process as usual; otherwise try to match text
          Case Else
            If ChrW$(KeyAscii) <> vbBack Then
              .SelText = ChrW$(KeyAscii)
              
              iStart = .SelStart
              
              If LeftOffPos <> 0 Then
                .SelStart = LeftOffPos
                iStart = LeftOffPos
              End If
              
              sSearchKey = CStr(Left$(.Text, iStart))
              .ListIndex = SendMessage(.hwnd, CB_FINDSTRING, -1, _
                  ByVal CStr(Left$(.Text, iStart)))
              
              If .ListIndex = -1 Then
                LeftOffPos = Len(sSearchKey)
              End If
              
              .SelStart = iStart
              .SelLength = Len(.Text)
              LeftOffPos = 0
              
              KeyAscii = 0
            End If
        End Select
    End With
End Sub


