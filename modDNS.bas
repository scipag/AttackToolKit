Attribute VB_Name = "modDNS"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ************************************************************************************
' * Developement History                                                             *
' *                                                                                  *
' * Version 3.1 2004-11-20                                                           *
' * - Fixed a bug in GetIPFromHostName(). The new code is done by                    *
' *   lijian1976 at hotmail.com                                                      *
' ************************************************************************************

Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSADATA) As Long
  
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function inet_addr Lib "wsock32" _
  (ByVal s As String) As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
Public Const IP_SUCCESS As Long = 0
Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1

Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
Private Const AF_INET As Long = 2

Public Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32" _
  (ByVal Hostname As String) As Long
  
Private Declare Function lstrlenA Lib "kernel32" _
  (lpString As Any) As Long

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim success As Long
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function

Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub

Public Function GetIPFromHostName(ByVal sHostName As String) As String

  'converts a host name to an IP address.

   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   hostent_addr = gethostbyname(sHostName & vbNullChar)

   If hostent_addr <> 0 Then

      RtlMoveMemory host, hostent_addr, LenB(host)
      RtlMoveMemory hostip_addr, host.hAddrList, 4

      ReDim temp_ip_address(1 To host.hLen)
      RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLen
  
      For i = 1 To host.hLen
          ip_address = ip_address & temp_ip_address(i) & "."
      Next
      ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

      GetIPFromHostName = ip_address

   End If
   
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetHostNameFromIP(ByVal sAddress As String) As String

   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
   
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               GetHostNameFromIP = sAddress
            End If
         
         'Else: MsgBox "Call to gethostbyaddr failed."
         End If 'If ptrHosent
      
      SocketsCleanup
      
      'Else: MsgBox "String passed is an invalid IP."
      End If 'If hAddress
   
   'Else: MsgBox "Sockets failed to initialize."
   End If  'If SocketsInitialize
      
End Function
