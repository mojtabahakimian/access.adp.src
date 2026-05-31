Attribute VB_Name = "POS"
Option Compare Database
'reference
'http://msdn.microsoft.com/en-us/library/windows/desktop/ms740673(v=vs.85).aspx
Const INVALID_SOCKET = -1
Const WSADESCRIPTION_LEN = 256

Enum AF
  AF_UNSPEC = 0
  AF_INET = 2
  AF_IPX = 6
  AF_APPLETALK = 16
  AF_NETBIOS = 17
  AF_INET6 = 23
  AF_IRDA = 26
  AF_BTH = 32

End Enum

Enum sock_type
   SOCK_STREAM = 1
   SOCK_DGRAM = 2
   SOCK_RAW = 3
   SOCK_RDM = 4
   SOCK_SEQPACKET = 5
End Enum

Enum Protocol
   IPPROTO_ICMP = 1
   IPPROTO_IGMP = 2
   BTHPROTO_RFCOMM = 3
   IPPROTO_TCP = 6
   IPPROTO_UDP = 17
   IPPROTO_ICMPV6 = 58
   IPPROTO_RM = 113


End Enum

Type sockaddr
   sa_family As Integer
   sa_data(0 To 13) As Byte
End Type

Type sockaddr_in
  sin_family As Integer
  sin_port As Integer
  sin_addr(0 To 3) As Byte
  sin_zero(0 To 7) As Byte
End Type


'typedef UINT_PTR        SOCKET;
Type socket
   pointer As Long
End Type



' typedef struct WSAData {
'  WORD           wVersion;
'  WORD           wHighVersion;
'  char           szDescription[WSADESCRIPTION_LEN+1];
'  char           szSystemStatus[WSASYS_STATUS_LEN+1];
'  unsigned short iMaxSockets;
'  unsigned short iMaxUdpDg;
'  char FAR       *lpVendorInfo;
'} WSADATA, *LPWSADATA;

Type LPWSADATA_Type
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADESCRIPTION_LEN) As Byte
   szSystemStatus(0 To WSADESCRIPTION_LEN) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpVendorInfo As Long
End Type

'int errorno = WSAGetLastError()
Public Declare Function WSAGetLastError Lib "Ws2_32.dll" _
   () As Integer

'   int WSAStartup(
'  __in   WORD wVersionRequested,
'  __out  LPWSADATA lpWSAData
');
Public Declare Function WSAStartup Lib "Ws2_32.dll" _
    (ByVal wVersionRequested As Integer, ByRef lpWSAData As LPWSADATA_Type) As Long



'int sendto(
'  __in  SOCKET s,
'  __in  const char *buf,
'  __in  int len,
'  __in  int flags,
'  __in  const struct sockaddr *to,
'  __in  int tolen
');
Public Declare Function sendto Lib "Ws2_32.dll" _
    (ByVal socket As Long, ByRef buf() As Byte, _
     ByVal length As Long, ByVal Flags As Long, _
     ByRef toaddr As sockaddr_in, tolen As Long) As Long


'    SOCKET WSAAPI socket(
'  __in  int af,
'  __in  int type,
'  __in  int protocol
');

Public Declare Function f_socket Lib "Ws2_32.dll" Alias "socket" _
    (ByVal AF As Long, ByVal stype As Long, ByVal Protocol As Long) As Long

Public Declare Function closesocket Lib "Ws2_32.dll" _
    (ByVal socket As Long) As Long

Public Declare Sub WSACleanup Lib "Ws2_32.dll" ()
Sub main()
        Dim ConnectSocket As socket

        Dim wsaData As LPWSADATA_Type


        Dim iResult As Integer
           iResult = 0

        Dim send_sock As sock_type
          send_sock = INVALID_SOCKET

        Dim iFamily As AF
           iFamily = AF_INET

        Dim iType As Integer
           iType = SOCK_DGRAM


        Dim iProtocol As Integer
           iProtocol = IPPROTO_TCP

        Dim SendBuf(0 To 1023) As Byte
        Dim BufLen As Integer
        Dim ST As String
        ST = "{" & """ServiceCode""" & ":" & """1""" & "," & """Amount""" & ":" & """1000""" & "," & """PayerId""" & ":" & """108""" & "," & """PcID""" & ":" & """1234""" & "}"
        For i = 1 To Len(ST)
          SendBuf(i - 1) = Asc(MID(ST, i, 1))
        Next i
           BufLen = Len(ST)
        Dim RecvAddr As sockaddr_in
        RecvAddr.sin_family = AF_INET
        RecvAddr.sin_port = 1024
        RecvAddr.sin_addr(0) = 192
        RecvAddr.sin_addr(1) = 168
        RecvAddr.sin_addr(2) = 1
        RecvAddr.sin_addr(3) = 254


        iResult = WSAStartup(&H202, wsaData)
        If iResult <> 0 Then
           MsgBox ("WSAStartup failed: " & iResult)
           Exit Sub
        End If


        send_sock = f_socket(iFamily, iType, iProtocol)

        If send_sock = INVALID_SOCKET Then
           Errno = WSAGetLastError()
           Exit Sub
        End If

        iResult = sendto(send_sock, SendBuf, BufLen, 0, RecvAddr, Len(RecvAddr))
        If iResult = -1 Then
           MsgBox ("sendto failed with error: " & WSAGetLastError())
                closesocket (SendSocket)
                Call WSACleanup
                Exit Sub
        End If


        iResult = closesocket(send_sock)
        If iResult <> 0 Then
           MsgBox ("closesocket failed with error : " & WSAGetLastError())
           Call WSACleanup
        End If

End Sub
