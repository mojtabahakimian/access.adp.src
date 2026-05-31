Attribute VB_Name = "network"
'This project requires the following components:
' - a form (Form1) with a textbox (Text1, Multiline=True)
'   and a command button (Command1)
' - a module (Module1)

'in Form1:

'In Module1:

'******************************************************************
'Created By Verburgh Peter.
' 07-23-2001
' verburgh.peter@skynet.be
'-------------------------------------
'With this small application , you can detect the IP's installed on your computer,
'including subnet mask , BroadcastAddr..
'
'I've wrote this because i've a programm that uses the winsock control, but,
'if you have multiple ip's  installed on your pc , you could get by using the Listen
' method the wrong ip ...
'Because Winsock.Localip => detects the default ip installed on your PC ,
' and in most of the cases it could be the LAN (nic) not the WAN (nic)
'So then you have to use the Bind function ,to bind to your right ip..
'but how do you know & find that ip ?
'you can find it now by this appl.. it check's in the api.. IP Table..
'******************************************************************


Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..

Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Type MIB_IPADDRTABLE
    dEntrys As Long   'NUMBER of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type
Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 1
Public Const MIB_IF_TYPE_TOKENRING = 2
Public Const MIB_IF_TYPE_FDDI = 3
Public Const MIB_IF_TYPE_PPP = 4
Public Const MIB_IF_TYPE_LOOPBACK = 5
Public Const MIB_IF_TYPE_SLIP = 6

Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    ADDRESS(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Boolean
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type

Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

'converts a Long  to a string
Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function GETIPADD() As String
Dim Ret As Long, TEL As Long
Dim bBytes() As Byte
Dim Listing As MIB_IPADDRTABLE
Dim ST As String
ST = ""

                 On Error Resume Next
GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Function
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, False

    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    'MsgBox "IP's found : " & Listing.dEntrys    => Founded ip installed on your PC..
'    ST = Listing.dEntrys & "   IP addresses found on your PC !!" & vbCrLf
'    ST = ST & "----------------------------------------" & vbCrLf
'    For TEL = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
       ' MsgBox bBytes(tel) & "."
         CopyMemory Listing.mIPInfo(TEL), bBytes(4 + (TEL * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(TEL))
         GETIPADD = ConvertAddressToString(Listing.mIPInfo(TEL).dwAddr)
'    Next

'MsgBox ConvertAddressToString(Listing.mIPInfo(1).dwAddr)
  Exit Function
End Function
