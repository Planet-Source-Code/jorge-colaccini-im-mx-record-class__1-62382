VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imMXRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Class         : imMXRecord
' Date          : 29/Ago/2005
' Author        : Jorge Colaccini (JRC) <software(AT)informas.com>
' Purpose       : functions to access to MX Record of a domain
'---------------------------------------------------------------------------------------
' BASED ON:
'MX_Query:
'   modMXQuery from (vbSendMail v. 3.65) based on Gregg Housh's MX.OCX code
'       vbSendMail v. 3.65
'           Dean Dusenbery & FreeVBCode.com
'           http://www.freevbcode.com
'           http://www.freevbcode.com/ShowCode.Asp?ID=109
'
'if you are some credits on the based code, please, make to know me
'
'GetIPfromHostname: Ip Resolver
'   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31176&lngWId=1
'---------------------------------------------------------------------------------------

' winsock
Private Const DNS_RECURSION As Byte = 1
Private Const AF_INET = 2
Private Const SOCKET_ERROR = -1
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const SOCK_DGRAM = 2
Private Const INADDR_NONE = &HFFFFFFFF
Private Const INADDR_ANY = &H0

Private m_lWinSockTimeout As Long  'Timeout for winsock operations in milliseconds

' registry access
Private Const REG_SZ = 1&
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY

' winsock
Private Type WSADATA
    wversion                As Integer
    wHighVersion            As Integer
    szDescription(256)      As Byte
    szSystemStatus(128)     As Byte
    iMaxSockets             As Integer
    iMaxUdpDg               As Integer
    lpVendorInfo            As Long
End Type

Private Type DNS_HEADER
    qryID                   As Integer
    options                 As Byte
    response                As Byte
    qdcount                 As Integer
    ancount                 As Integer
    nscount                 As Integer
    arcount                 As Integer
End Type

Private Type IP_ADDRESS_STRING
    IpAddressStr(4 * 4 - 1) As Byte
End Type
 
Private Type IP_MASK_STRING
    IpMaskString(4 * 4 - 1) As Byte
End Type
 
Private Type IP_ADDR_STRING
    Next                    As Long
    IpAddress               As IP_ADDRESS_STRING
    IpMask                  As IP_MASK_STRING
    Context                 As Long
End Type

Private Type FIXED_INFO
    hostname(128 + 4 - 1)   As Byte
    DomainName(128 + 4 - 1) As Byte
    CurrentDnsServer        As Long
    DnsServerList           As IP_ADDR_STRING
    NodeType                As Long
    ScopeId(256 + 4 - 1)    As Byte
    EnableRouting           As Long
    EnableProxy             As Long
    EnableDns               As Long
End Type

Private Type SOCKADDR
    sin_family              As Integer
    sin_port                As Integer
    sin_addr                As Long
    sin_zero                As String * 8
End Type

Private Type HOSTENT
    h_name                  As Long
    h_aliases               As Long
    h_addrtype              As Integer
    h_length                As Integer
    h_addr_list             As Long
End Type

' registry
Private Type FILETIME
    dwLowDateTime           As Long
    dwHighDateTime          As Long
End Type

'type for passing DNS info
Private Type DNS_INFO
    Servers()               As String
    Count                   As Long
    LocalDomain             As String
    RootDomain              As String
End Type

' used below
Private Type MX_RECORD
    Server                  As String
    Pref                    As Integer
End Type

' type for passing MX info
Private Type MX_INFO
    Best                    As String
    Domain                  As String
    List()                  As MX_RECORD
    Count                   As Long
End Type

Private DNS                  As DNS_INFO
Private MX                   As MX_INFO


' API prototypes

' winsock, 'wsock32.dll' used instead of 'ws2_32.dll' for wider compatibility
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Private Declare Function recvfrom Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As SOCKADDR, fromlen As Long) As Long
Private Declare Function socket Lib "WSOCK32.DLL" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function htons Lib "WSOCK32.DLL" (ByVal hostshort As Long) As Integer
Private Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Integer
Private Declare Function sendto Lib "WSOCK32.DLL" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As SOCKADDR, ByVal tolen As Long) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

'Added by JRC to trap TIMEOUTs
Private Declare Function setsockopt Lib "WSOCK32.DLL" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Const SOL_SOCKET = 65535
Private Const SO_SNDTIMEO = &H1005&
Private Const SO_RCVTIMEO = &H1006&

' Registry access
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

' misc
Private Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Version used in GetIPfromHostname, note ByVal
Private Declare Sub CopyMemoryByVal Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)


Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long


'Sockets Initialize / CleanUp
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private m_sLastErrorMsg As String
Private m_bInitialized As Boolean
Public Sub GetDNSInfo()

    ' get the DNS servers and the local IP Domain name
    
    Dim sBuffer                 As String
    Dim sDNSBuff                As String
    Dim sDomainBuff             As String
    Dim sKey                    As String
    Dim lngFixedInfoNeeded      As Long
    Dim bytFixedInfoBuffer()    As Byte
    Dim udtFixedInfo            As FIXED_INFO
    Dim lngIpAddrStringPtr      As Long
    Dim udtIpAddrString         As IP_ADDR_STRING
    Dim strDnsIpAddress         As String
    Dim nRet                    As Long
    Dim sTmp()                  As String
    Dim i                       As Long
       
    ' get dns servers with the new GetNetworkParams call (only works on 98/ME/2000)
    ' if GetNetworkParams is not supported then try reading from the registry
    If Exported("iphlpapi.dll", "GetNetworkParams") Then
        nRet = GetNetworkParams(ByVal vbNullString, lngFixedInfoNeeded)
        If nRet = ERROR_BUFFER_OVERFLOW Then
            ReDim bytFixedInfoBuffer(lngFixedInfoNeeded)
            nRet = GetNetworkParams(bytFixedInfoBuffer(0), lngFixedInfoNeeded)
            CopyMemory udtFixedInfo, bytFixedInfoBuffer(0), Len(udtFixedInfo)
            With udtFixedInfo
                ' get the DNS servers
                lngIpAddrStringPtr = VarPtr(.DnsServerList)
                Do While lngIpAddrStringPtr
                    CopyMemory udtIpAddrString, ByVal lngIpAddrStringPtr, Len(udtIpAddrString)
                    With udtIpAddrString
                        strDnsIpAddress = StripTerminator(StrConv(.IpAddress.IpAddressStr, vbUnicode))
                        sDNSBuff = sDNSBuff & strDnsIpAddress & ","
                        lngIpAddrStringPtr = .Next
                    End With
                Loop
                ' get the ip domain name
                sDomainBuff = StripTerminator(StrConv(.DomainName, vbUnicode))
            End With
        End If
    End If
    
    ' if GetNetworkParams didn't get the data we need,
    ' try known locations in the registry for DNS & domain info
    If Len(sDNSBuff) = 0 Or Len(sDomainBuff) = 0 Then

        ' DNS servers configured through Network control panel applet (95/98/ME)
        sKey = "System\CurrentControlSet\Services\VxD\MSTCP"
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "NameServer", "")
        If Len(sBuffer) Then sDNSBuff = sBuffer & ","
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "Domain", "")
        If Len(sBuffer) Then sDomainBuff = sBuffer

        ' DNS servers configured through Network control panel applet (NT/2000)
        sKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "NameServer", "")
        If Len(sBuffer) Then sDNSBuff = sBuffer & ","
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "Domain", "")
        If Len(sBuffer) Then sDomainBuff = sBuffer

        ' DNS servers configured DHCP (NT/2000/XP)
        sKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "DhcpNameServer", "")
        If Len(sBuffer) Then sDNSBuff = sBuffer & ","
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "DHCPDomain", "")
        If Len(sBuffer) Then sDomainBuff = sBuffer

        ' DNS servers configured through Network control panel applet (XP)
        sKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces"
        sTmp = EnumRegKey(HKEY_LOCAL_MACHINE, sKey)
        For i = 0 To UBound(sTmp)
            sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey & "\" & sTmp(i), "NameServer", "")
            If Len(sBuffer) Then sDNSBuff = sBuffer & ","
            sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey & "\" & sTmp(i), "Domain", "")
            If Len(sBuffer) Then sDomainBuff = sBuffer
        Next
        
        ' DNS servers configured DHCP (95/98/ME)
        ' *** haven't found one ***
    
    End If

    ' get rid of any space delimiters (2000)
    sDNSBuff = Replace(sDNSBuff, " ", ",")

    ' trim any trailing commas
    If Right(sDNSBuff, 1) = "," Then sDNSBuff = Left(sDNSBuff, Len(sDNSBuff) - 1)

    ' load our type struc
    DNS.Servers = Split(sDNSBuff, ",")
    DNS.Count = UBound(DNS.Servers) + 1
    DNS.LocalDomain = sDomainBuff

    ' cheap trick
    If sDomainBuff = "" And DNS.Count > 0 Then
        sDomainBuff = GetRemoteHostName(DNS.Servers(0))
        nRet = InStr(sDomainBuff, ".")
        If nRet Then
            DNS.LocalDomain = Mid$(sDomainBuff, nRet + 1)
        End If
    End If

    sTmp = Split(sDomainBuff, ".")
    nRet = UBound(sTmp)
    If nRet > 0 Then
        DNS.RootDomain = sTmp(nRet - 1) & "." & sTmp(nRet)
    Else
        DNS.RootDomain = sDomainBuff
    End If

End Sub

Private Function MX_Query(ByVal ms_Domain As String) As String
'---------------------------------------------------------------------------------------
' Procedimiento : MX_Query
' FPO           : 27/Ago/2005 20:11
' Autor         : JRC
' Propósito     :
'---------------------------------------------------------------------------------------
' NOTEs
'       Original function
'---------------------------------------------------------------------------------------
    ' Performs the actual IP work to contact the DNS server,
    ' calls the other functions to parse and return the
    ' best server to send email through
    
    'Dim StartupData     As WSADATA
    Dim SocketBuffer    As SOCKADDR
    Dim IpAddr          As Long
    Dim iRC             As Integer
    Dim dnsHead         As DNS_HEADER
    Dim iSock           As Integer
    Dim dnsQuery()      As Byte
    Dim sQName          As String
    Dim dnsQueryNdx     As Integer
    Dim iTemp           As Integer
    Dim iNdx            As Integer
    Dim dnsReply(2048)  As Byte
    Dim iAnCount        As Integer
    Dim dwFlags         As Long


    'Para que funcione en dominios "loc.al" sin DNS
  'On Error GoTo MX_Query_Error
  
  '--------------------------------------------------------
  'Cuerpo del procedimiento. Inicio
  '--------------------------------------------------------
    MX.Count = 0
    MX.Best = vbNullString
    ReDim MX.List(0)
  
    
    '--------------------------------------------------------
    ' The following piece of code must be remark.
    ' The purpose is to trap local domains when I test in
    ' my intranet
    '--------------------------------------------------------
    If LCase$(Right$(ms_Domain, 6)) = "loc.al" _
       Or LCase$(Right$(ms_Domain, 5)) = "local" _
       Or LCase$(ms_Domain) = "localhost" _
       Then
         MX.Count = 1
         MX.Best = ms_Domain
         MX.List(0).Pref = 10
         MX.List(0).Server = ms_Domain
         MX_Query = ms_Domain
         Exit Function
    End If
    '--------------------------------------------------------
    
    ' if DNSInfo hasn't been called, call it now
    If DNS.Count = 0 Then GetDNSInfo
    
    ' check to see that we found a dns server
    If DNS.Count = 0 Then
        ' problem
        m_sLastErrorMsg = "No DNS entries found, MX Query cannot continue."
        Err.Raise 20000, "MXQuery", m_sLastErrorMsg
        MX_Query = ""
        Exit Function
    End If
   
    ' if null was passed in then use the local domain name
    If Len(ms_Domain) = 0 Then ms_Domain = DNS.LocalDomain
    
    ' validate domain name
    If Len(ms_Domain) < 5 Then
        m_sLastErrorMsg = "No Valid Domain Specified"
        Err.Raise 20000, "MXQuery", m_sLastErrorMsg
        MX_Query = ""
        Exit Function
    End If
   
    MX.Domain = ms_Domain
   
    ' Create a socket
    iSock = socket(AF_INET, SOCK_DGRAM, 0)
    If iSock = SOCKET_ERROR Then Exit Function
    
    '----------------------- IMPORTANTE --------------------------
    '
    ' THIS IS THE RELEVANT IMPROVMENT FROM ORIGINAL CODE
    ' IT PREVENT FREEZING WHEN THERE ARE A NON-EXISTENT
    ' DOMAIN OR DNS
    ' Jorge Colaccini
    '
    ' SET TIMEOUT IN MILLISECONDS TO PREVENT BLOCKING RECEIVE OPS
    '----------------------- IMPORTANTE --------------------------
    iRC = setsockopt(iSock, SOL_SOCKET, SO_SNDTIMEO, (m_lWinSockTimeout), 4&)
    iRC = setsockopt(iSock, SOL_SOCKET, SO_RCVTIMEO, (m_lWinSockTimeout), 4&)

    ' convert the IP address string to a network ordered long
    IpAddr = GetHostByNameAlias(DNS.Servers(0))
    If IpAddr = -1 Then Exit Function
    
    ' Setup the connnection parameters
    SocketBuffer.sin_family = AF_INET
    SocketBuffer.sin_port = htons(53)
    SocketBuffer.sin_addr = IpAddr
    SocketBuffer.sin_zero = String$(8, 0)
    
    ' Set the DNS parameters
    dnsHead.qryID = htons(&H11DF)
    dnsHead.options = DNS_RECURSION
    dnsHead.qdcount = htons(1)
    dnsHead.ancount = 0
    dnsHead.nscount = 0
    dnsHead.arcount = 0
    
    dnsQueryNdx = 0
    
    ReDim dnsQuery(4000)
    
    ' Setup the dns structure to send the query in
    ' First goes the DNS header information
    CopyMemory dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    ' Then the domain name (as a QNAME)
    sQName = MakeQName(MX.Domain)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
        'DoEvents   'JRC
    Wend
    
    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    ' Null terminate the string
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    ' The type of query (15 means MX query)
    iTemp = htons(15)
    CopyMemory dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ' The class of query (1 means INET)
    iTemp = htons(1)
    CopyMemory dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    
   
    
    ' Send the query to the DNS server
    iRC = sendto(iSock, dnsQuery(0), dnsQueryNdx + 1, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Or (iRC = 0) Then
        m_sLastErrorMsg = "Problem sending MX query"
        Err.Raise 20000, "MXQuery", m_sLastErrorMsg
        
        'iRC = WSACleanup
        MX_Query = ""
        Exit Function
    End If
    
    ' Wait for answer from the DNS server
    DoEvents
    Debug.Print Time, Timer
    iRC = recvfrom(iSock, dnsReply(0), 2048, 0, SocketBuffer, Len(SocketBuffer))
    Debug.Print Time, Timer
    
    If (iRC = SOCKET_ERROR) Or (iRC = 0) Then
        m_sLastErrorMsg = "Problem receiving MX query"
        Err.Raise 20000, "MXQuery", m_sLastErrorMsg
        'iRC = WSACleanup
        MX_Query = ""
        Exit Function
    End If
    
    ' Get the number of answers
    CopyMemory iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    
    
    If iAnCount Then
        ' Parse the answer buffer
        MX_Query = GetMXName(dnsReply(), 12, iAnCount)
        
    Else
        ' if we didn't find anything and we are part of
        ' a sub domain, go up one level and try again
        ' the last pass is at the root domain level
        If InStr(MX.Domain, DNS.RootDomain) > 1 Then
            MX.Domain = Mid$(MX.Domain, InStr(MX.Domain, ".") + 1)
            MX_Query = MX_Query(MX.Domain)
        End If
    End If


  '--------------------------------------------------------
  'Cuerpo del procedimiento. Salida
  '--------------------------------------------------------

MX_Query_Exit:
  On Error GoTo 0
  Exit Function

MX_Query_Error:
  '--------------------------------------------------------
  'Tratamiento de errores
  '--------------------------------------------------------
  
  Debug.Print Err.Number, Err.Description
  'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MX_Query of Módulo modMXQuery"
  Select Case Err
    'Case xx
    '  bResult = False
    '  Resume Next
    Case Else
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento MX_Query de Módulo modMXQuery"
  End Select
  GoTo MX_Query_Exit
    
End Function



' ************************************************************
' * Coder Name       : kegham
' * Web Site         : http://www.geocities.com/vbdotlb
' * E-Mail           : kegham_d@hotmail.com
' * Date             : 1/25/02
' **********************************************************************
' * Comments         :
' * Find IP address ginving the hostname
' **********************************************************************

Private Function hibyte(ByVal wParam As Integer)
   
   hibyte = wParam \ &H100 And &HFF&
   
End Function

Private Function lobyte(ByVal wParam As Integer)
   
   lobyte = wParam And &HFF&
   
End Function

Private Function SocketsInitialize() As Boolean
'---------------------------------------------------------------------------------------
' Procedimiento : SocketsInitialize
' FPO           : 29/Ago/2005 21:41
' Autor         : JRC
' Propósito     :
'---------------------------------------------------------------------------------------
'
   
   'Dim WSAD As WSADATA
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, m_sLastErrorMsg As String
   
   Dim bResult As Boolean
   
   
   Dim StartupData     As WSADATA
   ' Dim SocketBuffer    As SOCKADDR
   ' Dim IpAddr          As Long
   ' Dim iRC             As Integer
   ' Dim dnsHead         As DNS_HEADER
   ' Dim iSock           As Integer
   ' Dim dnsQuery()      As Byte
   ' Dim sQName          As String
   ' Dim dnsQueryNdx     As Integer
   ' Dim iTemp           As Integer
   ' Dim iNdx            As Integer
   ' Dim dnsReply(2048)  As Byte
   ' Dim iAnCount        As Integer
   ' Dim dwFlags         As Long
   
   bResult = False
   
    ' Initialize the Winsock, request v1.1
    iReturn = WSAStartup(WS_VERSION_REQD, StartupData)
    If iReturn <> ERROR_SUCCESS Then
       'iReturn = WSACleanup
       Exit Function
    End If
   
   'iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
   
   
   
'   If iReturn <> 0 Then
'      MsgBox "Winsock.dll is not responding."
'      End
'   End If
   
   If lobyte(StartupData.wversion) < WS_VERSION_MAJOR Or (lobyte(StartupData.wversion) = WS_VERSION_MAJOR And hibyte(StartupData.wversion) < WS_VERSION_MINOR) Then
      sHighByte = Trim$(Str$(hibyte(StartupData.wversion)))
      sLowByte = Trim$(Str$(lobyte(StartupData.wversion)))
      m_sLastErrorMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
      m_sLastErrorMsg = m_sLastErrorMsg & " is not supported by winsock.dll "
      'MsgBox m_sLastErrorMsg
      'End
      Exit Function
   End If
   
   If StartupData.iMaxSockets < MIN_SOCKETS_REQD Then
      m_sLastErrorMsg = "This application requires a minimum of "
      m_sLastErrorMsg = m_sLastErrorMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
      'MsgBox m_sLastErrorMsg
      'End
      Exit Function
   End If
   SocketsInitialize = True
End Function

Private Sub SocketsCleanup()
   Dim lReturn As Long
   
   lReturn = WSACleanup()
   
   If lReturn <> 0 Then
'      MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
'      End
   End If
   
End Sub

Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    
' Parse the server name out of the MX record, returns it in variable sName.
' iNdx is also modified to point to the end of the parsed structure.
    
    Dim iCompress       As Integer      ' Compression index (index to original buffer)
    Dim iChCount        As Integer      ' Character count (number of chars to read from buffer)
        
    ' While we dont encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
    
End Sub

Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    
' Parses the buffer returned by the DNS server, returns the best
' MX server (lowest preference number), iNdx is modified to point
' to the current buffer position (should be the end of the buffer
' by the end, unless a record other than MX is found)
    
    Dim iChCount        As Integer     ' Character counter
    Dim sTemp           As String      ' Holds the original query string
    Dim iBestPref       As Integer     ' Holds the "best" preference number (lowest)
    Dim iMXCount        As Integer
    
    
    MX.Count = 0
    MX.Best = vbNullString
    ReDim MX.List(0)

    iMXCount = 0
    iBestPref = -1
    
    ParseName dnsReply(), iNdx, sTemp
    
    ' Step over null
    iNdx = iNdx + 2
    
    ' Step over 6 bytes, not sure what the 6 bytes are, but
    ' all other documentation shows steping over these 6 bytes
    iNdx = iNdx + 6
    
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6
            
            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            CopyMemory iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            ParseName dnsReply(), iNdx, sName
            
            If Trim(sName) <> "" Then
                iMXCount = iMXCount + 1
                ReDim Preserve MX.List(iMXCount - 1)
                MX.List(iMXCount - 1).Server = sName
                MX.List(iMXCount - 1).Pref = iPref
                MX.Count = iMXCount
                If (iBestPref = -1 Or iPref < iBestPref) Then
                    iBestPref = iPref
                    MX.Best = sName
                End If
            End If
            ' Step over 3 useless bytes
            iNdx = iNdx + 3
        Else
            GetMXName = MX.Best
            SortMX MX.List
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    SortMX MX.List
        
    GetMXName = MX.Best

End Function

Private Function MakeQName(sDomain As String) As String
    
' Takes sDomain and converts it to the QNAME-type string.
' QNAME is how a DNS server expects the string.
'
' Example:  Pass -        mail.com
'           Returns -     &H4mail&H3com
'                          ^      ^
'                          |______|____ These two are character counters, they count
'                                       the number of characters appearing after them
    
    Dim iQCount         As Integer      ' Character count (between dots)
    Dim iNdx            As Integer      ' Index into sDomain string
    Dim iCount          As Integer      ' Total chars in sDomain string
    Dim sQName          As String       ' QNAME string
    Dim sDotName        As String       ' Temp string for chars between dots
    Dim sChar           As String       ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName
    
End Function

Private Function GetHostByNameAlias(ByVal sHostName As String) As Long
    
    'Return IP address as a long, in network byte order
    
    Dim phe             As Long
    Dim heDestHost      As HOSTENT
    Dim addrList        As Long
    Dim retIP           As Long
    
    retIP = inet_addr(sHostName)
    
    If retIP = INADDR_NONE Then
        phe = gethostbyname(sHostName)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, LenB(heDestHost)
            CopyMemory addrList, ByVal heDestHost.h_addr_list, 4
            CopyMemory retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    
    GetHostByNameAlias = retIP
    
End Function

Private Function StripTerminator(ByVal strString As String) As String
    
    ' strip off trailing NULL's from API calls
    
    Dim intZeroPos      As Integer

    intZeroPos = InStr(strString, vbNullChar)
    
    If intZeroPos > 1 Then
        StripTerminator = Trim$(Left$(strString, intZeroPos - 1))
    ElseIf intZeroPos = 1 Then
        StripTerminator = vbNullString
    Else
        StripTerminator = strString
    End If
    
End Function

Private Function GetRegStr(hKeyRoot As Long, ByVal sKeyName As String, ByVal sValueName As String, Optional ByVal Default As String = "") As String
   
   Dim lRet             As Long
   Dim hKey             As Long
   Dim lType            As Long
   Dim lBytes           As Long
   Dim sBuff            As String
   
   ' in case there's a permissions violation
   On Local Error GoTo Err_Reg

   ' Assume failure and set return to Default
   GetRegStr = Default

   ' Open the key
   lRet = RegOpenKeyEx(hKeyRoot, sKeyName, 0&, KEY_READ, hKey)
   If lRet = ERROR_SUCCESS Then
      
      ' Determine the buffer size
      lRet = RegQueryValueEx(hKey, sValueName, 0&, lType, ByVal sBuff, lBytes)
      If lRet = ERROR_SUCCESS Then
         ' size the buffer & call again
         If lBytes > 0 Then
            sBuff = Space(lBytes)
            lRet = RegQueryValueEx(hKey, sValueName, 0&, lType, ByVal sBuff, Len(sBuff))
            If lRet = ERROR_SUCCESS Then
               ' Trim NULL and return
               GetRegStr = Left(sBuff, lBytes - 1)
            End If
         End If
      End If
      Call RegCloseKey(hKey)
   End If
   
   Exit Function
   
Err_Reg:

  If hKey Then Call RegCloseKey(hKey)
   
End Function

Private Function EnumRegKey(hKeyRoot As Long, sKeyName As String) As String()
    

    Dim lRet            As Long
    Dim ft              As FILETIME
    Dim hKey            As Long
    Dim CurIdx          As Long
    Dim KeyName         As String
    Dim ClassName       As String
    Dim KeyLen          As Long
    Dim ClassLen        As Long
    Dim RESERVED        As Long
    Dim sEnum()         As String
    
    On Local Error GoTo Err_Enum
    
    ' initialize array
    EnumRegKey = Split("", "")
    
    ' Open the key
    lRet = RegOpenKeyEx(hKeyRoot, sKeyName, 0&, KEY_READ, hKey)
    If lRet <> ERROR_SUCCESS Then Exit Function
    
    ' the key opened so get all the sub keys
    Do
        ' get each sub key until lRet = error
        KeyLen = 2000
        ClassLen = 2000
        KeyName = String$(KeyLen, 0)
        ClassName = String$(ClassLen, 0)
        lRet = RegEnumKeyEx(hKey, CurIdx, KeyName, KeyLen, RESERVED, ClassName, ClassLen, ft)

        If lRet = ERROR_SUCCESS Then
            ReDim Preserve sEnum(CurIdx)
            sEnum(CurIdx) = Left$(KeyName, KeyLen)
        End If
    
        CurIdx = CurIdx + 1
        
    Loop While lRet = ERROR_SUCCESS
      
Err_Enum:

    EnumRegKey = sEnum
    If hKey Then Call RegCloseKey(hKey)

End Function

Private Function Exported(ByVal ModuleName As String, ByVal ProcName As String) As Boolean
   
    ' see if the api supports a call
    
    Dim hModule         As Long
    Dim lpProc          As Long
    Dim FreeLib         As Boolean
   
    ' check to see if the module is already
    ' mapped into this process.
    hModule = GetModuleHandle(ModuleName)
    If hModule = 0 Then
        ' not mapped, load the module into this process.
        hModule = LoadLibrary(ModuleName)
        FreeLib = True
    End If
   
    ' check the procedure address to verify it's exported.
    If hModule Then
        lpProc = GetProcAddress(hModule, ProcName)
        Exported = (lpProc <> 0)
    End If
   
    ' unload library if we loaded it here.
    If FreeLib Then Call FreeLibrary(hModule)
    
End Function

Private Sub SortMX(arr() As MX_RECORD, Optional ByVal bSortDesc As Boolean = False)

    ' simple bubble sort

    Dim ValMX           As MX_RECORD
    Dim index           As Long
    Dim firstItem       As Long
    Dim indexLimit      As Long
    Dim lastSwap        As Long

    firstItem = LBound(arr)
    lastSwap = UBound(arr)
    
    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For index = firstItem To indexLimit
            ValMX.Pref = arr(index).Pref
            ValMX.Server = arr(index).Server
            If (ValMX.Pref > arr(index + 1).Pref) Xor bSortDesc Then
                ' if the items are not in order, swap them
                arr(index).Pref = arr(index + 1).Pref
                arr(index).Server = arr(index + 1).Server
                arr(index + 1).Pref = ValMX.Pref
                arr(index + 1).Server = ValMX.Server
                lastSwap = index
            End If
        Next
    Loop While lastSwap

End Sub

Private Function GetRemoteHostName(ByVal strIpAddress As String) As String

    Dim udtHostEnt      As HOSTENT  ' HOSTENT structure
    Dim lngPtrHostEnt   As Long     ' pointer to HOSTENT
    Dim lngInetAddr     As Long     ' address as a Long value
    Dim strHostName     As String   ' string buffer for host name

    ' initialize the buffer
    strHostName = String(256, 0)

    ' Convert IP address to Long
    lngInetAddr = inet_addr(strIpAddress)
    If lngInetAddr = INADDR_NONE Then Exit Function
        
    ' Get the HostEnt structure pointer
    lngPtrHostEnt = gethostbyaddr(lngInetAddr, 4, AF_INET)
    If lngPtrHostEnt = 0 Then Exit Function
            
    ' Copy data into the HostEnt structure
    CopyMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
    CopyMemory ByVal strHostName, ByVal udtHostEnt.h_name, Len(strHostName)

    GetRemoteHostName = StripTerminator(strHostName)

End Function



Public Function MXRecordList(ByVal email As String) As String()
'---------------------------------------------------------------------------------------
' Procedimiento : MXRecord
' FPO           : 29/Ago/2005 18:15
' Autor         : JRC
' Propósito     : Retrieve de list of MX Records
'---------------------------------------------------------------------------------------
    Dim sResult As String
    Dim iPosArroba As Integer, i  As Integer
    Dim sArr() As String  '1, 0)
    If m_bInitialized Then
        sResult = Trim$(LCase$(email))
        iPosArroba = InStr(sResult, "@")
        If iPosArroba > 0 Then
          sResult = Mid$(sResult, iPosArroba + 1)
        End If
        sResult = MX_Query(sResult)
        If MX.Count > 0 Then
    '    MX.Count = 5
          ReDim sArr(1, MX.Count)
          For i = 0 To MX.Count - 1
    '        sArr(0, i) = "0, " & Format$(i)  'MX.List(i).Server
    '        sArr(1, i) = "1, " & Format$(i)  ' MX.List(i).Pref
            sArr(0, i) = MX.List(i).Server
            sArr(1, i) = MX.List(i).Pref
          Next
        End If
    End If
    MXRecordList = sArr
End Function

Public Function MXRecord(ByVal email As String, Optional ByVal Ordinal As Integer = 1) As String
'---------------------------------------------------------------------------------------
' Procedimiento : MXRecord
' FPO           : 29/Ago/2005 18:15
' Autor         : JRC
' Propósito     : Retrieve de best domain of MX Record,
'                   or the Ordinals if there more than one
'---------------------------------------------------------------------------------------
    Dim sResult As String
    Dim iPosArroba As Integer
    
    If m_bInitialized Then
        sResult = Trim$(LCase$(email))
        iPosArroba = InStr(sResult, "@")
        If iPosArroba > 0 Then
          sResult = Mid$(sResult, iPosArroba + 1)
        End If
        sResult = MX_Query(sResult)
        If Ordinal > 1 Then
          If DNS.Count >= Ordinal Then
            sResult = MX.List(Ordinal - 1).Server
          End If
        End If
    End If
    MXRecord = sResult
End Function

Public Property Get Count() As Integer
'---------------------------------------------------------------------------------------
' Procedimiento : Count
' FPO           : 29/Ago/2005 18:35
' Autor         : JRC
' Propósito     :
'---------------------------------------------------------------------------------------
'

    Count = MX.Count

End Property



Public Function GetIPfromHostname(sHostName As String) As String
'---------------------------------------------------------------------------------------
' Procedimiento : GetIPfromHostname
' FPO           : 29/Ago/2005 20:19
' Autor         : JRC
' Propósito     :
'---------------------------------------------------------------------------------------
'Basado en: Ip Resolver
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31176&lngWId=1
'---------------------------------------------------------------------------------------

    Dim sResult As String


   Dim HOSTENT_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   
   HOSTENT_addr = gethostbyname(sHostName)
   Debug.Print "HOSTENT_addr", HOSTENT_addr
   If HOSTENT_addr <> 0 Then
        'CopyMemory host, HOSTENT_addr, LenB(host)
        CopyMemoryByVal host, HOSTENT_addr, LenB(host)
             Debug.Print "host.hAddrList", host.h_addr_list
        
        CopyMemoryByVal hostip_addr, host.h_addr_list, 4
    
        Debug.Print "hostip_Addr", hostip_addr
        Debug.Print "host.hLength", host.h_length
    
        ReDim temp_ip_address(1 To host.h_length)
        CopyMemoryByVal temp_ip_address(1), hostip_addr, host.h_length
        
        For i = 1 To host.h_length
           ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
   End If

   GetIPfromHostname = ip_address  'sResult

End Function

Private Sub Class_Initialize()
    m_bInitialized = SocketsInitialize()
    m_lWinSockTimeout = 2000 ' 2 secs
End Sub


Private Sub Class_Terminate()
    SocketsCleanup
End Sub



Public Property Get LastErrorMsg() As String
'---------------------------------------------------------------------------------------
' Procedimiento : LastErrorMsg
' FPO           : 29/Ago/2005 23:25
' Autor         : JRC
' Propósito     : Return Last Error Messages
'---------------------------------------------------------------------------------------
'

    LastErrorMsg = m_sLastErrorMsg

End Property

Public Property Get Initialized() As Boolean
'---------------------------------------------------------------------------------------
' Procedimiento : Initialized
' FPO           : 29/Ago/2005 23:37
' Autor         : JRC
' Propósito     :
'---------------------------------------------------------------------------------------
'

    Initialized = m_bInitialized

End Property

Public Property Get Timeout() As Long

    Timeout = m_lWinSockTimeout

End Property

Public Property Let Timeout(ByVal lTimeout As Long)

    m_lWinSockTimeout = lTimeout

End Property
