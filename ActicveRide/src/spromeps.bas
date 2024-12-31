Attribute VB_Name = "SentinelProKeyLock"
' (C) Copyright 1991-2002 Rainbow Technologies, Inc. All rights reserved.

DefLng A-Z
' SSP API return code
Global Const SP_SUCCESS = 0
Global Const SP_INVALID_FUNCTION_CODE = 1
Global Const SP_INVALID_PACKET = 2
Global Const SP_UNIT_NOT_FOUND = 3
Global Const SP_ACCESS_DENIED = 4
Global Const SP_INVALID_MEMORY_ADDRESS = 5
Global Const SP_INVALID_ACCESS_CODE = 6
Global Const SP_PORT_IS_BUSY = 7
Global Const SP_WRITE_NOT_READY = 8
Global Const SP_NO_PORT_FOUND = 9
Global Const SP_ALREADY_ZERO = 10
Global Const SP_DRIVER_OPEN_ERROR = 11
Global Const SP_DRIVER_NOT_INSTALLED = 12
Global Const SP_IO_COMMUNICATIONS_ERROR = 13
Global Const SP_PACKET_TOO_SMALL = 15
Global Const SP_INVALID_PARAMETER = 16
Global Const SP_MEM_ACCESS_ERROR = 17
Global Const SP_VERSION_NOT_SUPPORTED = 18
Global Const SP_OS_NOT_SUPPORTED = 19
Global Const SP_QUERY_TOO_LONG = 20
Global Const SP_INVALID_COMMAND = 21
Global Const SP_MEM_ALIGNMENT_ERROR = 29
Global Const SP_DRIVER_IS_BUSY = 30
Global Const SP_PORT_ALLOCATION_FAILURE = 31
Global Const SP_PORT_RELEASE_FAILURE = 32
Global Const SP_ACQUIRE_PORT_TIMEOUT = 39
Global Const SP_SIGNAL_NOT_SUPPORTED = 42
Global Const SP_UNKNOWN_MACHINE = 44
Global Const SP_SYS_API_ERROR = 45
Global Const SP_UNIT_IS_BUSY = 46
Global Const SP_INVALID_PORT_TYPE = 47
Global Const SP_INVALID_MACH_TYPE = 48
Global Const SP_INVALID_IRQ_MASK = 49
Global Const SP_INVALID_CONT_METHOD = 50
Global Const SP_INVALID_PORT_FLAGS = 51
Global Const SP_INVALID_LOG_PORT_CFG = 52
Global Const SP_INVALID_OS_TYPE = 53
Global Const SP_INVALID_LOG_PORT_NUM = 54
Global Const SP_INVALID_ROUTER_FLGS = 56
Global Const SP_INIT_NOT_CALLED = 57
Global Const SP_DRVR_TYPE_NOT_SUPPORTED = 58
Global Const SP_FAIL_ON_DRIVER_COMM = 59
Global Const SP_SERVER_PROBABLY_NOT_UP = 60
Global Const SP_UNKNOWN_HOST = 61
Global Const SP_SENDTO_FAILED = 62
Global Const SP_SOCKET_CREATION_FAILED = 63
Global Const SP_NORESOURCES = 64
Global Const SP_BROADCAST_NOT_SUPPORTED = 65
Global Const SP_BAD_SERVER_MESSAGE = 66
Global Const SP_NO_SERVER_RUNNING = 67
Global Const SP_NO_NETWORK = 68
Global Const SP_NO_SERVER_RESPONSE = 69
Global Const SP_NO_LICENSE_AVAILABLE = 70
Global Const SP_INVALID_LICENSE = 71
Global Const SP_INVALID_OPERATION = 72
Global Const SP_BUFFER_TOO_SMALL = 73
Global Const SP_INTERNAL_ERROR = 74
Global Const SP_PACKET_ALREADY_INITIALIZED = 75
Global Const SP_PROTOCOL_NOT_INSTALLED = 76


'constants required for SetProtocol
Global Const NSPRO_TCP_PROTOCOL = 1
Global Const NSPRO_IPX_PROTOCOL = 2
Global Const NSPRO_NETBEUI_PROTOCOL = 4
Global Const NSPRO_SAP_PROTOCOL = 8

'constants required for Enum Flag
Global Const NSPRO_RET_ON_FIRST = 1
Global Const NSPRO_GET_ALL_SERVERS = 2
Global Const NSPRO_RET_ON_FIRST_AVAILABLE = 4

'constants required for HeartBeat
Global Const MAX_HEARTBEAT = 2592000
Global Const MIN_HEARTBEAT = 60
Global Const INFINITE_HEARTBEAT = &HFFFFFFFF

'constants required for showing OS driver type
Global Const RB_WINNT_SYS_DRVR = 5 ' Windows NT system driver
Global Const RB_WIN95_SYS_DRVR = 7  'Windows 95 system driver
Global Const RB_NW_LOCAL_DRVR = 8  'Netware local driver
  
        
        
'constants required for SetContactServer API
Global Const RNBO_STANDALONE = "RNBO_STANDALONE"
Global Const RNBO_SPN_DRIVER = "RNBO_SPN_DRIVER"
Global Const RNBO_SPN_LOCAL = "RNBO_SPN_LOCAL"
Global Const RNBO_SPN_BROADCAST = "RNBO_SPN_BROADCAST"
Global Const RNBO_SPN_ALL_MODES = "RNBO_SPN_ALL_MODES"
Global Const RNBO_SPN_SERVER_MODES = "RNBO_SPN_SERVER_MODES"

'Global constants required for SSP API
Global Const MAX_NAME_LEN = 64
Global Const MAX_ADDR_LEN = 32
Global Const API_PACKET_SZ = 4112
Global Const MAX_NUM_SERVERS = 10
Global Const SPRO_MAX_QUERY_SIZE = 56

'Global constants required for User input validation
Global Const MIN_QUERY_LEN = 8
Global Const MIN_ACCESS_CODE = 0
Global Const MAX_ACCESS_CODE = 3
Global Const MAX_CELL_DATA_LEN = 4
Global Const MAX_CELL_ADD_LEN = 2

Type APIPACKET
 Data(API_PACKET_SZ - 1) As Byte
End Type

Type DATAQUERY
 Data(SPRO_MAX_QUERY_SIZE - 1) As Byte
End Type

Type NSPRO_SERVER_INFO
  srvrAdd(MAX_ADDR_LEN - 1) As Byte
  numLicAvail As Integer
End Type

Type SrvrInfoArr
   srvrInfo(MAX_NUM_SERVERS - 1) As NSPRO_SERVER_INFO
End Type

Type NSPRO_KEY_MONITOR_INFO
    devID       As Integer
    hrdLmt      As Integer
    LicInUse    As Integer
    numTimedOut As Integer
    highestUse  As Integer
End Type

Type NSPRO_MONITOR_INFO
    srvrName(MAX_NAME_LEN - 1) As Byte
    srvrIPAdd(MAX_ADDR_LEN - 1) As Byte
    srvrIPXAdd(MAX_ADDR_LEN - 1) As Byte
    version(MAX_NAME_LEN - 1)  As Byte
    protocol       As Integer
    keyInfo        As NSPRO_KEY_MONITOR_INFO
End Type

Global Datain As DATAQUERY
Global Dataout As DATAQUERY
Global ApiPack As APIPACKET

'SSP APIs
Declare Function RNBOsproFormatPacket% Lib "Sx32w.dll" _
                                           (ApiPack As APIPACKET, _
                                            ByVal ApiPackSize As Integer)
                                            
Declare Function RNBOsproInitialize% Lib "Sx32w.dll" _
                                         (ApiPack As APIPACKET)
                                         
Declare Function RNBOsproGetFullStatus% Lib "Sx32w.dll" _
                                            (ApiPack As APIPACKET)
                                            
Declare Function RNBOsproGetVersion% Lib "Sx32w.dll" _
                                         (ApiPack As APIPACKET, _
                                          majv As Integer, _
                                          minv As Integer, _
                                          rev As Integer, _
                                          ostype As Integer)
                                          
Declare Function RNBOsproFindFirstUnit% Lib "Sx32w.dll" _
                                            (ApiPack As APIPACKET, _
                                             ByVal DEVELOPERID As Integer)
                                             
Declare Function RNBOsproFindNextUnit% Lib "Sx32w.dll" _
                                           (ApiPack As APIPACKET)
                                           
Declare Function RNBOsproRead% Lib "Sx32w.dll" _
                                   (ApiPack As APIPACKET, _
                                    ByVal address As Integer, _
                                    Datum As Integer)
                                    
Declare Function RNBOsproExtendedRead% Lib "Sx32w.dll" _
                                           (ApiPack As APIPACKET, _
                                            ByVal address As Integer, _
                                            Datum As Integer, _
                                            accessCode As Integer)
                                            
Declare Function RNBOsproWrite% Lib "Sx32w.dll" _
                                    (ApiPack As APIPACKET, _
                                     ByVal wPass As Integer, _
                                     ByVal address As Integer, _
                                     ByVal Datum As Integer, _
                                     ByVal accessCode As Integer)
                                     
Declare Function RNBOsproOverwrite% Lib "Sx32w.dll" _
                                        (ApiPack As APIPACKET, _
                                         ByVal wPass As Integer, _
                                         ByVal oPass1 As Integer, _
                                         ByVal oPass2 As Integer, _
                                         ByVal address As Integer, _
                                         ByVal Datum As Integer, _
                                         ByVal accessCode As Integer)
                                         
Declare Function RNBOsproDecrement% Lib "Sx32w.dll" _
                                        (ApiPack As APIPACKET, _
                                         ByVal wPass As Integer, _
                                         ByVal address As Integer)
                                         
Declare Function RNBOsproActivate% Lib "Sx32w.dll" _
                                       (ApiPack As APIPACKET, _
                                        ByVal wPass As Integer, _
                                        ByVal aPass1 As Integer, _
                                        ByVal aPass2 As Integer, _
                                        ByVal address As Integer)
                                        
Declare Function RNBOsproQuery% Lib "Sx32w.dll" _
                                    (ApiPack As APIPACKET, _
                                     ByVal address As Integer, _
                                     query As DATAQUERY, _
                                     response As DATAQUERY, _
                                     response32 As Long, _
                                     ByVal length As Integer)

Declare Function RNBOsproSetContactServer% Lib "Sx32w.dll" _
                                               (ApiPack As APIPACKET, _
                                                ByVal srvr As String)
                                                
Declare Function RNBOsproGetContactServer% Lib "Sx32w.dll" _
                                               (ApiPack As APIPACKET, _
                                                ByVal srvr As String, _
                                                ByVal strlen As Integer)
                                                
Declare Function RNBOsproGetSubLicense% Lib "Sx32w.dll" _
                                            (ApiPack As APIPACKET, _
                                             ByVal cellAdd As Integer)
                                             
Declare Function RNBOsproReleaseLicense% Lib "Sx32w.dll" _
                                             (ApiPack As APIPACKET, _
                                              ByVal cellAdd As Integer, _
                                              numSubLic As Long)
                                              
Declare Function RNBOsproGetHardLimit% Lib "Sx32w.dll" _
                                           (ApiPack As APIPACKET, _
                                            hrdLmt As Integer)
                                            
Declare Function RNBOsproEnumServer% Lib "Sx32w.dll" _
                                         (ByVal enumFlag As Integer, _
                                          ByVal devID As Long, _
                                          serverInfo As SrvrInfoArr, _
                                          numServers As Integer)
                                          
Declare Function RNBOsproGetKeyInfo% Lib "Sx32w.dll" _
                                         (ApiPack As APIPACKET, _
                                          ByVal devID As Long, _
                                          ByVal keyIndex As Integer, _
                                          monitorInfo As NSPRO_MONITOR_INFO)
                                          
Declare Function RNBOsproSetProtocol% Lib "Sx32w.dll" _
                                          (ApiPack As APIPACKET, _
                                           ByVal ProtocolFlag As Integer)
                                           
Declare Function RNBOsproSetHeartBeat% Lib "Sx32w.dll" _
                                           (ApiPack As APIPACKET, _
                                           ByVal heartbeat As Long)






'some global constants
Global XreadD, wPass, oPass1, oPass2, Datum, dID
Global XreadAcc%, aCode%, Data%, ProtocolFlag%
Global valid$, nl$
Global IsInitialized As Byte 'will tell whether the apipkt is initialized or not
Global errFlag As Integer ' the global error flag

Public Function ToHex(arg As String, Optional maxLen As Integer, Optional msgStr As String, Optional titleStr As String) As Long
    errFlag% = 0 'initialize the flag to "NO ERROR" mode
    arg = UCase(arg) 'CAPitalize it
    ln% = Len(arg)     'scan entire string
    If (maxLen) Then
        If ln% > maxLen Then
           MsgBox msgStr & " must be in hex with length not exceeding " & maxLen & " chars.", 0, titleStr
           errFlag% = -1
           Exit Function
        End If
    End If
    
    i% = ln%
    While i% > 0        'backwards for hex chars, blanks
        h$ = Mid$(arg$, i%, 1) '1-at-a-time
        If InStr(valid$, h$) = 0 Then errFlag% = 1 'non-hex
        If h$ = " " Then
            arg$ = Left$(arg$, i% - 1) + Right$(arg$, i% - 1)
            ln% = ln% - 1
        End If
        i% = i% - 1
    Wend
    '
    If errFlag% = 1 Then Adr = -1    'return err flag if problem
    If errFlag% = 0 And ln% > 0 Then
        For i% = 1 To ln%     'compute dec. addr. from hex digit
            t% = InStr(valid$, Mid$(arg$, i%, 1))   'next digit
            Adr = Adr * 16 + t% - 1 'make room for newest digit
        Next i%
    End If
    ToHex = Adr
    'input of 'FFFF' results in dec. value of 65535
    'BUT, we cannot take HEX$(no.>32767)
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
    'The function takes an unsigned Integer from and API and
    'converts it to a Long for display or arithmetic purposes
    If Value < 0 Then
        IntegerToUnsigned = Value + 65536 'the limit of unsigned int in C is 65535
    Else
        IntegerToUnsigned = Value
    End If
    '
End Function

Public Function ErrorPresent(invalidStr As String, title As String, maxLen As Integer) As Integer
   ErrorPresent = 1 'set the functions return value
    If errFlag% = 1 Then 'check if error flag has been set by the subroutine "getchars"
       MsgBox invalidStr + " must be hex with length not exceeding " + CStr(maxLen) + " digits.", vbOKOnly, title
       Exit Function
    ElseIf errFlag% = -1 Then 'user has not enterd any INPUT or has pressed cancel
'       ErrorPresent = 2
       Exit Function
    End If
    ErrorPresent = 0
End Function



