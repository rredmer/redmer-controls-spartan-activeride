VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form UsbKeyDiagnostics 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame UsbKeyMessages 
      Height          =   7995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   14102
      _Version        =   262144
      Caption         =   "Messages"
      Begin VB.TextBox UsbKeyText 
         Height          =   7665
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   210
         Width           =   11865
      End
   End
   Begin ActiveToolBars.SSActiveToolBars UsbToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "UsbKeyDiagnostics.frx":0000
      ToolBars        =   "UsbKeyDiagnostics.frx":0CF7
   End
End
Attribute VB_Name = "UsbKeyDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shared Routines                                           **
'**                                                                        **
'** Module.....: UsbKeyDiagnostics                                         **
'**                                                                        **
'** Description: Provides hardware dongle interface.                       **
'**                                                                        **
'** History....:                                                           **
'**    09/25/03 v1.00 RDR Implemented Class from existing code.            **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

' SSP API return code
Private Const SP_SUCCESS = 0
Private Const SP_INVALID_FUNCTION_CODE = 1
Private Const SP_INVALID_PACKET = 2
Private Const SP_UNIT_NOT_FOUND = 3
Private Const SP_ACCESS_DENIED = 4
Private Const SP_INVALID_MEMORY_ADDRESS = 5
Private Const SP_INVALID_ACCESS_CODE = 6
Private Const SP_PORT_IS_BUSY = 7
Private Const SP_WRITE_NOT_READY = 8
Private Const SP_NO_PORT_FOUND = 9
Private Const SP_ALREADY_ZERO = 10
Private Const SP_DRIVER_OPEN_ERROR = 11
Private Const SP_DRIVER_NOT_INSTALLED = 12
Private Const SP_IO_COMMUNICATIONS_ERROR = 13
Private Const SP_PACKET_TOO_SMALL = 15
Private Const SP_INVALID_PARAMETER = 16
Private Const SP_MEM_ACCESS_ERROR = 17
Private Const SP_VERSION_NOT_SUPPORTED = 18
Private Const SP_OS_NOT_SUPPORTED = 19
Private Const SP_QUERY_TOO_LONG = 20
Private Const SP_INVALID_COMMAND = 21
Private Const SP_MEM_ALIGNMENT_ERROR = 29
Private Const SP_DRIVER_IS_BUSY = 30
Private Const SP_PORT_ALLOCATION_FAILURE = 31
Private Const SP_PORT_RELEASE_FAILURE = 32
Private Const SP_ACQUIRE_PORT_TIMEOUT = 39
Private Const SP_SIGNAL_NOT_SUPPORTED = 42
Private Const SP_UNKNOWN_MACHINE = 44
Private Const SP_SYS_API_ERROR = 45
Private Const SP_UNIT_IS_BUSY = 46
Private Const SP_INVALID_PORT_TYPE = 47
Private Const SP_INVALID_MACH_TYPE = 48
Private Const SP_INVALID_IRQ_MASK = 49
Private Const SP_INVALID_CONT_METHOD = 50
Private Const SP_INVALID_PORT_FLAGS = 51
Private Const SP_INVALID_LOG_PORT_CFG = 52
Private Const SP_INVALID_OS_TYPE = 53
Private Const SP_INVALID_LOG_PORT_NUM = 54
Private Const SP_INVALID_ROUTER_FLGS = 56
Private Const SP_INIT_NOT_CALLED = 57
Private Const SP_DRVR_TYPE_NOT_SUPPORTED = 58
Private Const SP_FAIL_ON_DRIVER_COMM = 59
Private Const SP_SERVER_PROBABLY_NOT_UP = 60
Private Const SP_UNKNOWN_HOST = 61
Private Const SP_SENDTO_FAILED = 62
Private Const SP_SOCKET_CREATION_FAILED = 63
Private Const SP_NORESOURCES = 64
Private Const SP_BROADCAST_NOT_SUPPORTED = 65
Private Const SP_BAD_SERVER_MESSAGE = 66
Private Const SP_NO_SERVER_RUNNING = 67
Private Const SP_NO_NETWORK = 68
Private Const SP_NO_SERVER_RESPONSE = 69
Private Const SP_NO_LICENSE_AVAILABLE = 70
Private Const SP_INVALID_LICENSE = 71
Private Const SP_INVALID_OPERATION = 72
Private Const SP_BUFFER_TOO_SMALL = 73
Private Const SP_INTERNAL_ERROR = 74
Private Const SP_PACKET_ALREADY_INITIALIZED = 75
Private Const SP_PROTOCOL_NOT_INSTALLED = 76


'constants required for SetProtocol
Private Const NSPRO_TCP_PROTOCOL = 1
Private Const NSPRO_IPX_PROTOCOL = 2
Private Const NSPRO_NETBEUI_PROTOCOL = 4
Private Const NSPRO_SAP_PROTOCOL = 8

'constants required for Enum Flag
Private Const NSPRO_RET_ON_FIRST = 1
Private Const NSPRO_GET_ALL_SERVERS = 2
Private Const NSPRO_RET_ON_FIRST_AVAILABLE = 4

'constants required for HeartBeat
Private Const MAX_HEARTBEAT = 2592000
Private Const MIN_HEARTBEAT = 60
Private Const INFINITE_HEARTBEAT = &HFFFFFFFF

'constants required for showing OS driver type
Private Const RB_WINNT_SYS_DRVR = 5 ' Windows NT system driver
Private Const RB_WIN95_SYS_DRVR = 7  'Windows 95 system driver
Private Const RB_NW_LOCAL_DRVR = 8  'Netware local driver
        
'constants required for SetContactServer API
Private Const RNBO_STANDALONE = "RNBO_STANDALONE"
Private Const RNBO_SPN_DRIVER = "RNBO_SPN_DRIVER"
Private Const RNBO_SPN_LOCAL = "RNBO_SPN_LOCAL"
Private Const RNBO_SPN_BROADCAST = "RNBO_SPN_BROADCAST"
Private Const RNBO_SPN_ALL_MODES = "RNBO_SPN_ALL_MODES"
Private Const RNBO_SPN_SERVER_MODES = "RNBO_SPN_SERVER_MODES"

'private constants required for SSP API
Private Const MAX_NAME_LEN = 64
Private Const MAX_ADDR_LEN = 32
Private Const API_PACKET_SZ = 4112
Private Const MAX_NUM_SERVERS = 10
Private Const SPRO_MAX_QUERY_SIZE = 56

'private constants required for User input validation
Private Const MIN_QUERY_LEN = 8
Private Const MIN_ACCESS_CODE = 0
Private Const MAX_ACCESS_CODE = 3
Private Const MAX_CELL_DATA_LEN = 4
Private Const MAX_CELL_ADD_LEN = 2

Private Type APIPACKET
 Data(API_PACKET_SZ - 1) As Byte
End Type

Private Type DATAQUERY
 Data(SPRO_MAX_QUERY_SIZE - 1) As Byte
End Type

Private Type NSPRO_SERVER_INFO
  srvrAdd(MAX_ADDR_LEN - 1) As Byte
  numLicAvail As Integer
End Type

Private Type SrvrInfoArr
   srvrInfo(MAX_NUM_SERVERS - 1) As NSPRO_SERVER_INFO
End Type

Private Type NSPRO_KEY_MONITOR_INFO
    DevID       As Integer
    hrdLmt      As Integer
    LicInUse    As Integer
    numTimedOut As Integer
    highestUse  As Integer
End Type

Private Type NSPRO_MONITOR_INFO
    srvrName(MAX_NAME_LEN - 1) As Byte
    srvrIPAdd(MAX_ADDR_LEN - 1) As Byte
    srvrIPXAdd(MAX_ADDR_LEN - 1) As Byte
    version(MAX_NAME_LEN - 1)  As Byte
    protocol       As Integer
    keyInfo        As NSPRO_KEY_MONITOR_INFO
End Type

Private Datain As DATAQUERY
Private Dataout As DATAQUERY
Private ApiPack As APIPACKET

'SSP APIs
Private Declare Function RNBOsproFormatPacket% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal ApiPackSize As Integer)
Private Declare Function RNBOsproInitialize% Lib "Sx32w.dll" (ApiPack As APIPACKET)
Private Declare Function RNBOsproGetFullStatus% Lib "Sx32w.dll" (ApiPack As APIPACKET)
Private Declare Function RNBOsproGetVersion% Lib "Sx32w.dll" (ApiPack As APIPACKET, majv As Integer, minv As Integer, rev As Integer, ostype As Integer)
Private Declare Function RNBOsproFindFirstUnit% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal DeveloperID As Integer)
Private Declare Function RNBOsproFindNextUnit% Lib "Sx32w.dll" (ApiPack As APIPACKET)
Private Declare Function RNBOsproRead% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal address As Integer, Datum As Integer)
Private Declare Function RNBOsproExtendedRead% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal address As Integer, Datum As Integer, accessCode As Integer)
Private Declare Function RNBOsproWrite% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal address As Integer, ByVal Datum As Integer, ByVal accessCode As Integer)
Private Declare Function RNBOsproOverwrite% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal oPass1 As Integer, ByVal oPass2 As Integer, ByVal address As Integer, ByVal Datum As Integer, ByVal accessCode As Integer)
Private Declare Function RNBOsproDecrement% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal address As Integer)
Private Declare Function RNBOsproActivate% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal wPass As Integer, ByVal aPass1 As Integer, ByVal aPass2 As Integer, ByVal address As Integer)
Private Declare Function RNBOsproQuery% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal address As Integer, query As DATAQUERY, response As DATAQUERY, response32 As Long, ByVal length As Integer)
Private Declare Function RNBOsproSetContactServer% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal srvr As String)
Private Declare Function RNBOsproGetContactServer% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal srvr As String, ByVal strlen As Integer)
Private Declare Function RNBOsproGetSubLicense% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal cellAdd As Integer)
Private Declare Function RNBOsproReleaseLicense% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal cellAdd As Integer, numSubLic As Long)
Private Declare Function RNBOsproGetHardLimit% Lib "Sx32w.dll" (ApiPack As APIPACKET, hrdLmt As Integer)
Private Declare Function RNBOsproEnumServer% Lib "Sx32w.dll" (ByVal enumFlag As Integer, ByVal DevID As Long, serverInfo As SrvrInfoArr, numServers As Integer)
Private Declare Function RNBOsproGetKeyInfo% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal DevID As Long, ByVal keyIndex As Integer, monitorInfo As NSPRO_MONITOR_INFO)
Private Declare Function RNBOsproSetProtocol% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal ProtocolFlag As Integer)
Private Declare Function RNBOsproSetHeartBeat% Lib "Sx32w.dll" (ApiPack As APIPACKET, ByVal heartbeat As Long)

Private XreadD, wPass, oPass1, oPass2, Datum, dID
Private XreadAcc%, aCode%, Data%, ProtocolFlag%
Private valid$, nl$
Private IsInitialized As Byte 'will tell whether the apipkt is initialized or not
Private ErrFlag As Integer ' the private error flag
Private DeveloperID As Integer

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetSecurity                                           **
'**                                                                        **
'**  Description..:  This routine checks for Security Dongle (True=NonDemo)**
'**  The security dongle is a Rainbow Sentinnel Pro USB Device.  The app   **
'**  talks to the device using the Rainbow API DLL, defined by spromeps.bas**
'**                                                                        **
'****************************************************************************
Public Function GetSecurity(DevID As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim stval As Integer, Apisize As Integer, Datum As Integer, PC As New PerformanceCounter
    DeveloperID = DevID
    AddText "Getting application security mode from hardware dongle."
    GetSecurity = False
    PC.StartTimer True
    If IsInitialized = 1 Then
       IsInitialized = 0
       stval = RNBOsproReleaseLicense(ApiPack, 0, 0)
       AddText "Release liscense," & stval
    End If
    Apisize = 4096
    Datum = 0
    stval = RNBOsproFormatPacket(ApiPack, Apisize)
    AddText "FormatPacket," & stval
    stval = RNBOsproInitialize(ApiPack)
    AddText "Initialize," & stval
    If stval = 0 Then IsInitialized = 1
    stval = RNBOsproFindFirstUnit(ApiPack, DeveloperID)
    If stval = 0 Then
        AddText "FindFirstUnit," & stval
        stval = RNBOsproRead(ApiPack, 1, Datum)
        If stval = 0 Then
            If Datum < 0 Then Datum = 65536 + Datum 'no negative!
            If Datum <> DeveloperID Then
                AddText "Read wrong code," & Hex(Datum)
            Else
                '---- Following line for debugging key information
                AddText "Found and validated dongle."
                GetSecurity = True
            End If
        Else
            AddText "Read Failed," & stval
        End If
    Else
        AddText "FindFirstUnit did not find dongle," & stval
    End If
    AddText "Timed, " & Format(PC.StopTimer, "####.####") & " seconds."
    Set PC = Nothing
    Exit Function
ErrorHandler:
    PC.StopTimer
    Set PC = Nothing
    AddText "Error communicating with Usb Key.  Driver may be corrupted or not installed."
    ErrorForm.ReportError "Hardware:GetSecurity", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

Private Sub Form_Load()
    UsbKeyText.Text = ""
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrorHandler
    If Me.Width - 100 > 0 Then
        UsbKeyMessages.Width = Me.Width - 100
        UsbKeyText.Width = UsbKeyMessages.Width - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Sub UsbToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Refresh"
            UsbKeyText.Text = ""
            GetSecurity DeveloperID
    End Select
End Sub

Private Sub AddText(Msg As String)
    If Len(UsbKeyText.Text) > 32000 Then
        UsbKeyText.Text = ""
    Else
        UsbKeyText.Text = UsbKeyText.Text & Msg & vbCrLf
        UsbKeyText.Refresh
        AppLog InfoMsg, "UsbKeyDiagnostics," & Msg
    End If
End Sub

Public Function ToHex(arg As String, Optional maxLen As Integer, Optional msgStr As String, Optional titleStr As String) As Long
    Dim ln As Integer, i As Integer, ErrFlag As Integer, h As String, T As Integer, Adr As Integer
    ErrFlag% = 0 'initialize the flag to "NO ERROR" mode
    arg = UCase(arg) 'CAPitalize it
    ln% = Len(arg)     'scan entire string
    If (maxLen) Then
        If ln% > maxLen Then
           MsgBox msgStr & " must be in hex with length not exceeding " & maxLen & " chars.", 0, titleStr
           ErrFlag% = -1
           Exit Function
        End If
    End If
    
    i% = ln%
    While i% > 0        'backwards for hex chars, blanks
        h$ = Mid$(arg$, i%, 1) '1-at-a-time
        If InStr(valid$, h$) = 0 Then ErrFlag% = 1 'non-hex
        If h$ = " " Then
            arg$ = Left$(arg$, i% - 1) + Right$(arg$, i% - 1)
            ln% = ln% - 1
        End If
        i% = i% - 1
    Wend
    '
    If ErrFlag% = 1 Then Adr = -1    'return err flag if problem
    If ErrFlag% = 0 And ln% > 0 Then
        For i% = 1 To ln%     'compute dec. addr. from hex digit
            T% = InStr(valid$, Mid$(arg$, i%, 1))   'next digit
            Adr = Adr * 16 + T% - 1 'make room for newest digit
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
    If ErrFlag% = 1 Then 'check if error flag has been set by the subroutine "getchars"
       MsgBox invalidStr + " must be hex with length not exceeding " + CStr(maxLen) + " digits.", vbOKOnly, title
       Exit Function
    ElseIf ErrFlag% = -1 Then 'user has not enterd any INPUT or has pressed cancel
'       ErrorPresent = 2
       Exit Function
    End If
    ErrorPresent = 0
End Function





