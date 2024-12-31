VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{943CA7D1-C26F-4EA9-901A-2EA9BCAB0A49}#1.0#0"; "SaxComm8.ocx"
Begin VB.Form SetupSerialPort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Serial Port"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin SaxComm8Ctl.SaxComm SaxComm1 
      Height          =   1335
      Left            =   2400
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   1935
      _cx             =   3413
      _cy             =   2355
      Enabled         =   -1  'True
      Settings        =   "57600,n,8,1"
      BackColor       =   1
      Columns         =   80
      AutoProcess     =   0
      AutoScrollColumn=   0   'False
      AutoScrollKeyboard=   -1  'True
      AutoScrollRow   =   -1  'True
      AutoSize        =   0
      BackSpace       =   0
      CaptureFilename =   ""
      CaptureMode     =   0
      CDTimeOut       =   0
      ColorFilter     =   0
      Columns         =   80
      CommEcho        =   0   'False
      CommPort        =   "TOSHIBA Software Modem"
      CommSpy         =   0   'False
      CommSpyInput    =   -1  'True
      CommSpyOutput   =   -1  'True
      CommSpyProperties=   -1  'True
      CommSpyWarnings =   -1  'True
      CommSpyEvents   =   -1  'True
      CTSTimeOut      =   0
      DialMode        =   0
      DialTimeOut     =   60000
      DSRTimeOut      =   0
      Echo            =   0   'False
      Emulation       =   2
      EndOfLineMode   =   0
      ForeColor       =   15
      Handshaking     =   4
      IgnoreOnComm    =   -1  'True
      InBufferSize    =   65536
      InputEcho       =   0   'False
      InputLen        =   0
      InTimeOut       =   0
      OutTimeOut      =   0
      LookUpSeparator =   "|"
      LookUpText      =   ""
      LookUpTimeOut   =   10000
      NullDiscard     =   0   'False
      OutBufferSize   =   65536
      ParityReplace   =   ""
      Rows            =   25
      RThreshold      =   1
      RTSEnable       =   -1  'True
      ScrollRows      =   0
      SThreshold      =   0
      XferProtocol    =   5
      XferStatusDialog=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StatusbarVisible=   0   'False
      ToolbarVisible  =   0   'False
      StatusDialog    =   0
      UseTAPI         =   -1  'True
      BorderStyle     =   1
      SerialNumber    =   "1180-2431098-63"
      PhoneNumber     =   ""
      ProjectFilename =   ""
      CommSpyTransfer =   0   'False
      AutoZModem      =   -1  'True
   End
   Begin VB.TextBox PeriodicDelay 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "100"
      Top             =   450
      Width           =   885
   End
   Begin VB.Timer CommTimer 
      Interval        =   100
      Left            =   1920
      Top             =   780
   End
   Begin VB.ComboBox CommPortCombo 
      Height          =   315
      ItemData        =   "SetupSerialPort.frx":0000
      Left            =   930
      List            =   "SetupSerialPort.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   60
      Width           =   2835
   End
   Begin VB.CheckBox SendPeriodic 
      Caption         =   "Send periodic message every"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   2385
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1290
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupSerialPort.frx":003E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupSerialPort.frx":0360
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupSerialPort.frx":0682
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupSerialPort.frx":09A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   840
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Serial Port"
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "SetupSerialPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_BUFFER = 2048
Private Const HEADER_LEN = 3
Private Const CHECKSUM As Byte = &HAA
'Private Const MINIMUM_PACKET_SIZE = HEADER_LEN + 1
Private Const PACKET_HEADER_SIZE = HEADER_LEN
'Private Const DATA_COMMAND_OFFSET = PACKET_HEADER_SIZE + 1
'Private Const DEFAULT_PACKET_SEND_RETRY_COUNT = 3
'Private Const DEFAULT_PACKET_TIMEOUT = 1000

'Private sDiagnosticMessages(16) As String
Private sDiagnosticECU(7) As String

'Public sDiagnosticErrors As String          'Contains the diagnostic error code set
Public PacketsSent As Long                  'Number of packets sent
Public PacketsReceived As Long              'Number of packets received
Public PacketErrors As Long                 'Number of packet errors
Public CommStatus As String                 'Communications status
'Private CommBuffer As String                'Serial Communications buffer
'Private CommBufferLen As Integer            'Length of current serial communication command
Private CommTimeOut As Long
Private PeriodicTicks As Long               'Period tick count
Private bytSequence As Byte                 'Message Sequence (Actually implemented as Upper-Nibble of Command Byte)
'Private PacketHeaderReceived As Boolean     'Indicates whether 3 byte header has been received
Private WaitingForData As Boolean           'Indicates if we are waiting for a ACK or Data from the CCM
'Private InCommRoutine As Boolean
Private LastCommMessage As String
Private m_blnHeaderFound As Boolean
Private m_abytBuffer() As Byte
Private m_intCommandLength As Integer
Private m_intBufferLength As Integer

Private m_blnInReceive As Boolean

Private Const SIZEOF_SET_PARAMETER = 7
'Private Const SIZEOF_GET_PARAMETER = 5

Enum eCommand                                   'These are the CCM Command (Packet) Types
    Cmd = &HA0                                  'Command Packet
    Ack = &H60                                  'Acknowledge Packet
    Ret = &HB0                                  'Retry Packet
End Enum
Enum eDataCommand                               'These are the CCM Data Command Codes
    eGetParameter = &H1
    eSetParamater = &H2
    eGetAllParameters = &H3
    eBroadcastMessage = &H4
    eStartBroadcast = &H5
    eStopBroadcast = &H6
    eStartFirmwareDownload = &H7
    eDownloadStatus = &H8
    eDownloadData = &H9
    eDownloadComplete = &HA
    eDownloadVerified = &HB
    eDownloadEraseStatus = &HC
    eCCMStartup = &HE                            'Sent once from CCM when it activates
    eCommandError = &H11
    eDiagnosticMessage = &H12
    eClearDiagnosticMessage = &H13
    eKonect = &H16                              'Connection test message to CCM
End Enum
Enum eNodeMask
    CCM_NODE_MASK = 1
    LEFT_FRONT_NODE_MASK = 2
    RIGHT_FRONT_NODE_MASK = 4
    LEFT_REAR_NODE_MASK = 8
    RIGHT_REAR_NODE_MASK = 16
    LEFT_TAG_NODE_MASK = 32
    RIGHT_TAG_NODE_MASK = 64
    NODE_MASK_ALL = 128  ' Changed from 126 to 128 MCS 9/22/2003
End Enum
Enum DiagnosticLevel
    CCM_DIAGNOSTIC = 0
    LEFT_FRONT_DIAGNOSTIC = 1
    RIGHT_FRONT_DIAGNOSTIC = 2
    LEFT_REAR_DIAGNOSTIC = 3
    RIGHT_REAR_DIAGNOSTIC = 4
    LEFT_TAG_DIAGNOSTIC = 5
    RIGHT_TAG_DIAGNOSTIC = 6
    LVIT_POSITION = 7
    LVIT_VELOCITY = 8
    TEMPERATURE_SENSOR = 9
    SPEED_SENSOR = 10
    PITCH_SENSOR = 11
    ROLL_SENSOR = 12
    COMPRESSION = 13
    REBOUND = 14
    CANBUS = 15
End Enum
Enum DiagnosticMask
    CCM_DIAGNOSTIC = 1
    LEFT_FRONT_DIAGNOSTIC = 2
    RIGHT_FRONT_DIAGNOSTIC = 4
    LEFT_REAR_DIAGNOSTIC = 8
    RIGHT_REAR_DIAGNOSTIC = 16
    LEFT_TAG_DIAGNOSTIC = 32
    RIGHT_TAG_DIAGNOSTIC = 64
    LVIT_POSITION = 128
    LVIT_VELOCITY = 256
    TEMPERATURE_SENSOR = 512
    SPEED_SENSOR = 1024
    PITCH_SENSOR = 2048
    ROLL_SENSOR = 4096
    COMPRESSION = 8192
    REBOUND = 16384
    CANBUS = 32768
End Enum

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
100     sDiagnosticECU(1) = "CCM"
102     sDiagnosticECU(2) = "LF"
104     sDiagnosticECU(3) = "RF"
106     sDiagnosticECU(4) = "LR"
108     sDiagnosticECU(5) = "RR"
110     sDiagnosticECU(6) = "LT"
112     sDiagnosticECU(7) = "RT"

    ' {"CCM Failed", "Left Front REM Failed", "Right Front REM Failed", "Left Rear REM Failed", "Right Rear REM Failed", "Left Tag REM Failed", "Right Tag REM Failed", "LVIT Position Sensor Failed", "LVIT Velocity Sensor Failed", "Temperature Sensor Failed", "Speed Sensor Failed", "Pitch Sensor Failed", "Roll Sensor Failed", "Compression Solenoid Failed", "Rebound Solenoid Failed", "CAN Bus Failed"}
    '    sDiagnosticMessages(1) = "CCM Failed"
    '    sDiagnosticMessages(2) = "Left Front REM Failed"
    '    sDiagnosticMessages(3) = "Right Front REM Failed"
    '    sDiagnosticMessages(4) = "Left Rear REM Failed"
    '    sDiagnosticMessages(5) = "Right Rear REM Failed"
    '    sDiagnosticMessages(6) = "Left Tag REM Failed"
    '    sDiagnosticMessages(7) = "Right Tag REM Failed"
    '    sDiagnosticMessages(8) = "LVIT Position Sensor Failed"
    '    sDiagnosticMessages(9) = "LVIT Velocity Sensor Failed"
    '    sDiagnosticMessages(10) = "Temperature Sensor Failed"
    '    sDiagnosticMessages(11) = "Speed Sensor Failed"
    '    sDiagnosticMessages(12) = "Pitch Sensor Failed"
    '    sDiagnosticMessages(13) = "Roll Sensor Failed"
    '    sDiagnosticMessages(14) = "Compression Solenoid Failed"
    '    sDiagnosticMessages(15) = "Rebound Solenoid Failed"
    '    sDiagnosticMessages(16) = "CAN Bus Failed"

114     CommStatus = "Port Closed."    'Communications status
116     PacketsSent = 0                  'Number of packets sent
    '    sDiagnosticErrors = ""          'Contains the diagnostic error code set
118     PacketsReceived = 0              'Number of packets received
120     PacketErrors = 0                 'Number of packet errors
122     PeriodicTicks = 0               'Period tick count
124     bytSequence = 0                  'Message Sequence (Actually implemented as Upper-Nibble of Command Byte)
    '    CommBuffer = ""                'Serial Communications buffer
    '    CommBufferLen = 0                     'Length of current serial communication command
    '    PacketHeaderReceived = False          'Indicates whether 3 byte header has been received
126     WaitingForData = False       'Indicates if we are waiting for a ACK or Data from the CCM
128     CommTimeOut = 0
    '    InCommRoutine = False
        ' MCS
130     m_blnHeaderFound = False
132     m_intBufferLength = 0
134     m_blnInReceive = False
    
        ' MCS  Fill in the Serial Comm Connection
136     CommPortCombo.Text = GetSetting(App.CompanyName, App.ProductName, "CommPort", "Com1")

        ' Connect to the com port
138     ConnectCOM
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    ' Turn off the Comm Timer
    CommTimer.Enabled = False
    
'    CommPort.PortOpen = False
    If SaxComm1.PortOpen Then
        SaxComm1.PortOpen = False
    End If
    
'    If SaxComm1.Busy Then
        SaxComm1.ShutDown = True
'    End If

End Sub

' My version of the Comm Routine
Private Sub SaxComm1_Receive()

    On Error GoTo ErrHandler
    
'    Dim strCommand As String
    Dim intIndex As Integer
    Dim intBufferLength As Integer
'    Dim bytArray(256) As Byte
    Dim abytCommand() As Byte
'    Dim intByteIndex As Integer
    Dim intUpperBound As Integer
    Dim intLowerBound As Integer
    Dim intIndex2 As Integer
    Dim intBufferIndex As Integer
    Dim bytBuffer() As Byte
'    Dim intBufferLength As Integer
    
    m_blnInReceive = True
    
    bytBuffer = SaxComm1.Input
    
    ' Get the Comm Buffer
    If m_intBufferLength <> 0 Then
        intUpperBound = UBound(bytBuffer)
        intLowerBound = LBound(bytBuffer)
        ' -1 When the buffer from the comm control is empty
        If intUpperBound <> -1 Then
            intBufferLength = UBound(m_abytBuffer) + 1
            ReDim Preserve m_abytBuffer(intBufferLength + intUpperBound) As Byte
            For intBufferIndex = intLowerBound To intUpperBound
                m_abytBuffer(intBufferLength + intBufferIndex) = bytBuffer(intBufferIndex)
            Next
            m_intBufferLength = UBound(m_abytBuffer) + 1
        End If
'        m_abytBuffer = m_abytBuffer & SaxComm1.Input
    Else
        m_abytBuffer = bytBuffer
    End If

    ' Get the Upper and Lower Bound of the Buffer Array
    intUpperBound = UBound(m_abytBuffer)
    intLowerBound = LBound(m_abytBuffer)

    ' Check for length
    If intUpperBound - intLowerBound >= 2 Then
        ' Look for Header
        If Not m_blnHeaderFound Then
            ' Check the command Length
            For intIndex = intLowerBound To intUpperBound - 2
                m_intCommandLength = CInt(m_abytBuffer(intIndex)) + CInt(m_abytBuffer(intIndex + 1))
                ' 5/24/2003 Added length 7
                If m_intCommandLength = 11 Or m_intCommandLength = 34 Or _
                        m_intCommandLength = 4 Or m_intCommandLength = 14 Or _
                        m_intCommandLength = 6 Or m_intCommandLength = 8 Or _
                        m_intCommandLength = 7 Then
                     ' Found a Good Command Length
                     m_blnHeaderFound = True
                     ReDim abytCommand(m_intCommandLength - 1) As Byte
                     Exit For
                End If
            Next intIndex
        Else
            ReDim abytCommand(m_intCommandLength - 1) As Byte
        End If

        If m_blnHeaderFound Then
            intBufferLength = intUpperBound - intLowerBound + 1
            ' Trim Extra buffer information
            If intIndex <> 0 Then
'                m_abytBuffer = Right$(m_abytBuffer, intBufferLength - intIndex + 1)
                ' New buffer length
                intBufferLength = intUpperBound - intIndex
            End If

            ' Grab command if buffers long enough
            If intBufferLength > m_intCommandLength Then
                ' Grab the Command
'                strCommand = Left$(m_abytBuffer, m_intCommandLength)
'                ReDim abytCommand(m_intCommandLength - 1) As Byte
                For intIndex2 = intIndex To intIndex + m_intCommandLength - 1
                    abytCommand(intIndex2 - intIndex) = m_abytBuffer(intIndex2)
                Next
                ' Save the remaining Buffer
'                m_abytBuffer = Right$(m_abytBuffer, intBufferLength - m_intCommandLength)
                For intBufferIndex = intIndex + m_intCommandLength To intUpperBound
                    m_abytBuffer(intBufferIndex - intIndex - m_intCommandLength) = m_abytBuffer(intBufferIndex)
                Next
                ReDim Preserve m_abytBuffer(intUpperBound - intIndex - m_intCommandLength) As Byte
                m_intBufferLength = UBound(m_abytBuffer) + 1
            ElseIf intBufferLength = m_intCommandLength Then
                abytCommand = m_abytBuffer
                ' Clear Buffer
                Erase m_abytBuffer
                ' Set Buffer Length to Zero
                m_intBufferLength = 0
            Else
                ' Save the remaining Buffer
'                m_abytBuffer = Right$(m_abytBuffer, intBufferLength - m_intCommandLength)
                For intBufferIndex = intIndex + m_intCommandLength To intUpperBound
                    m_abytBuffer(intBufferIndex - intIndex - m_intCommandLength) = m_abytBuffer(intBufferIndex)
                Next
                ReDim Preserve m_abytBuffer(intUpperBound - intIndex) As Byte
                m_intBufferLength = UBound(m_abytBuffer) + 1

                ' Go Wait for more data
                Exit Sub
            End If
            
            ' Process the Command
            Call ProcessCommand(abytCommand, m_intCommandLength)
            
            m_blnHeaderFound = False
            WaitingForData = False
            CommStatus = "Port Open"
            CommTimeOut = 0
        Else
            ' Added MCS 1/23/2003 see what's being dumped
                    Dim sMsg As String
                    
                    sMsg = ""
                    For intIndex = 0 To intUpperBound
                        sMsg = sMsg & m_abytBuffer(intIndex) & " "
                    Next
                    ViewLog.Log ErrorMsg, "Bad Data " & sMsg
            ' Clear buffer
            Erase m_abytBuffer
            ' Set Buffer Length to Zero
            m_intBufferLength = 0
            WaitingForData = False
            CommStatus = "Port Open"
            CommTimeOut = 0
        End If
    End If
    WaitingForData = False
    CommStatus = "Port Open"
    CommTimeOut = 0
    
    m_blnInReceive = False

    Exit Sub
    
ErrHandler:

    ViewLog.Log ErrorMsg, "Error occurred in Saxcomm1.Receive - Error: " & Err.Description

End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo ToolBar_ButtonClick_Err
        '</EhHeader>
    
100     Select Case Button.index
            Case 1
                'Exit
102             Me.Hide
                ' Save the com port setting
104             SaveSetting App.CompanyName, App.ProductName, "CommPort", CommPortCombo.Text
106         Case 2
                'Reset
108             Call CommReset
        End Select
    
        '<EhFooter>
        Exit Sub

ToolBar_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.ToolBar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub UpdateGraphs(ByVal slot0 As Long, ByVal slot1 As Long, ByVal slot2 As Long, ByVal slot3 As Long)
        '<EhHeader>
        On Error GoTo UpdateGraphs_Err
        '</EhHeader>

        '---- Process data for each graph slot
        Dim GraphNum As Integer
        Dim lngData As Long
        Dim dblData As Double
    
100     With MainForm
102         For GraphNum = 0 To 3
104             If g_blnGraphEnable(GraphNum) Then
106                 Select Case GraphNum
                        Case 0
108                         lngData = slot0
110                     Case 1
112                         lngData = slot1
114                     Case 2
116                         lngData = slot2
118                     Case 3
120                         lngData = slot3
                    End Select
                    ' Fixed negative Numbers only if min isn't zero
122                 If lngData > 32768 And g_adblMinValue(GraphNum) <> 0 Then
124                     lngData = lngData - 65536
                    End If
                
126                 If g_intGraphDivisor(GraphNum) > 0 Then
128                     dblData = (lngData / (10 ^ g_intGraphDivisor(GraphNum)))
                    Else
130                     dblData = lngData
                    End If
                
132                 .GraphChart(GraphNum).OpenDataEx COD_VALUES Or COD_ADDPOINTS, 1, 1
134                 .GraphChart(GraphNum).Value(0) = dblData
136                 .GraphValueLabel(GraphNum).Caption = dblData
138                 .GraphChart(GraphNum).CloseData COD_VALUES Or COD_REALTIMESCROLL
                End If
            Next
        End With
        '<EhFooter>
        Exit Sub

UpdateGraphs_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.UpdateGraphs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function IncomingCheckSum(aBytes() As Byte) As Byte

    Dim iByte As Integer
    Dim bytChecksum As Integer
    Dim bytReturn As Byte
    Dim intUpperBound As Integer
    
    intUpperBound = UBound(aBytes)
    
    bytChecksum = CHECKSUM
    For iByte = 0 To intUpperBound
        bytChecksum = bytChecksum - aBytes(iByte)
        If bytChecksum < 0 Then
            bytChecksum = 256 + bytChecksum
        End If
    Next
    bytReturn = CByte(IIf(bytChecksum < 0, 256 + bytChecksum, bytChecksum))
    IncomingCheckSum = bytReturn

End Function

Private Sub ConnectCOM()
    
    On Error GoTo ErrorHandler
'    With CommPort                           'This is simply an MS-COMM Control
'        If .PortOpen = True Then            'If the port is open
'            .PortOpen = False               'Close the comm port
'        End If
'        .InBufferSize = MAX_BUFFER
'        .CommPort = CommPortCombo.ListIndex + 1
'        .Settings = "57600,n,8,1"           'Communications settings
'        .RThreshold = 1
'        If .PortOpen = False Then           'If the port is not open
'            .PortOpen = True                'Open the port
'        End If
'    End With
    With SaxComm1                           'This is simply an MS-COMM Control
        If .PortOpen = True Then            'If the port is open
            .PortOpen = False               'Close the comm port
        End If
        .InBufferSize = MAX_BUFFER
        .CommPort = CommPortCombo.ListIndex + 1
        .Settings = "57600,n,8,1"           'Communications settings
        .RThreshold = 1
        If .PortOpen = False Then           'If the port is not open
            .PortOpen = True                'Open the port
        End If
        ' Added 3/26/2003 Changed the Input Type to Binary the Conversion
        .InputMode = InputMode_Binary
    End With
    CommStatus = "Port Open."
    CommTimer.Interval = 100    'mcs
    CommTimer.Enabled = True
    MainForm.StatusBar.Panels(1).Text = SaxComm1.CommPort & ":" & SaxComm1.Settings
    MainForm.StatusBar.Panels(2).Text = CommStatus
    Exit Sub

ErrorHandler:
    
    ErrorForm.ReportError Me.Name & ":FileSave", Err.Number, Err.LastDllError, Err.Source, Err.Description, True

End Sub

Private Sub CommTimer_Timer()
    '--- Timer for COMM Port Timeout
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    If WaitingForData = True Then
        CommTimeOut = CommTimeOut + 1
        If CommTimeOut > 10 Then
            '---- COMM TIMEOUT ERROR
            CommTimeOut = 0
            WaitingForData = False
            PacketErrors = PacketErrors + 1
            CommStatus = " TIMEOUT ERROR."
            Call ViewLog.Log(LogMsgTypes.ErrorMsg, CommStatus)
            Beep
            MainForm.StatusBar.Panels(2).Text = CommStatus
            MainForm.UpdateStatus
        End If
    End If
    '---- Send periodic messages
    If SendPeriodic.Value = 1 Then
        '--- Send eKonnect message (tests communications)
        If Val(PeriodicDelay.Text) <> 0 And PeriodicTicks >= Val(PeriodicDelay.Text) Then
            Call ComSend(CreateCommandPacket(eDataCommand.eKonect, eNodeMask.CCM_NODE_MASK, 0, 0))
            PeriodicTicks = 0
        Else
            PeriodicTicks = PeriodicTicks + 100
        End If
    End If
    ' MCS Trying to get the comm to pull data faster
'    DoEvents
    If Not m_blnInReceive And m_intBufferLength <> 0 Then
        Call SaxComm1_Receive
    End If
    
End Sub

Public Function ComSend(ByRef bytOut() As Byte) As Boolean

    If SaxComm1.PortOpen Then
        
        SaxComm1.Output = bytOut
        PacketsSent = PacketsSent + 1
        If ViewLog.DebugEnable.Value = 1 Then
            Dim sMsg As String
            Dim iByte As Integer
    '
            sMsg = "TX:"
            For iByte = 0 To UBound(bytOut)
                sMsg = sMsg & bytOut(iByte) & " "
            Next
            Call ViewLog.Log(LogMsgTypes.DebugMsg, sMsg)
        End If
        CommStatus = "Port Open."
        MainForm.StatusBar.Panels(2).Text = CommStatus
        MainForm.UpdateStatus
    End If
    
    'mcs
'    DoEvents
        
End Function

Public Function CreateCommandPacket(ByVal DatCmd As eDataCommand, _
                                    ByVal byNodeMask As eNodeMask, _
                                    ByVal bytParm As Byte, _
                                    ByVal lData As Currency) As Byte()
    
    On Error GoTo ErrHandler
    
    Dim bytCommand() As Byte
    Dim sCmd As String
    
    ReDim bytCommand(HEADER_LEN + SIZEOF_SET_PARAMETER)
    
    ' Length
    bytCommand(0) = HEADER_LEN + SIZEOF_SET_PARAMETER + 1
    bytCommand(1) = 0
    ' ID & Sequence
    bytCommand(2) = SetCommandByte(eCommand.Cmd)
    ' Node
    bytCommand(3) = byNodeMask
    ' Command
    bytCommand(4) = DatCmd
    bytCommand(5) = bytParm
    sCmd = MyHex(lData, 8)                '.ToString("x8")
    bytCommand(6) = CByte(Val("&H" & Mid$(sCmd, 7, 2)))     'Byte order determined from Rideware
    bytCommand(7) = CByte(Val("&H" & Mid$(sCmd, 5, 2)))
    bytCommand(8) = CByte(Val("&H" & Mid$(sCmd, 3, 2)))
    bytCommand(9) = CByte(Val("&H" & Mid$(sCmd, 1, 2)))
    bytCommand(10) = CreateCheckSum(bytCommand)
    
    WaitingForData = True
    CreateCommandPacket = bytCommand
    
    Exit Function
    
ErrHandler:

    Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Error in CreateCommandPacket " & _
                        Err.Description & " Node: " & byNodeMask & _
                        " Command: " & DatCmd & " Parameter: " & bytParm)
    Resume Next

End Function

Public Function CreateDownloadCompletePacket(ByVal DatCmd As eDataCommand, ByVal byNodeMask As eNodeMask, ByVal lData As Currency) As Byte()
    
    On Error GoTo ErrHandler
    
    Dim bytCommand() As Byte
    Dim sCmd As String
    
    ReDim bytCommand(9)
    
    ' Length
    bytCommand(0) = 10
    bytCommand(1) = 0
    ' Id & Sequence
    bytCommand(2) = SetCommandByte(eCommand.Cmd)
    ' Node
    bytCommand(3) = byNodeMask
    ' Command
    bytCommand(4) = DatCmd
    
    sCmd = MyHex(lData, 8)                '.ToString("x8")
    
    bytCommand(5) = CByte(Val("&H" & Mid$(sCmd, 7, 2)))     'Byte order determined from Rideware
    bytCommand(6) = CByte(Val("&H" & Mid$(sCmd, 5, 2)))
    bytCommand(7) = CByte(Val("&H" & Mid$(sCmd, 3, 2)))
    bytCommand(8) = CByte(Val("&H" & Mid$(sCmd, 1, 2)))
    bytCommand(9) = CreateCheckSum(bytCommand)
    
    WaitingForData = True
    CreateDownloadCompletePacket = bytCommand
    
    Exit Function
    
ErrHandler:

    Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Error in CreateDownloadCompletePacket " & _
                        Err.Description & " Node: " & byNodeMask)
    Resume Next

End Function

Public Function CreateBroadcastPacket(ByVal DatCmd As eDataCommand, ByVal byNodeMask As eNodeMask, ByVal bytParm As Byte, ByVal bytSlot As Byte) As Byte()
    
    Dim bytPacket(8) As Byte
    
    ' 2 bytes Packet Length
    bytPacket(0) = 9
    bytPacket(1) = 0
    ' 1 byte ID & Sequence
    bytPacket(2) = SetCommandByte(eCommand.Cmd)
    ' 1 byte Node
    bytPacket(3) = 1                                           'All broadcasts go thru CCM
    ' 1 byte Command
    bytPacket(4) = DatCmd
    bytPacket(5) = byNodeMask
    bytPacket(6) = bytSlot
    bytPacket(7) = bytParm
    bytPacket(8) = CreateCheckSum(bytPacket)
    
    WaitingForData = True
    CreateBroadcastPacket = bytPacket
    
End Function

Private Function CreateCheckSum(ByRef bytCommand() As Byte) As Byte
    
    Dim iByte As Integer
    Dim bytChecksum As Integer
    Dim bytReturn As Byte
    
    bytChecksum = CHECKSUM
    ' MCS Changed start at 1 go to end of string not -1
    For iByte = 0 To UBound(bytCommand)
        bytChecksum = bytChecksum - bytCommand(iByte)
        If bytChecksum < 0 Then
            bytChecksum = 256 + bytChecksum
        End If
    Next
    
    bytReturn = CByte(IIf(bytChecksum < 0, 256 + bytChecksum, bytChecksum))
    CreateCheckSum = bytReturn

End Function

Public Function CreateStartDownloadPacket(ByVal DatCmd As eDataCommand, ByVal byNodeMask As eNodeMask, ByVal bytParm As Byte, ByVal lStartAddr As Currency, ByVal lEndAddr As Currency) As Byte()
    
    Dim sCmd As String
    Dim sCmd2 As String
    Dim bytPacket(14) As Byte
    
    sCmd = MyHex(lStartAddr, 8)             '.ToString("x8")
    sCmd2 = MyHex(lEndAddr, 8)              '.ToString("x8")
'    sCmd = "00030000"             '.ToString("x8")
'    sCmd2 = "00039F3E"              '.ToString("x8")
    
'    Dim sMsg As String
    
    ' 2 bytes Length
    bytPacket(0) = 15
    bytPacket(1) = 0
    ' 1 byte Id & Sequence
    bytPacket(2) = SetCommandByte(eCommand.Cmd)
    bytPacket(3) = byNodeMask
    ' 1 byte Command
    bytPacket(4) = DatCmd
    bytPacket(5) = bytParm
    bytPacket(6) = CByte(Val("&H" & Mid$(sCmd, 7, 2)))   'Byte order determined from Rideware
    bytPacket(7) = CByte(Val("&H" & Mid$(sCmd, 5, 2)))
    bytPacket(8) = CByte(Val("&H" & Mid$(sCmd, 3, 2)))
    bytPacket(9) = CByte(Val("&H" & Mid$(sCmd, 1, 2)))
    bytPacket(10) = CByte(Val("&H" & Mid$(sCmd2, 7, 2)))     'Byte order determined from Rideware
    bytPacket(11) = CByte(Val("&H" & Mid$(sCmd2, 5, 2)))
    bytPacket(12) = CByte(Val("&H" & Mid$(sCmd2, 3, 2)))
    bytPacket(13) = CByte(Val("&H" & Mid$(sCmd2, 1, 2)))
    bytPacket(14) = CreateCheckSum(bytPacket)
    
    WaitingForData = True
    CreateStartDownloadPacket = bytPacket
    
End Function

Private Function CreateAckPacket(ByVal byNodeMask As eNodeMask, ByVal bytCmd As Byte) As Byte()
    
    Dim bytPacket(3) As Byte
    
    bytPacket(0) = 4
    bytPacket(1) = 0
    bytPacket(2) = eCommand.Ack + (bytCmd And &HF)
    bytPacket(3) = CreateCheckSum(bytPacket)
    
    CreateAckPacket = bytPacket

End Function

Public Function CreateGetPacket(ByVal DatCmd As eDataCommand, ByVal byNodeMask As eNodeMask, ByVal bytParm As Byte) As Byte()
    
    Dim bytPacket(8) As Byte
    
    bytPacket(0) = 9
    bytPacket(1) = 0
    bytPacket(2) = SetCommandByte(eCommand.Cmd)
    bytPacket(3) = byNodeMask
    bytPacket(4) = DatCmd
    bytPacket(5) = bytParm
    bytPacket(6) = 0
    bytPacket(7) = 0
    bytPacket(8) = CreateCheckSum(bytPacket)
    
    WaitingForData = False
    CreateGetPacket = bytPacket
    
End Function

Private Function SetCommandByte(ByVal bytCmd As eCommand) As Byte
    
    Dim bytNew As Byte
    
    If bytSequence >= 16 Then
        bytSequence = 0
    End If
    
    bytNew = bytCmd + bytSequence
    bytSequence = bytSequence + 1
    SetCommandByte = bytNew

End Function

Public Function MyHex(DecValue As Currency, StringLength As Integer) As String
    
    Dim ByteNum As Integer
    Dim HexString As String
    
    HexString = Hex(DecValue)
'    HexString = CStr(DecValue)
    For ByteNum = Len(HexString) To StringLength - 1
        HexString = "0" & HexString
    Next
    MyHex = HexString

End Function

' MCS
Public Sub GetAllParameters()
        '<EhHeader>
        On Error GoTo GetAllParameters_Err
        '</EhHeader>
    
        '----  Get Parameters from each of the modules
100     Call ViewLog.Log(DebugMsg, "Get All Parameters - CCM")
102     Call ComSend(CreateCommandPacket(eGetAllParameters, CCM_NODE_MASK, 0, 0))
    '    Call Sleep(1000)
    '    Call TimingForm.Delay(1000)
104     Call ViewLog.Log(DebugMsg, "Get All Parameters - Left Front")
106     Call ComSend(CreateCommandPacket(eGetAllParameters, LEFT_FRONT_NODE_MASK, 0, 0))
    '    Call TimingForm.Delay(1000)
108     Call ViewLog.Log(DebugMsg, "Get All Parameters - Right Front")
110     Call ComSend(CreateCommandPacket(eGetAllParameters, RIGHT_FRONT_NODE_MASK, 0, 0))
    '    Call TimingForm.Delay(1000)
112     Call ViewLog.Log(DebugMsg, "Get All Parameters - Left Rear")
114     Call ComSend(CreateCommandPacket(eGetAllParameters, LEFT_REAR_NODE_MASK, 0, 0))
    '    Call TimingForm.Delay(1000)
116     Call ViewLog.Log(DebugMsg, "Get All Parameters - Right Rear")
118     Call ComSend(CreateCommandPacket(eGetAllParameters, RIGHT_REAR_NODE_MASK, 0, 0))
    '    Call TimingForm.Delay(1000)
    
        '<EhFooter>
        Exit Sub

GetAllParameters_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.GetAllParameters " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' MCS
Public Sub GetDiagnosticMessages()
        '<EhHeader>
        On Error GoTo GetDiagnosticMessages_Err
        '</EhHeader>

        ' Clear the Diagnostic Message Text Box
100     MainForm.DiagnosticMessagesText.Text = ""
        ' Clear the Error String
    '    sDiagnosticErrors = ""
        ' Send the command to get the current Errors
102     Call ViewLog.Log(DebugMsg, "Get Error Messages")
104     Call ComSend(CreateCommandPacket(eDataCommand.eDiagnosticMessage, eNodeMask.CCM_NODE_MASK, 0, 0))

        '<EhFooter>
        Exit Sub

GetDiagnosticMessages_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.GetDiagnosticMessages " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' MCS
Public Sub ClearDiagnosticMessages()
        '<EhHeader>
        On Error GoTo ClearDiagnosticMessages_Err
        '</EhHeader>
    
        ' Clear the Diagnostic Message TextBox
100     MainForm.DiagnosticMessagesText.Text = ""
        ' Clear the Error String
    '    sDiagnosticErrors = ""
        ' Send the Command to Clear the Messages
102     Call ViewLog.Log(DebugMsg, "Clear Error Messages")
104     Call ComSend(CreateCommandPacket(eDataCommand.eClearDiagnosticMessage, eNodeMask.CCM_NODE_MASK, 0, 0))

        '<EhFooter>
        Exit Sub

ClearDiagnosticMessages_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.ClearDiagnosticMessages " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function CreateFirmwareDownloadPacket(ByVal DatCmd As String, _
                                            ByVal lStartAddr As Currency, _
                                            ByVal byNodeMask As eNodeMask) As Byte()

'    Dim sMsg As String
    Dim sCmd As String
    Dim intIndex As Integer
    Dim bytPacket() As Byte
    Dim intDataLength As Integer
    Dim strLength As String
    
    ' Convert the Start Address to Hex
    sCmd = MyHex(lStartAddr, 8)             '.ToString("x8")
    
    intDataLength = Len(DatCmd)
    
    ' Convert the Length of the String to Hex
    strLength = MyHex(intDataLength + 10, 4)
    
    ReDim bytPacket(intDataLength + 9) As Byte
    
    ' Command Length
    bytPacket(0) = CByte(Val("&H" & Mid$(strLength, 3, 2)))
    bytPacket(1) = CByte(Val("&H" & Mid$(strLength, 1, 2)))
    ' ID & Sequence
    bytPacket(2) = SetCommandByte(eCommand.Cmd)
    bytPacket(3) = byNodeMask
    ' Command
    bytPacket(4) = 9
'    ' Node
'    bytPacket(5) = byNodeMask
    ' Starting Address
    bytPacket(5) = CByte(Val("&H" & Mid$(sCmd, 7, 2)))
    bytPacket(6) = CByte(Val("&H" & Mid$(sCmd, 5, 2)))
    bytPacket(7) = CByte(Val("&H" & Mid$(sCmd, 3, 2)))
    bytPacket(8) = CByte(Val("&H" & Mid$(sCmd, 1, 2)))
    ' Data
    For intIndex = 1 To intDataLength
        bytPacket(intDataLength - intIndex + 1 + 8) = Asc(Mid$(DatCmd, ((intDataLength - intIndex) + 1), 1))
    Next
    
    bytPacket(intDataLength + 9) = CreateCheckSum(bytPacket)
    
    WaitingForData = True
    CreateFirmwareDownloadPacket = bytPacket

End Function

Private Sub ProcessCommand(ByRef bytCommand() As Byte, intCommandLength As Integer)
        
    On Error GoTo ErrHandler
    
'    Dim bytCommand(256) As Byte
    Dim intByteIndex As Integer
    Dim strMessages() As String
    Dim intIndex As Integer
        
    LastCommMessage = "RX:"
    
    ' Convert the Command into a byte Array
    For intByteIndex = 0 To intCommandLength - 1
        ' MCS Added +1 to intByteIndex
'        bytCommand(intByteIndex) = Asc(Mid(strCommand, intByteIndex + 1, 1))

        LastCommMessage = LastCommMessage & bytCommand(intByteIndex) & " "        'RDR NEEDS HEX CONVERSION   x2
    Next

    LastCommMessage = LastCommMessage & "(" & IncomingCheckSum(bytCommand) & ")"   'RDR NEEDS HEX CONVERSION x2
    
    '---- Parse the protocol
    ' MCS Changed 2 to 3
    ' Changed back to 2 3/26/2003
'    Select Case (Asc(Mid(strCommand, 3, 1)) - (Asc(Mid(strCommand, 3, 1)) And &HF))
    Select Case (bytCommand(2) - (bytCommand(2) And &HF))
        Case eCommand.Ack                   'Acknowledgement Packet
            LastCommMessage = LastCommMessage & "ACK"
            PacketsReceived = PacketsReceived + 1
        Case eCommand.Cmd                   'Command Packet
            Select Case bytCommand(4)            'bDataCmd
                ' Diagnostic Message
                Case eDataCommand.eDiagnosticMessage
                    Dim curMask As Currency              'Each Failure Code is 32 bits
                    Dim sDiagMsg As String
                    Dim iTarget As Integer
'                    Dim strFailure As String
                    
                    LastCommMessage = LastCommMessage & "DIAG"
                    iTarget = 1
                    sDiagMsg = ""
                    MainForm.DiagnosticMessagesText.Text = ""
    
                    '---- Loop through diagnostic code bytes
                    ' MCS change to 34 Vs ubound(bytCommand)
                    For intByteIndex = 5 To intCommandLength - 2 Step 4   'Data starts at 5th byte
                        '---- Compute bitmask
                        ' 5/7/2003 Added currency conversion to avoid overflow error
                        curMask = CCur(bytCommand(intByteIndex)) + (CCur(256) * CCur(bytCommand(intByteIndex + 1)))
                        '---- Loop through the diagnostic codes and check for match
                        sDiagMsg = SetupDiagnosticCodes.GetCode(curMask)
                        
                        ' Don't Display the Left and Right Tag
                        If iTarget < 6 Then
                            If sDiagMsg <> "" Then
                                If InStr(1, sDiagMsg, " ") <> 0 Then
                                    strMessages() = Split(Trim$(sDiagMsg), " ")
                                    For intIndex = 0 To UBound(strMessages)
                                        MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
                                                                            sDiagnosticECU(iTarget) & " - " & strMessages(intIndex) & " Failure " & vbCrLf
                                    Next intIndex
                                Else
                                    MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
                                                                        sDiagnosticECU(iTarget) & " - " & sDiagMsg & " Failure " & vbCrLf
                                End If
                            Else
                                If curMask = 0 Then
                                    MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
                                                                            sDiagnosticECU(iTarget) & " - Passed " & vbCrLf
                                Else
                                    MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
                                                                            sDiagnosticECU(iTarget) & " - Failure Code " & curMask & vbCrLf
                                End If
                            End If
                        End If
                        sDiagMsg = ""
                        iTarget = iTarget + 1
                    Next
                    
                    '--- RDR ADDED ACK
                    Call ComSend(CreateAckPacket(eNodeMask.CCM_NODE_MASK, bytCommand(2)))
                    
                Case eDataCommand.eCCMStartup
                    LastCommMessage = LastCommMessage & "CCM STARTUP"
                    ' 5/7/2003 Added Go get the parameters and Error Codes
                    Call GetAllParameters
                    ' Trying to Clear Application before starting app code
'                    Call SetupSerialPort.ComSend(SetupSerialPort.CreateStartDownloadPacket( _
'                        eDataCommand.eStartFirmwareDownload, CCM_NODE_MASK, 1, _
'                        196608, 242440))
'                    Call SetupSerialPort.ComSend(SetupSerialPort.CreateStartDownloadPacket( _
                        eDataCommand.eStartFirmwareDownload, LEFT_FRONT_NODE_MASK, 1, _
                        196608, 243146))

                Case eDataCommand.eCommandError
                    LastCommMessage = LastCommMessage & "ERROR"
                    PacketErrors = PacketErrors + 1
                Case eDataCommand.eGetParameter
                    LastCommMessage = LastCommMessage & "GET"
                Case eDataCommand.eSetParamater
                    LastCommMessage = LastCommMessage & "SET"
                    Dim iRow As Integer
                    Dim lValue As Currency
'                    Dim blnSkip As Boolean
'
'                    blnSkip = False
                    ' MCS
                    lValue = bytCommand(6) + (256 * CLng(bytCommand(7))) + (65536 * CLng(bytCommand(8))) + (16777216@ * CCur(bytCommand(9)))
                    If lValue > 256@ * 256@ * 256@ Then
                        lValue = lValue - (256@ * 256@ * 256@ * 256@)
                    End If
    
                    With MainForm.ParameterSpread
                        .Col = 4                                'Parameter codes are in column 5
                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
                            .Row = iRow                         'Set active row
                            If Val(.Text) = bytCommand(5) Then      'If the cell value is same as command byte
                                '---- Get the data format
                                .Col = 3
                                Dim iDecimalPlaces As Integer
    
                                iDecimalPlaces = Val(Mid$(.Text, InStr(1, .Text, ".") + 1, 1))
    
                                Select Case bytCommand(3)           'Byte 3 has the target id
                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
                                        .Col = 7
                                    Case eNodeMask.LEFT_REAR_NODE_MASK
                                        .Col = 9
                                    Case eNodeMask.LEFT_TAG_NODE_MASK
                                        .Col = 11
                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
                                        .Col = 8
                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
                                        .Col = 10
                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
                                        .Col = 12
                                    Case eNodeMask.CCM_NODE_MASK
                                        .Col = 13
                                End Select
                                .Text = IIf(iDecimalPlaces > 0, (lValue / (10 ^ iDecimalPlaces)), lValue)
                                Exit For
                            End If
                        Next
                    End With
    
                    '---- Check firmware table
                    With MainForm.ConfigurationSpread
                        .Col = 6                                'Parameter codes are in column 5
                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
                            .Row = iRow                         'Set active row
                            If Val(.Text) = bytCommand(5) Then   'If the cell value is same as command byte
                                Select Case bytCommand(3)           'Byte 3 has the target id
                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
                                        .Col = 2
                                        .Text = lValue
                                        Exit For
                                    Case eNodeMask.LEFT_REAR_NODE_MASK
                                        .Col = 4
                                        .Text = lValue
                                        Exit For
                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
                                        .Col = 3
                                        .Text = lValue
                                        Exit For
                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
                                        .Col = 5
                                        .Text = lValue
                                        Exit For
                                    Case eNodeMask.LEFT_TAG_NODE_MASK
                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
                                    Case eNodeMask.CCM_NODE_MASK
                                    Case Else
                                End Select
                            End If
                        Next
                    End With
    
                Case eDataCommand.eGetAllParameters
                    LastCommMessage = LastCommMessage & "GETALL"
                Case eDataCommand.eBroadcastMessage
    
                    '---- Graphing parameters are received in Broadcast messages
                    LastCommMessage = LastCommMessage & "BCAST"
    
                    Dim slot0 As Long, slot1 As Long, slot2 As Long, slot3 As Long
    
                    slot0 = 256& * CLng(bytCommand(6)) + bytCommand(5)
                    slot1 = 256& * CLng(bytCommand(8)) + bytCommand(7)
                    slot2 = 256& * CLng(bytCommand(10)) + bytCommand(9)
                    slot3 = 256& * CLng(bytCommand(12)) + bytCommand(11)
'                    slot2 = CLng("&h" & Trim$(Hex(bytCommand(10))) & Trim$(Hex(bytCommand(9))))
'                    slot3 = CLng("&h" & Trim$(Hex(bytCommand(12))) & Trim$(Hex(bytCommand(11))))
    
                    Call UpdateGraphs(slot0, slot1, slot2, slot3)
    
                '---- FIRMWARE MESSAGES:
                Case eDataCommand.eDownloadData
                    LastCommMessage = LastCommMessage & "DOWNLOADDATA"
        
                ' Flash Program Status
                ' 9/22/2003 Added error message and status for each node, instead of group status
                Case eDataCommand.eDownloadStatus
                    LastCommMessage = LastCommMessage & "DOWNLOADSTATUS " & NodeName(bytCommand(3))
                    Select Case bytCommand(5)
                        ' OK
                        Case 0
                            LastCommMessage = LastCommMessage & " OK "
                            DownLoadMode(NodeIndex(bytCommand(3))) = eDownLoadMode.ProgramData
                        ' Config Error: Vehicle is moving or command sequence error
                        Case 1
                            LastCommMessage = LastCommMessage & " Config Error "
                            DownLoadStatus(NodeIndex(bytCommand(3))) = CONFIG_ERR
'                            MsgBox "Vehicle is moving or Command " & _
                                    "sequence error", vbCritical, "Config Error"
                        ' Program Address Error: Address in Download Data
                        ' command is invalid
                        Case 6
                            LastCommMessage = LastCommMessage & " Program Address Error "
                            DownLoadStatus(NodeIndex(bytCommand(3))) = ADDRESS_RANGE_ERR
'                            MsgBox "Address in Download Data command is invalid", _
                                    vbInformation, "Program Address Error"
                        ' Program Len Error: Attempt to program data beyond Max
                        ' Address specified
                        Case 7
                            LastCommMessage = LastCommMessage & " Program Length Error "
                            DownLoadStatus(NodeIndex(bytCommand(3))) = PROGRAM_LEN_ERR
'                            MsgBox "Attempt to program data beyond Max Address " & _
                                    "specified", vbCritical, "Program Len Error"
                        ' Program H/W Error: The Flash chip reported an error during
                        ' programming
                        Case 8
                            LastCommMessage = LastCommMessage & " Program H/W Error "
                            DownLoadStatus(NodeIndex(bytCommand(3))) = PROGRAM_HW_ERR
'                            MsgBox "The Flash Chip Reported an error during " & _
                                    "programming", vbCritical, "Program H/W Error"
                    End Select
                    
                    
                    ' Error in Programming Node Remove the node Program list
                    'RDR Subscript Out Of Range Error here...
                    'If bytCommand(5) > 0 Then
                    '    Call RemoveNode(NodeIndex(bytCommand(3)))
                    'End If
                    
                    
                    
                Case eDataCommand.eDownloadEraseStatus
                    LastCommMessage = LastCommMessage & "DOWNLOADERASESTATUS " & NodeName(bytCommand(3))
                    Select Case bytCommand(5)
                        Case 0  ' OK
                            DownLoadMode(NodeIndex(bytCommand(3))) = eDownLoadMode.StartDownload
                            LastCommMessage = LastCommMessage & " OK"
                        Case 1  ' Config Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = CONFIG_ERR
                            LastCommMessage = LastCommMessage & " Config Error"
                        Case 3  ' Erase Command Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = ERASE_CMD_ERR
                            LastCommMessage = LastCommMessage & " Erase Command Error"
                        Case 4  ' Erase Start Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = ERASE_START_ERR
                            LastCommMessage = LastCommMessage & " Erase Start Error"
                        Case 5  ' Erase Timeout Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = ERASE_TIMEOUT_ERR
                            LastCommMessage = LastCommMessage & " Erase Timeout Error"
                        Case 9  ' Memory Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = MEMORY_ERR
                            LastCommMessage = LastCommMessage & " Memory Error"
                    End Select
                    
                    ' Error in Programming Node Remove the node Program list
                    If bytCommand(5) > 0 Then
                        Call RemoveNode(bytCommand(3))
                    End If
                    
                    
                Case eDataCommand.eDownloadComplete
                    LastCommMessage = LastCommMessage & "DOWNLOADCOMPLETE"
                Case eDataCommand.eDownloadVerified
                    LastCommMessage = LastCommMessage & "DOWNLOADVERIFIED " & NodeName(bytCommand(3))
                    Select Case bytCommand(5)
                        Case 0  ' OK
                            DownLoadMode(NodeIndex(bytCommand(3))) = eDownLoadMode.DOWNLOADCOMPLETE
                            LastCommMessage = LastCommMessage & " OK"
                        Case 1  ' Config Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = CONFIG_ERR
                            LastCommMessage = LastCommMessage & " Config Error"
                        Case 2  ' Byte Count Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = BYTE_COUNT_ERR
                            LastCommMessage = LastCommMessage & " Byte Count Error"
                        Case 8  ' Program H/W Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = ERASE_TIMEOUT_ERR
                            LastCommMessage = LastCommMessage & " Program H/W Error"
                        Case 9  ' Memory Error
                            DownLoadStatus(NodeIndex(bytCommand(3))) = MEMORY_ERR
                            LastCommMessage = LastCommMessage & " Memory Error"
                    End Select
                Case eDataCommand.eKonect
                    LastCommMessage = LastCommMessage & "E-CONNECT"
                Case Else
                    LastCommMessage = LastCommMessage & "UNKNOWN"
                    PacketErrors = PacketErrors + 1
            End Select
            PacketsReceived = PacketsReceived + 1
    
            '---- Send ACK
            Call ComSend(CreateAckPacket(eNodeMask.CCM_NODE_MASK, bytCommand(2)))
    
        Case eCommand.Ret                   'Retry Packet
            LastCommMessage = LastCommMessage & "Retry"
            PacketsReceived = PacketsReceived + 1
            
            '---- RDR ADDED ACK
            Call ComSend(CreateAckPacket(eNodeMask.CCM_NODE_MASK, bytCommand(2)))
            
        Case Else
            LastCommMessage = LastCommMessage & "Error"
            PacketErrors = PacketErrors + 1
    End Select

    If InStr(1, UCase$(LastCommMessage), "ERROR") <> 0 Or InStr(1, LastCommMessage, "UNKNOWN") Then
        Call ViewLog.Log(ErrorMsg, LastCommMessage)
    Else
        Call ViewLog.Log(DebugMsg, LastCommMessage)
    End If

    Exit Sub
    
ErrHandler:

    ViewLog.Log ErrorMsg, "Error occurred in ProcessCommand - Error: " & Err.Description

End Sub

Public Function NodeName(bytNode As Byte) As String
        '<EhHeader>
        On Error GoTo NodeName_Err
        '</EhHeader>

100     Select Case bytNode
            Case eNodeMask.LEFT_FRONT_NODE_MASK
102             NodeName = "LEFT FRONT NODE"
104         Case eNodeMask.LEFT_REAR_NODE_MASK
106             NodeName = "LEFT REAR NODE"
108         Case eNodeMask.RIGHT_FRONT_NODE_MASK
110             NodeName = "RIGHT FRONT NODE"
112         Case eNodeMask.RIGHT_REAR_NODE_MASK
114             NodeName = "RIGHT REAR NODE"
116         Case eNodeMask.LEFT_TAG_NODE_MASK
118             NodeName = "LEFT TAG NODE"
120         Case eNodeMask.RIGHT_TAG_NODE_MASK
122             NodeName = "RIGHT TAG NODE"
124         Case eNodeMask.CCM_NODE_MASK
126             NodeName = "CCM NODE"
        End Select

        '<EhFooter>
        Exit Function

NodeName_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.NodeName " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub RemoveNode(bytNode As Byte)
    On Error GoTo ErrorHandler
    
    Dim intNodeIndex As Integer
    
    For intNodeIndex = 1 To 7
        If g_abytNodes(intNodeIndex) = bytNode Then
            ' Clear the Node because of an error
            g_abytNodes(intNodeIndex) = 0
        End If
    Next
ErrorHandler:
End Sub

Public Sub CommReset()
        '<EhHeader>
        On Error GoTo CommReset_Err
        '</EhHeader>

100     Call ConnectCOM

        '<EhFooter>
        Exit Sub

CommReset_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupSerialPort.CommReset " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       NodeIndex
' Description:
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       09/22/2003-11:20:45
'
' Parameters :       bytNode (Byte)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function NodeIndex(bytNode As Byte) As Integer

    Dim intIndex As Integer
    
    For intIndex = 1 To 7
        If g_abytNodes(intIndex) = bytNode Then
            NodeIndex = intIndex
        End If
    Next intIndex

End Function

