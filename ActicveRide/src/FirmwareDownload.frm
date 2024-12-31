VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FirmwareDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firmware Download"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDownloadProgress 
      Caption         =   "Download Progress"
      Height          =   855
      Left            =   1980
      TabIndex        =   10
      Top             =   7320
      Width           =   4335
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Targets"
      Height          =   2655
      Left            =   60
      TabIndex        =   7
      Top             =   4620
      Width           =   6315
      Begin FPSpread.vaSpread vaTarget 
         Height          =   2295
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6075
         _Version        =   393216
         _ExtentX        =   10716
         _ExtentY        =   4048
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "FirmwareDownload.frx":0000
      End
   End
   Begin VB.Frame fraFirmwareFile 
      Caption         =   "Firmware File"
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6315
      Begin VB.Timer TimeOutTimer 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4560
         Top             =   2280
      End
      Begin VB.Frame fraType 
         Caption         =   "Type"
         Height          =   1215
         Left            =   3120
         TabIndex        =   4
         Top             =   180
         Width           =   3075
         Begin VB.OptionButton optType 
            Caption         =   "Application Code"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton optType 
            Caption         =   "Boot Code"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   300
            Width           =   2655
         End
      End
      Begin VB.FileListBox File1 
         ForeColor       =   &H80000007&
         Height          =   1650
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   2895
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   2895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   840
      Left            =   120
      TabIndex        =   9
      Top             =   7380
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1482
      ButtonWidth     =   1429
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
            Caption         =   "Download"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1680
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FirmwareDownload.frx":01D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FirmwareDownload.frx":04F6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FirmwareDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strPreviousDrive As String

Public Enum eDownLoadMode
    Idle = 0
    StartDownload = 1
    ProgramData = 2
    DOWNLOADCOMPLETE = 3
    Waiting = 4
End Enum

Public Enum eDownloadStatus
    STATUS_OK = 0
    CONFIG_ERR = 1
    BYTE_COUNT_ERR = 2
    ERASE_CMD_ERR = 3
    ERASE_START_ERR = 4
    ERASE_TIMEOUT_ERR = 5
    PROGRAM_PAGE_ERR = 6
    PROGRAM_LEN_ERR = 7
    PROGRAM_HW_ERR = 8
    ADDRESS_RANGE_ERR = 9
    MEMORY_ERR = 10
    FLASH_ID_ERR = 11
End Enum

Public g_intNumberNodes As Integer

' Hex File Information
Private m_alngAddress() As Long
Private m_astrData() As String
Private m_astrDataFile() As String
Private m_blnExitLoop As Boolean
Private DownloadTimeOut As Integer
Private Const m_cintBufferSendSize As Integer = 512


Private Sub Dir1_Change()
    
    On Error GoTo ErrHandler
    
    File1.Path = Dir1.Path
    
    Exit Sub

ErrHandler:
    
    Select Case Err.Number
        Case 68
            
            ' Device Not Available
            File1.Path = App.Path
            MsgBox Err.Description
            
    End Select

End Sub

Private Sub Drive1_Change()
        
    On Error GoTo ErrHandler
    
    Dir1.Path = Drive1.Drive
    strPreviousDrive = Drive1.Drive
    
    Exit Sub

ErrHandler:
    
    Select Case Err.Number
        Case 68
            ' Device Not Available
            MsgBox Err.Description
            
            Drive1.Drive = strPreviousDrive
    End Select

End Sub

Private Sub Form_Load()
    
    Dim iRow As Integer

    ' For Exit Do loops
    m_blnExitLoop = False
    
    ' Store Previous Drive
    strPreviousDrive = Drive1.Drive
    ' Default directory
    
    
    SetFirmwarePath
    File1.Pattern = "*.h86"
    
    ' Setup Target SpreadSheet
    With vaTarget
        .LoadTextFile "FirmwareTargets.csv", ",", ",", vbCrLf, LoadTextFileColHeaders, "FirmwareTargets.log"
        .ColWidth(1) = 10
        .ColWidth(2) = 20
        .LockBackColor = vbCyan
        For iRow = 1 To .MaxRows
            .Col = 1
            .Row = iRow
            .CellType = CellTypeCheckBox
            .TypeCheckType = TypeCheckTypeNormal
            .TypeCheckCenter = True
            .Lock = False
            .Col = 2
            .Lock = True
        Next
    End With

End Sub


Private Sub SetFirmwarePath()
    On Error GoTo ErrorHandler
    Dim sPath As String
    sPath = GetSetting(App.CompanyName, App.ProductName, "FirmwareDirectory", App.Path)
    Dir1.Path = sPath
    
    Exit Sub
    
ErrorHandler:
    Dir1.Path = App.Path
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    m_blnExitLoop = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

100     m_blnExitLoop = True
    
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.FirmwareDownload.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub TimeOutTimer_Timer()
    If DownloadTimeOut < 32000 Then
        DownloadTimeOut = DownloadTimeOut + 1
    Else
        DownloadTimeOut = 0
    End If
End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo ToolBar_ButtonClick_Err
        '</EhHeader>

100     Select Case Button.index
            ' Exit
            Case 1
102             m_blnExitLoop = True
104             Unload Me
            ' Download
106         Case 2
108             If File1.ListIndex <> -1 Then
                    ' Save the Directory of the Firmware
110                 SaveSetting App.CompanyName, App.ProductName, "FirmwareDirectory", Dir1.Path
                    ' Validate Firmware File
112                 ValidateFirmwareFile
                Else
114                 MsgBox "Please pick a file for Download", vbInformation + vbOKOnly
                End If
        End Select

        '<EhFooter>
        Exit Sub

ToolBar_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.FirmwareDownload.ToolBar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function ValidateFirmwareFile() As Boolean
    
'    On Error GoTo ErrHandler
    On Error GoTo 0

    Dim sFileName As String
    Dim sBuffer As String
    Dim lngFileNumber As Long
    Dim iByte As Long
    Dim iDatByte As Long
    Dim iLength As Long
    Dim iRecordType As Long
    Dim iAddress As Long
    Dim iSegmentAddress As Long
    Dim lngNumberBytes As Long
    Dim iMaxAddress As Long
    Dim sHexData As String
    Dim intNodeMask As Integer
    Dim intRowIndex As Integer
    Dim lngNumberLines As Long
    Dim lngDataIndex As Long
    Dim intNodeIndex As Integer
    Dim intIndex As Integer
    
    sFileName = ""
    sBuffer = ""
    lngFileNumber = 0
    iByte = 0
    iDatByte = 0
    iLength = 0
    iRecordType = 0
    iAddress = 0
    iSegmentAddress = 0
    lngNumberBytes = 0
    iMaxAddress = 0
    sHexData = ""
    intNodeMask = 0
    lngNumberLines = 0
    lngDataIndex = 0
    intNodeIndex = 0
    intIndex = 0
    lngDataIndex = 0
    m_blnExitLoop = False
    DownloadTimeOut = 0
    TimeOutTimer.Enabled = False
    
    'RDR Verify that a module was selected!
    Dim ModuleSelected As Boolean
    ModuleSelected = False
    With vaTarget
        For intRowIndex = 1 To .MaxRows
            .Row = intRowIndex
            .Col = 1
            If .Value <> 0 Then
                ModuleSelected = True
            End If
        Next
    End With
    
    If ModuleSelected = False Then
        MsgBox "Please select a module.", vbApplicationModal + vbInformation + vbOKOnly, "Download Firmware"
        Exit Function
    End If
    
    sFileName = Dir1.Path & "\" & File1.FileName
    sBuffer = ""

    '---- Read the file into string buffer
    Call ViewLog.Log(InfoMsg, "Opening Firmware File: " & sFileName)
    lngFileNumber = FreeFile
    Open sFileName For Input As #lngFileNumber
    
    ReDim m_alngAddress(1) As Long
    ReDim m_astrData(1) As String
    
    ' Preprocessing the Firmware File
    ' Look for the Starting and Finish Address
    ' Count then number of lines
    ProgressBar1.Max = 5000
    ProgressBar1.Value = 0
    
    ' Read in file information
    Do
        ' Read line from the file
        Input #lngFileNumber, sBuffer
'        lngNumberLines = lngNumberLines + 1
        '---- Check first character of buffer-line
        If Left$(sBuffer, 1) = ":" Then
            iByte = 2
            ' Record Length
            iLength = CLng("&H" & Mid$(sBuffer, iByte, 2))
            iByte = iByte + 2
            ' Record Address
            iAddress = CLng("&H" & Mid$(sBuffer, iByte, 4))
            iByte = iByte + 4
            ' Record Type
            iRecordType = CLng("&H" & Mid$(sBuffer, iByte, 2))
            iByte = iByte + 2

            Select Case iRecordType
                Case 0                      'DATA RECORD
                    ' Increment number Lines
                    lngNumberLines = lngNumberLines + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    ' Redimension the Arrays
                    ReDim Preserve m_alngAddress(lngDataIndex) As Long
                    ReDim Preserve m_astrData(lngDataIndex) As String
                    
                    sHexData = ""
                    
                    ' Store the Data
                    '---- Copy the Hex Bytes into our Byte Array
                    For iDatByte = iByte To iByte + (iLength * 2 - 2) Step 2
                        'bytHexData(lngNumberBytes) = Val("&H" & sBuffer.Substring(iDatByte, 2))
                        sHexData = sHexData & Chr(Val("&H" & Mid$(sBuffer, iDatByte, 2)))
                        lngNumberBytes = lngNumberBytes + 1
                    Next
                    
                    ' Insert into Data Array
                    Call InsertData(iAddress, sHexData)
                
                    ' Increment Data Index
                    lngDataIndex = lngDataIndex + 1
                Case 1                      'END OF FILE RECORD
                    ' Increment number Lines
                    lngNumberLines = lngNumberLines + 1
                    Exit Do
                Case 2                      'Extended SEGMENT RECORD
                    ' Increment number Lines
                    lngNumberLines = lngNumberLines + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                                        
                    iAddress = CLng("&H" & Mid$(sBuffer, iByte, 4))
'                    iSegmentAddress = iAddress * &HFF
                    iSegmentAddress = iAddress * CLng("&H10")

                    If optType(0).Value Then
                        If iAddress = 0 Then 'Boot File
                            If MsgBox("Boot File Selected - Enter OK to Download", vbOKCancel, "Boot File") = vbCancel Then
                                Exit Function
                            End If
                        Else 'Application File
                            If MsgBox("Application File Selected!", vbOKOnly, "Application File") = vbOK Then
                                Exit Function
                            End If
                        End If
                    Else 'application button selected
                        If iAddress = 0 Then 'Boot File
                            If MsgBox("Boot File Selected!", vbOKOnly, "Boot File") = vbOK Then
                                Exit Function
                            End If
                        Else 'Application File

                        End If
                    End If
'                    iByte = iByte + 4
                Case 4                      'EXTENDED LINEAR ADDRESS RECORD
                    lngNumberLines = lngNumberLines + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    
            End Select

            If iAddress >= iMaxAddress Then
                iMaxAddress = iAddress + iLength
            End If
'            iByte = iByte + 3 'need to increments for Carriage Return and Line Feed
        Else
            Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Invalid firmware file.")
            Call MsgBox("Invalid firmware file.", vbCritical + vbOKOnly, "Error")
'            Exit Do
        End If
    Loop Until EOF(lngFileNumber)
    
    ' Close the file
    Close #lngFileNumber
    
    ' Clear Node Array
    For intNodeIndex = 1 To 7
        g_abytNodes(intNodeIndex) = 0
    Next
    
    intNodeIndex = 1
    
    ' Store the Nodes that are to be programmed
    With vaTarget
        For intRowIndex = 1 To .MaxRows
            .Row = intRowIndex
            .Col = 1
            If .Value <> 0 Then
                Select Case .Row
                    Case 1
                        intNodeMask = eNodeMask.CCM_NODE_MASK
                    Case 2
                        intNodeMask = eNodeMask.LEFT_FRONT_NODE_MASK
                    Case 3
                        intNodeMask = eNodeMask.RIGHT_FRONT_NODE_MASK
                    Case 4
                        intNodeMask = eNodeMask.LEFT_REAR_NODE_MASK
                    Case 5
                        intNodeMask = eNodeMask.RIGHT_REAR_NODE_MASK
                    Case 6
                        intNodeMask = eNodeMask.LEFT_TAG_NODE_MASK
                    Case 7
                        intNodeMask = eNodeMask.RIGHT_TAG_NODE_MASK
                End Select
                g_abytNodes(intNodeIndex) = intNodeMask
                ' Increment the Node Index
                intNodeIndex = intNodeIndex + 1
            End If
        Next
    End With
    
'    MsgBox ("Starting download...")
    
    ' Create one big string for downloading the Data
    Call CreateFile
    
    Dim strCodeName As String
    
    ' Application Code
    If optType(1).Value Then
        ' Clear the Application Code Name
        strCodeName = ""
        ' Get the Application Code Name from the Data
        For intIndex = 512 To 517
            strCodeName = strCodeName + m_astrDataFile(intIndex)
        Next
        If g_abytNodes(1) > 1 Then
            ' REM
            If strCodeName <> "REMAPP" Then
                MsgBox "Wrong Application File Picked"
                Exit Function
            End If
        Else
            ' CCM
            If strCodeName <> "CCMAPP" Then
                MsgBox "Wrong Application File Picked"
                Exit Function
            End If
        End If
    End If
    
    ' Set the Current Download status to Idle, Used to know when to send the next
    ' Firmware Packet
    For intIndex = 1 To 7
        DownLoadMode(intIndex) = eDownLoadMode.Idle
        DownLoadStatus(intIndex) = STATUS_OK        'RDR initialize download status!
    Next intIndex
    
    ' Calculate the Max Address
    iMaxAddress = iMaxAddress + iSegmentAddress

    ' Get number of bytes sent
    If g_abytNodes(1) > 1 Then
        ' REM
        lngNumberBytes = UBound(m_astrDataFile)
    Else
        ' CCM
        lngNumberBytes = UBound(m_astrDataFile)
    End If
    
    '---- Send Start Download Command
    ' Hard coded to Application Code - 1
    For intNodeIndex = 1 To 7
        If g_abytNodes(intNodeIndex) <> 0 Then
            DownLoadMode(intNodeIndex) = Idle
'            Call ViewLog.Log(LogMsgTypes.DebugMsg, "Send Start Download Packet " & _
'                            SetupSerialPort.NodeName(g_abytNodes(intNodeIndex)))

            Call SetupSerialPort.ComSend(SetupSerialPort.CreateStartDownloadPacket( _
                        eDataCommand.eStartFirmwareDownload, g_abytNodes(intNodeIndex), 1, _
                        iSegmentAddress, iMaxAddress))
                        
            ' Loop Until the Erase was OK or Error Occurred or User Exits form
'            Do While DownLoadMode(intNodeIndex) <> eDownLoadMode.StartDownload
''            Do While NodesStartDownload
'                If m_blnExitLoop Then
'                    Exit Do
'                End If
'
'                DoEvents
'
'                ' Scan the DownLoadStatus for Errors
''                For intIndex = 1 To 7
'                    '---- Handle start download errors here
'                    If DownLoadStatus(intNodeIndex) = MEMORY_ERR Then
'                        MsgBox "Memory Error has occurred"
'                        DownLoadMode(intNodeIndex) = Idle
'                        Exit Do
'                    End If
''                Next intIndex
'                '---- Handle timeout here
'
'            Loop
        End If
    Next
    
'    ' Loop Until the Erase was OK or Error Occurred or User Exits form

    DownloadTimeOut = 0
    TimeOutTimer.Enabled = True
    Do While Not NodesStartDownload
        If m_blnExitLoop Then
            Exit Do
        End If

        DoEvents

        For intIndex = 1 To 7
            '---- Handle start download errors here
            If DownLoadStatus(intIndex) = MEMORY_ERR Then
                MsgBox "A module memory error has occurred."
                Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Memory Error downloading firmware file.")
                GoTo DownLoadError
            End If
        Next intIndex
        
        '---- Handle timeout here 40 * 100ms = 4sec)
        If DownloadTimeOut > 40 Then
            MsgBox "Download Timed Out (Waiting for erase message).", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
            Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Download timed out (Waiting for erase message).")
            GoTo DownLoadError
        End If
    Loop
    
    TimeOutTimer.Enabled = False
    
    ' Setup Progress Bar for Download of the File
    ProgressBar1.Max = lngNumberBytes \ m_cintBufferSendSize + 2
    ProgressBar1.Value = 0
    
    'Download the File
    If NodesStartDownload Then
        ' Download the Hex File to the CCM or REM(s)
        If DownloadHexFile(iSegmentAddress, intNodeMask) Then
            For intNodeIndex = 1 To 7
                If g_abytNodes(intNodeIndex) <> 0 Then
                    
                    'Send download complete message
                    Call SetupSerialPort.ComSend(SetupSerialPort.CreateDownloadCompletePacket( _
                                eDataCommand.eDownloadComplete, g_abytNodes(intNodeIndex), _
                                lngNumberBytes))
                End If
            Next
        Else
            MsgBox "Error sending firmware."
            Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Error sending firmware.")
            GoTo DownLoadError
        End If
    End If

    ' Look for Download Complete
    DownloadTimeOut = 0
    TimeOutTimer.Enabled = True
    Do While Not NodesDownloadComplete
        If m_blnExitLoop Then
            Exit Do
        End If
        
        DoEvents
        
        ' Scan Nodes for Errors
        For intIndex = 1 To 7
            '---- Handle start download errors here
            If DownLoadStatus(intIndex) <> -1 Then
                If DownLoadStatus(intIndex) <> STATUS_OK Then
                    MsgBox "Error in programming module"
                    Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Error downloading firmware file.")
                    GoTo DownLoadError
                End If
            End If
        Next intIndex
        
        '---- Handle timeout here 40 * 100ms = 4sec)
        If DownloadTimeOut > 40 Then
            MsgBox "Download Timed Out (waiting for complete).", vbApplicationModal + vbExclamation + vbOKOnly, "Error"
            Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Download timed out (waiting for complete).")
            GoTo DownLoadError
        End If
    
    Loop
    DownloadTimeOut = 0
    TimeOutTimer.Enabled = False
    
    ' Download Completed
    If NodesDownloadComplete Then
        ProgressBar1.Value = ProgressBar1.Max
        MsgBox "Download Completed"
    End If
    
    Exit Function

ErrHandler:

    Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Error reading firmware file.")
    Call MsgBox("Error reading firmware file.", vbCritical + vbOKOnly, "Error")
    ProgressBar1.Value = 0
    TimeOutTimer.Enabled = False
    DownloadTimeOut = 0
    Exit Function
    
DownLoadError:
    ProgressBar1.Value = 0
    TimeOutTimer.Enabled = False
    DownloadTimeOut = 0
End Function

Private Sub InsertData(lngAddress As Long, strData As String)

    Dim intIndex As Integer
    Dim intMoveIndex As Integer
    Dim intUpperBound As Integer
    
    intUpperBound = UBound(m_alngAddress)
    
    If intUpperBound = 0 Then
        m_alngAddress(intIndex) = lngAddress
        m_astrData(intIndex) = strData
    Else
        For intIndex = 0 To intUpperBound - 1
            If m_alngAddress(intIndex) < lngAddress Then
                ' Keep Looking
            Else
                For intMoveIndex = intUpperBound To intIndex + 1 Step -1
                    m_alngAddress(intMoveIndex) = m_alngAddress(intMoveIndex - 1)
                    m_astrData(intMoveIndex) = m_astrData(intMoveIndex - 1)
                Next intMoveIndex
                ' Jump out of the For to save the Data Sent into Sub
                Exit For
            End If
        Next
        m_alngAddress(intIndex) = lngAddress
        m_astrData(intIndex) = strData
    End If

End Sub




Private Function DownloadHexFile(lngOffset As Long, intNodeMask As Integer) As Boolean

    Dim strDownloadString As String
    Dim intLength As Integer
    Dim lngDataIndex As Long
    Dim lngUpperBound As Long
    Dim lngAddress As Long
    Dim intNodeIndex As Integer
    Dim intStart As Integer
    Dim intIndex As Integer
    
    DownloadHexFile = False
    
    lngUpperBound = UBound(m_astrDataFile)
    
    intLength = 1
    lngAddress = 0
    
    If g_abytNodes(1) > 1 Then
        ' REMs
        lngAddress = lngOffset
        intStart = 0
    Else
        ' CCM
        lngAddress = lngOffset
        intStart = 0
    End If
    
    For lngDataIndex = intStart To lngUpperBound - 1
        strDownloadString = strDownloadString + m_astrDataFile(lngDataIndex)
        intLength = intLength + 1
        If intLength > m_cintBufferSendSize Or lngDataIndex = lngUpperBound - 1 Then
            Call ViewLog.Log(DebugMsg, "Last Address = " & lngAddress - lngOffset & _
                                " Line = " & lngDataIndex & " Length = " & Len(strDownloadString))
            For intNodeIndex = 1 To 7
                If g_abytNodes(intNodeIndex) <> 0 Then
                    ' Send the Data to CCM or REM
                    Call SetupSerialPort.ComSend(SetupSerialPort.CreateFirmwareDownloadPacket( _
                                                 strDownloadString, lngAddress, g_abytNodes(intNodeIndex)))
                    DownLoadMode(intNodeIndex) = Waiting
                End If
            Next
            intLength = 1
            strDownloadString = ""
'            If lngUpperBound - lngDataIndex > m_cintBufferSendSize Then
                lngAddress = lngAddress + m_cintBufferSendSize
'            End If
            ' Make sure packet was sent and is good
            DownloadTimeOut = 0
            TimeOutTimer.Enabled = True
            Do
                ' Scan Nodes for Status
                For intIndex = 1 To 7
                    If g_abytNodes(intIndex) <> 0 Then
                        If DownLoadStatus(intIndex) = STATUS_OK And DownLoadMode(intIndex) = ProgramData Then
                            DownLoadMode(intIndex) = Idle
    '                        Exit Do
                        ElseIf DownLoadStatus(intIndex) <> STATUS_OK Then
                            MsgBox "Error in Program Unit " & SetupSerialPort.NodeName(g_abytNodes(intIndex)) & vbCrLf & "Error Number " & _
                                        DownLoadStatus(intIndex)
                            DownLoadMode(intIndex) = Idle
                            SetupSerialPort.RemoveNode (g_abytNodes(intIndex))
                            
                            'RDR - Added code to get out of loop!
                            DownloadTimeOut = 0
                            TimeOutTimer.Enabled = False
                            Exit Function
                            
                        Else
                            If m_blnExitLoop Then
                                Exit Function
                            End If
                            DoEvents
                        End If
                    End If
                Next intIndex
                
                '---- RDR Added Timeout timer
                If DownloadTimeOut > 40 Then
                    DownloadTimeOut = 0
                    TimeOutTimer.Enabled = False
                    Exit Function
                End If
                
                
            Loop Until NodesIdle
            DownloadTimeOut = 0
            TimeOutTimer.Enabled = False
            
            ' Display the Progress
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
    Next
    
    DownloadHexFile = True

End Function

Private Sub CreateFile()
        '<EhHeader>
        On Error GoTo CreateFile_Err
        '</EhHeader>

        Dim lngIndex As Long
        Dim intDataIndex As Integer
        Dim lngAddress As Long
        Dim lngLength As Long
        
104     For intDataIndex = 0 To UBound(m_astrData)
106         lngAddress = m_alngAddress(intDataIndex)
108         lngLength = Len(m_astrData(intDataIndex))
        
            ' File Data information
110         ReDim Preserve m_astrDataFile(lngAddress + lngLength) As String
112         For lngIndex = lngAddress To lngAddress + lngLength - 1
114             m_astrDataFile(lngIndex) = Mid$(m_astrData(intDataIndex), lngIndex - lngAddress + 1, 1)
116         Next lngIndex
118     Next intDataIndex
    
        ' Fill in the blanks
120     For lngIndex = 1 To UBound(m_astrDataFile)
122         If Len(m_astrDataFile(lngIndex)) = 0 Then
124             m_astrDataFile(lngIndex) = Chr$(Val("&HFF"))
            End If
        Next
    
        '<EhFooter>
        Exit Sub

CreateFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.FirmwareDownload.CreateFile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub vaTarget_Change(ByVal Col As Long, ByVal Row As Long)
        '<EhHeader>
        On Error GoTo vaTarget_Change_Err
        '</EhHeader>

        Dim intIndex As Integer
        Dim blnChecked As Boolean
    
100     With vaTarget
102         If Col = 1 And Row = 1 Then
104             .Col = Col
106             .Row = Row
108             If .Text = 1 Then
110                 For intIndex = 2 To 6
112                     .Row = intIndex
114                     .Lock = True
                    Next
                Else
116                 For intIndex = 2 To 6
118                     .Row = intIndex
120                     .Lock = False
                    Next
                End If
            Else
122             .Col = Col
124             .Row = Row
126             If .Text = 1 Then
128                 For intIndex = 2 To 6
130                     .Row = intIndex
132                     If .Text = 1 Then
134                         .Row = 1
136                         .Lock = True
                            Exit For
                        End If
                    Next
                Else
138                 blnChecked = False
140                 For intIndex = 2 To 6
142                     .Row = intIndex
144                     If Val(.Text) = 0 Then
146                         blnChecked = blnChecked Or False
                        Else
148                         blnChecked = blnChecked Or True
                        End If
                    Next
150                 If blnChecked Then
                
                    Else
152                     .Row = 1
154                     .Lock = False
                    End If
                End If
            End If
        End With

        '<EhFooter>
        Exit Sub

vaTarget_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.FirmwareDownload.vaTarget_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       NodesIdle
' Description:
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       09/22/2003-10:36:48
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NodesIdle() As Boolean

    Dim intIndex As Integer
    Dim blnIdle As Boolean
    
    blnIdle = True
    
    ' Scan Nodes for Status
    For intIndex = 1 To 7
        ' Check to see if Node is being programmed
        If g_abytNodes(intIndex) <> 0 Then
            If DownLoadMode(intIndex) = Idle Then
                blnIdle = True
            Else
                blnIdle = False
                Exit For
            End If
        End If
    Next intIndex

    NodesIdle = blnIdle

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       NodesStartDownload
' Description:
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       09/22/2003-10:48:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NodesStartDownload() As Boolean

    Dim intIndex As Integer
    Dim blnStart As Boolean
    
    For intIndex = 1 To 7
        ' Check to see if Node is being programmed
        If g_abytNodes(intIndex) <> 0 Then
            If DownLoadMode(intIndex) <> StartDownload Then
                blnStart = False
                Exit For
            Else
                blnStart = True
            End If
        End If
    Next intIndex

    NodesStartDownload = blnStart

End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       NodesDownloadComplete
' Description:
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       09/22/2003-11:03:29
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NodesDownloadComplete() As Boolean

    Dim intIndex As Integer
    Dim blnComplete As Boolean
    
    For intIndex = 1 To 7
        ' Check to see if Node is DownloadComplete
        If g_abytNodes(intIndex) <> 0 Then
            If DownLoadMode(intIndex) = DOWNLOADCOMPLETE Then
                blnComplete = True
            Else
                blnComplete = False
                Exit For
            End If
        End If
    Next intIndex

    NodesDownloadComplete = blnComplete

End Function

