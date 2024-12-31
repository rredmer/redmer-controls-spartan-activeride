Attribute VB_Name = "OldCode"
'Private Sub CommPort_OnComm()
'
'    Dim sMsg As String
'    Dim intIndex As Integer
'
'    If InCommRoutine = True Then
'        Exit Sub
'    End If
'
'    InCommRoutine = True
'
'    'CommPort.RThreshold = 0
'
'    Select Case CommPort.CommEvent
''        Case comEventBreak                          'A Break was received.
''        Case comEventFrame                          'Framing Error
''            PacketErrors += 1
''            Beep()
''        Case comEventOverrun                        'Data Lost
''            PacketErrors += 1
''            Beep()
''        Case comEventRxOver                         'Receive buffer overflow
''            PacketErrors += 1
''            Beep()
''        Case comEventRxParity                       'Parity Error
''            PacketErrors += 1
''            Beep()
''        Case comEventTxFull                         'Transmit buffer full
''            PacketErrors += 1
''            Beep()
''        Case comEventDCB                            'Unexpected error retrieving DCB
'        Case comEvCD                                'Change in the CD line
'            ViewLog.Log DebugMsg, "CD Line Event"
''        Case comEvCTS                               'Change in the CTS line
''        Case comEvDSR                               'Change in the DSR line
''        Case comEvRing                              'Change in the Ring Indicator
'        Case comEvReceive                           'Received RThreshold # of chars
'            If CommPort.InBufferCount > 0 Then        'If there are characters in the serial buffer
'
'                Dim sBuf As String
'                sBuf = ""
'                sBuf = CommPort.Input
'                CommBuffer = CommBuffer + sBuf
'                If Len(CommBuffer) >= 2 And PacketHeaderReceived = False Then
''                    PacketHeaderReceived = True
''                    CommBufferLen = Asc(Mid(CommBuffer, 1, 1)) + Asc(Mid(CommBuffer, 2, 1))
'                    ' MCS scan for valid lengths
'                    For intIndex = 1 To Len(CommBuffer) - 2
'                        CommBufferLen = Asc(Mid(CommBuffer, intIndex, 1)) + _
'                                                Asc(Mid(CommBuffer, intIndex + 1, 1))
'                        If CommBufferLen = 11 Or CommBufferLen = 34 Or CommBufferLen = 4 Or _
'                                CommBufferLen = 14 Or CommBufferLen = 6 Or CommBufferLen = 8 Then
'                            CommBuffer = Mid$(CommBuffer, intIndex, CommBufferLen)
'                            PacketHeaderReceived = True
'                            Exit For
'                        Else
'                            CommBufferLen = 0
'                        End If
'                    Next intIndex
'                End If
'                If Len(CommBuffer) >= CommBufferLen And CommBufferLen > 0 And PacketHeaderReceived = True Then
'                    Dim sCurMsg As String
'                    ' MCS changed from 1 to 2
'                    sCurMsg = Mid(CommBuffer, 1, CommBufferLen)
'                    Dim iLength As Integer
'                    ' MCS removed -1
'                    iLength = Len(sCurMsg)
'                    If Len(CommBuffer) > CommBufferLen Then
'                        CommBuffer = Mid(CommBuffer, CommBufferLen, Len(CommBuffer))
'                    Else
'                        CommBuffer = ""
'                    End If
'                    PacketHeaderReceived = False
'
'                    '---- Convert string message to hex bytes
'                    'Dim ByteArray(iLength) As Byte
'
'                    Dim ByteArray(256) As Byte
'                    Dim iByte As Integer
'
'                    LastCommMessage = LastCommMessage & "RX:"
'                    For iByte = 0 To iLength - 1
'                        ' MCS Added +1 to ibyte
'                        ByteArray(iByte) = Asc(Mid(sCurMsg, iByte + 1, 1))
'
'                        LastCommMessage = LastCommMessage & ByteArray(iByte) & " "        'RDR NEEDS HEX CONVERSION   x2
'                    Next
'
'                    LastCommMessage = LastCommMessage & "(" & IncomingCheckSum(ByteArray) & ")"   'RDR NEEDS HEX CONVERSION x2
'
'                    '---- Parse the protocol
'                    ' MCS Changed 2 to 3
'                    Select Case (Asc(Mid(sCurMsg, 3, 1)) - (Asc(Mid(sCurMsg, 3, 1)) And &HF))
'                        Case eCommand.Ack                   'Acknowledgement Packet
'                            LastCommMessage = LastCommMessage & "ACK"
'                            PacketsReceived = PacketsReceived + 1
'                        Case eCommand.Cmd                   'Command Packet
'                            Select Case ByteArray(4)            'bDataCmd
'                                ' Diagnostic Message
'                                Case eDataCommand.eDiagnosticMessage
'                                    LastCommMessage = LastCommMessage & "DIAG"
'                                    Dim iTarget As Integer
'                                    iTarget = 1
'                                    Dim iMask As Integer                          'Each Failure Code is 32 bits
'                                    Dim sDiagMsg As String
'                                    sDiagMsg = ""
'                                    MainForm.DiagnosticMessagesText.Text = ""
'
'                                    '---- Loop through diagnostic code bytes
'                                    ' MCS change to 34 vs ubound(bytearray)
'                                    For iByte = 5 To 34 - 2 Step 4   'Data starts at 5th byte
'                                        '---- Compute bitmask
'                                        iMask = ByteArray(iByte) + (256 * ByteArray(iByte + 1))
'                                        '---- Loop through the diagnostic codes and check for match
'                                        sDiagMsg = sDiagMsg + SetupDiagnosticCodes.GetCode(iMask)
'                                        ' Don't Display the Left and Right Tag
'                                        If iTarget < 6 Then
'                                            If sDiagMsg <> "" Then
'                                                MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
'                                                                                        sDiagnosticECU(iTarget) & " - " & sDiagMsg & " Failure " & vbCrLf
'                                            Else
'                                                MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
'                                                                                        sDiagnosticECU(iTarget) & " - Passed " & vbCrLf
'                                            End If
'                                        End If
'                                        sDiagMsg = ""
'                                        iTarget = iTarget + 1
'                                    Next
'                                Case eDataCommand.eCCMStartup
'                                    LastCommMessage = LastCommMessage & "CCM STARTUP"
'                                Case eDataCommand.eCommandError
'                                    LastCommMessage = LastCommMessage & "ERROR"
'                                    PacketErrors = PacketErrors + 1
'                                Case eDataCommand.eGetParameter
'                                    LastCommMessage = LastCommMessage & "GET"
'                                Case eDataCommand.eSetParamater
'                                    LastCommMessage = LastCommMessage & "SET"
'                                    Dim iRow As Integer
'                                    Dim lValue As Currency
'                                    Dim blnSkip As Boolean
'
'                                    blnSkip = False
'                                    ' MCS
'                                    lValue = ByteArray(6) + (256 * CLng(ByteArray(7))) + (65536 * CLng(ByteArray(8))) + (16777216@ * CCur(ByteArray(9)))
'                                    If lValue > 256@ * 256@ * 256@ Then
'                                        lValue = lValue - (256@ * 256@ * 256@ * 256@)
'                                    End If
'
'                                    With MainForm.ParameterSpread
'                                        .Col = 5                                'Parameter codes are in column 5
'                                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
'                                            .Row = iRow                         'Set active row
'                                            If Val(.Text) = ByteArray(5) Then      'If the cell value is same as command byte
'                                                '---- Get the data format
'                                                .Col = 4
'                                                Dim iDecimalPlaces As Integer
'
'                                                iDecimalPlaces = Val(Mid(.Text, InStr(1, .Text, ".") + 1, 1))
'
'                                                Select Case ByteArray(3)           'Byte 3 has the target id
'                                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
'                                                        .Col = 8
'                                                    Case eNodeMask.LEFT_REAR_NODE_MASK
'                                                        .Col = 10
'                                                    Case eNodeMask.LEFT_TAG_NODE_MASK
'                                                        .Col = 12
'                                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
'                                                        .Col = 9
'                                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
'                                                        .Col = 11
'                                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
'                                                        .Col = 13
'                                                    Case eNodeMask.CCM_NODE_MASK
'                                                        .Col = 14
'                                                End Select
'                                                .Text = IIf(iDecimalPlaces > 0, (lValue / (10 ^ iDecimalPlaces)), lValue)
'                                                Exit For
'                                            End If
'                                        Next
'                                    End With
'
'                                    '---- Check firmware table
'                                    With MainForm.ConfigurationSpread
'                                        .Col = 6                                'Parameter codes are in column 5
'                                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
'                                            .Row = iRow                         'Set active row
'                                            If Val(.Text) = ByteArray(5) Then   'If the cell value is same as command byte
'                                                Select Case ByteArray(3)           'Byte 3 has the target id
'                                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
'                                                        .Col = 2
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.LEFT_REAR_NODE_MASK
'                                                        .Col = 4
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
'                                                        .Col = 3
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
'                                                        .Col = 5
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.LEFT_TAG_NODE_MASK
'                                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
'                                                    Case eNodeMask.CCM_NODE_MASK
'                                                    Case Else
'                                                End Select
'                                            End If
'                                        Next
'                                    End With
'
'                                Case eDataCommand.eGetAllParameters
'                                    LastCommMessage = LastCommMessage & "GETALL"
'                                Case eDataCommand.eBroadcastMessage
'
'                                    '---- Graphing parameters are received in Broadcast messages
'                                    LastCommMessage = LastCommMessage & "BCAST"
'
'                                    Dim slot0 As Long, slot1 As Long, slot2 As Long, slot3 As Long
'
'                                    slot0 = "&h" & Trim(Hex(ByteArray(6))) & Trim(Hex(ByteArray(5)))
'                                    slot1 = "&h" & Trim(Hex(ByteArray(8))) & Trim(Hex(ByteArray(7)))
'                                    slot2 = "&h" & Trim(Hex(ByteArray(10))) & Trim(Hex(ByteArray(9)))
'                                    slot3 = "&h" & Trim(Hex(ByteArray(12))) & Trim(Hex(ByteArray(11)))
'
'                                    Call UpdateGraphs(slot0, slot1, slot2, slot3)
'
'
'                                '---- FIRMWARE MESSAGES:
'                                Case eDataCommand.eDownloadData
'                                    LastCommMessage = LastCommMessage & "DOWNLOADDATA"
'
'                                Case eDataCommand.eDownloadStatus
'                                    LastCommMessage = LastCommMessage & "DOWNLOADSTATUS"
'
'                                Case eDataCommand.eDownloadEraseStatus
'                                    LastCommMessage = LastCommMessage & "DOWNLOADERASESTATUS"
'                                    Select Case ByteArray(5)
'                                        Case 0  ' OK
'                                            FirmwareDownload.DownLoadMode = eDownLoadMode.StartDownload
'                                        Case 1  ' Config Error
'                                            FirmwareDownload.DownLoadStatus = CONFIG_ERR
'                                        Case 3  ' Erase Command Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_CMD_ERR
'                                        Case 4  ' Erase Start Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_START_ERR
'                                        Case 5  ' Erase Timeout Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_TIMEOUT_ERR
'                                        Case 9  ' Memory Error
'                                            FirmwareDownload.DownLoadStatus = MEMORY_ERR
'                                    End Select
'                                Case eDataCommand.eDownloadComplete
'                                    LastCommMessage = LastCommMessage & "DOWNLOADCOMPLETE"
'                                Case eDataCommand.eDownloadVerified
'                                    LastCommMessage = LastCommMessage & "DOWNLOADVERIFIED"
'                                Case eDataCommand.eKonect
'                                    LastCommMessage = LastCommMessage & "E-CONNECT"
'                                Case Else
'                                    LastCommMessage = LastCommMessage & "UNKNOWN"
'                                    PacketErrors = PacketErrors + 1
'                            End Select
'                            PacketsReceived = PacketsReceived + 1
'
'                            '---- Send ACK
'                            Call ComSend(CreateAckPacket(eNodeMask.CCM_NODE_MASK, ByteArray(2)))
'
'                        Case eCommand.Ret                   'Retry Packet
'                            LastCommMessage = LastCommMessage & "Retry"
'                            PacketsReceived = PacketsReceived + 1
'                        Case Else
'                            LastCommMessage = LastCommMessage & "Error"
'                            PacketErrors = PacketErrors + 1
'                    End Select
'                    If InStr(1, UCase(LastCommMessage), "ERROR") <> 0 Or InStr(1, LastCommMessage, "UNKNOWN") Then
'                        Call ViewLog.Log(ErrorMsg, LastCommMessage)
'                    Else
'                        Call ViewLog.Log(DebugMsg, LastCommMessage)
'                    End If
'                    Call MainForm.UpdateStatus
'                    LastCommMessage = ""
'
'                Else
'                    ' Added MCS 1/23/2003 see what's being dumped
''                    Dim sMsg As String
''                    Dim intIndex As Integer
'
'                    sMsg = ""
'                    For intIndex = 1 To Len(sBuf)
'                        sMsg = sMsg & (Asc((Mid(sBuf, intIndex, 1)))) & " "
'                    Next
'                    ViewLog.Log ErrorMsg, "Bad Data " & sMsg
'
'                End If
'                WaitingForData = False
'                CommStatus = "Port Open."
'                CommTimeOut = 0
'            Else
'            End If
'
'            Case comEvSend                              'SThreshold number of characters in the transmit buffer
'            Case comEvEOF                               'An EOF charater was found in the input stream
'            Case Else
'                ViewLog.Log DebugMsg, "Comm event (" & Trim(Str(CommPort.CommEvent)) & ") triggered."
'        End Select
'
'    'CommPort.RThreshold = 1
'    InCommRoutine = False
'
'End Sub


' my version of the OnComm
'Private Sub CommPort_OnComm()
'
'    Dim strCommand As String
'    Dim intIndex As Integer
'    Dim intBufferLength As Integer
'    Dim bytArray(256) As Byte
'    Dim intByteIndex As Integer
'
'    Select Case CommPort.CommEvent
'        Case comEvReceive
'            ' Get the Comm Buffer
'            If m_blnHeaderFound Then
'                m_strBuffer = m_strBuffer & CommPort.Input
'            Else
'                m_strBuffer = CommPort.Input
'            End If
'
'            ' Check for length
'            If Len(m_strBuffer) >= 2 Then
'                ' Look for Header
'                If Not m_blnHeaderFound Then
'                    ' Check the command Length
'                    For intIndex = 1 To Len(m_strBuffer) - 2
'                        m_intCommandLength = Asc(Mid$(m_strBuffer, intIndex, 1)) + _
'                                            Asc(Mid$(m_strBuffer, intIndex + 1, 1))
'                        If m_intCommandLength = 11 Or m_intCommandLength = 34 Or _
'                                m_intCommandLength = 4 Or m_intCommandLength = 14 Or _
'                                m_intCommandLength = 6 Or m_intCommandLength = 8 Then
'                             ' Found a Good Command Length
'                             m_blnHeaderFound = True
'                             Exit For
'                        End If
'                    Next intIndex
'                End If
'
'                If m_blnHeaderFound Then
'                    intBufferLength = Len(m_strBuffer)
'                    ' Trim Extra buffer information
'                    If intIndex <> 1 Then
'                        m_strBuffer = Right$(m_strBuffer, intBufferLength - intIndex + 1)
'                        ' New buffer length
'                        intBufferLength = Len(m_strBuffer)
'                    End If
'
'                    ' Grab command if buffers long enough
'                    If intBufferLength > m_intCommandLength Then
'                        ' Grab the Command
'                        strCommand = Left$(m_strBuffer, m_intCommandLength)
'                        ' Save the remaining Buffer
'                        m_strBuffer = Right$(m_strBuffer, intBufferLength - m_intCommandLength)
'                    ElseIf intBufferLength = m_intCommandLength Then
'                        strCommand = m_strBuffer
'                    Else
'                        ' Go Wait for more data
'                        Exit Sub
'                    End If
'
'                    Call ProcessCommand(strCommand, m_intCommandLength)
'                    m_blnHeaderFound = False
'                    WaitingForData = False
'                    CommStatus = "Port Open"
'                    CommTimeOut = 0
'                Else
'                    ' Clear buffer
'                    m_strBuffer = ""
'                    WaitingForData = False
'                    CommStatus = "Port Open"
'                    CommTimeOut = 0
'                End If
'            End If
'            WaitingForData = False
'            CommStatus = "Port Open"
'            CommTimeOut = 0
''        Case comEventBreak                          'A Break was received.
''        Case comEventFrame                          'Framing Error
''            PacketErrors += 1
''            Beep()
''        Case comEventOverrun                        'Data Lost
''            PacketErrors += 1
''            Beep()
''        Case comEventRxOver                         'Receive buffer overflow
''            PacketErrors += 1
''            Beep()
''        Case comEventRxParity                       'Parity Error
''            PacketErrors += 1
''            Beep()
''        Case comEventTxFull                         'Transmit buffer full
''            PacketErrors += 1
''            Beep()
''        Case comEventDCB                            'Unexpected error retrieving DCB
'        Case comEvCD                                'Change in the CD line
'            ViewLog.Log DebugMsg, "CD Line Event"
''        Case comEvCTS                               'Change in the CTS line
''        Case comEvDSR                               'Change in the DSR line
''        Case comEvRing                              'Change in the Ring Indicator
'        Case comEvSend                              'SThreshold number of characters in the transmit buffer
'        Case comEvEOF                               'An EOF charater was found in the input stream
'        Case Else
'            ViewLog.Log DebugMsg, "Comm event (" & Trim(Str(CommPort.CommEvent)) & ") triggered."
'    End Select
'
'End Sub

' Firmware download code
'    ' Get the Starting Address
'    Input #lngFileNumber, sBuffer
'
'    ' Starting Address
'    strStartingAddress = Mid(sBuffer, 2, 7)
'    strEndingAddress = Mid(sBuffer, 9, 14)
'
'    ' Application File Found
'    If Val(strStartingAddress) = 200000 Then
'        ' Application File Picked
'        If optType(1).Value Then
'            ' Continue
'        Else
'            MsgBox "File Selected isn't a Boot File"
'            Close #lngFileNumber
'            Exit Function
'        End If
'    ' Boot file Found
'    Else
'        If optType(1).Value Then
'            MsgBox "File Selected isn't an Application File"
'            Close #lngFileNumber
'            Exit Function
'        Else
'            ' Continue
'        End If
'    End If


' Backup Firmware download code
'Private Function ValidateFirmwareFile() As Boolean
'
''    Dim oFile As FileStream                     'Pointer to Firmware file
''    Dim oStream As StreamReader                 'Pointer to Firmware stream reader
'    Dim sFileName As String
'    Dim sBuffer As String
'    Dim bIsValid As Boolean
'    Dim lngFileNumber As Long
'
''    On Error GoTo ErrHandler
'
'    sFileName = Dir1.Path & "\" & File1.FileName
'    sBuffer = ""
'    bIsValid = False
'
'    '---- Read the file into string buffer
'    Call ViewLog.Log(InfoMsg, "Opening Firmware File: " & sFileName)
'
'    ' Get Next File Number
'    lngFileNumber = FreeFile
'
'    ' Open File
'    Open sFileName For Input As #1
''    Call ViewLog.Log(LogMsgTypes.InfoMsg, "Read: " & sBuffer.Length & " bytes from Firmware file.")
'
'    '---- Validate data and copy into byte buffer
'    Dim iByte As Long
'    Dim iDatByte As Long
'    Dim iLength As Long
'    Dim iRecordType As Long
'    Dim iAddress As Long
'    Dim iSegmentAddress As Long
'    Dim iHexByte As Long
'    Dim iMaxAddress As Long
'    Dim sHexData As String
'    Dim strStartingAddress As String
'    Dim strEndingAddress As String
'    Dim intNodeMask As Integer
'    Dim intRowIndex As Integer
'    Dim strHexAdress As String
'
'    iByte = 0
'    iDatByte = 0
'    iLength = 0
'    iRecordType = 0
'    iAddress = 0
'    iSegmentAddress = 0
'    iHexByte = 0
'    iMaxAddress = 0
'
'    ' Read in file information
''    For iByte = 0 To sBuffer.Length - 1     'Loop for each byte in the Hex File
'    Do
'        ' Read line from the file
'        Input #lngFileNumber, sBuffer
'
'        '---- Check first character of buffer-line
'        If Left$(sBuffer, 1) = ":" Then
'            iByte = 2
'            ' Record Length
'            iLength = CLng("&H" & Mid(sBuffer, iByte, 2))
'            iByte = iByte + 2
'            ' was 4
'            ' Record Address
'            iAddress = CLng("&H" & Mid(sBuffer, iByte, 4))
'            iByte = iByte + 4
'            ' Record Type
'            iRecordType = CLng("&H" & Mid(sBuffer, iByte, 2))
'            iByte = iByte + 2
'
'            Select Case iRecordType
'                Case 0                      'DATA RECORD
'                    '---- Copy the Hex Bytes into our Byte Array
'                    For iDatByte = iByte To iByte + (iLength * 2 - 2) Step 2
'                        'bytHexData(iHexByte) = Val("&H" & sBuffer.Substring(iDatByte, 2))
'                        sHexData = sHexData & Chr(Val("&H" & Mid(sBuffer, iDatByte, 2)))
'                        iHexByte = iHexByte + 1
'                        ' Remove MCS 3/19/2003
'                        'iByte = iByte + 2
'                    Next
'                Case 1                      'END OF FILE RECORD
'                    Exit Do
'                Case 2                      'Extended SEGMENT RECORD
'                    iAddress = CLng("&H" & Mid(sBuffer, iByte, 2))
'                    iSegmentAddress = iAddress * &HFF
'
''                    If iAddress <> 0 Then
''                        Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Invalid firmware file - 02 Record Type Failure.")
''                        Call MsgBox("Invalid firmware file.", vbCritical + vbOKOnly, "Error")
''                        Exit Do
''                    End If
'                    If optType(0).Value Then
'                        If iAddress = 0 Then 'Boot File
'                            If MsgBox("Boot File Selected - Enter OK to Download", vbOKCancel, "Boot File") = vbCancel Then
'                                Exit Function
'                            End If
'                        Else 'Aplication File
'                            If MsgBox("Application File Selected!", vbOKOnly, "Application File") = vbOK Then
'                                Exit Function
'                            End If
'                        End If
'                    Else 'application button selected
'                        If iAddress = 0 Then 'Boot File
'                            If MsgBox("Boot File Selected!", vbOKOnly, "Boot File") = vbOK Then
'                                Exit Function
'                            End If
'                        Else 'Aplication File
'
'                        End If
'                    End If
''                    iSegmentAddress = iAddress * &HFF
'                    'joe save iAddress to Segment Address
'                    'joe determine if application or boot code and if does not match button selection, prompt warning
'                    iByte = iByte + 4
'                Case 4                      'EXTENDED LINEAR ADDRESS RECORD
'
'            End Select
'
'            If (iAddress + iLength) > iMaxAddress Then
'                iMaxAddress = iAddress + iLength
'            End If
'
'            '---- Calculate the checksum joe
'            iByte = iByte + 3 'need to increments for Carriage Return and Line Feed
'        Else
'            Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Invalid firmware file.")
'            Call MsgBox("Invalid firmware file.", vbCritical + vbOKOnly, "Error")
'            Exit Do
'        End If
'    Loop Until EOF(lngFileNumber)
'
''    Next
'    With vaTarget
'        For intRowIndex = 1 To .MaxRows
'            .Row = intRowIndex
'            .Col = 1
'            If .Value <> 0 Then
'                Select Case .Row
'                    Case 1
'                        intNodeMask = eNodeMask.CCM_NODE_MASK
'                    Case 2
'                        intNodeMask = eNodeMask.LEFT_FRONT_NODE_MASK
'                    Case 3
'                        intNodeMask = eNodeMask.RIGHT_FRONT_NODE_MASK
'                    Case 4
'                        intNodeMask = eNodeMask.LEFT_REAR_NODE_MASK
'                    Case 5
'                        intNodeMask = eNodeMask.RIGHT_REAR_NODE_MASK
'                    Case 6
'                        intNodeMask = eNodeMask.LEFT_TAG_NODE_MASK
'                    Case 7
'                        intNodeMask = eNodeMask.RIGHT_TAG_NODE_MASK
'                End Select
'            End If
'        Next intRowIndex
'    End With
'
'    ' Close the file
'    Close #lngFileNumber
'
'    MsgBox ("Starting download...")
'    'joe continue to look for positive acknowledgment that erase has occured Oc 00
'    '07 00 a6 02 0c 00 ef
'    'Change DownloadMode to eDownloadMode.???
'    DownLoadMode = eDownLoadMode.Idle
''
'    '---- Send Start Download Command
'    Call SetupSerialPort.ComSend(SetupSerialPort.CreateStartDownloadPacket(eDataCommand.eStartFirmwareDownload, intNodeMask, 1, iSegmentAddress, iSegmentAddress + iMaxAddress))
'
'    Do While DownLoadMode <> eDownLoadMode.StartDownload
'        DoEvents
'
'        '---- Handle start download errors here
'
'        '---- Handle timeout here
'
'    Loop
'
'    'Download the File
'    If DownLoadMode = eDownLoadMode.StartDownload Then
'        Call SetupSerialPort.CreateFirmwareDownloadPacket(sHexData, strStartingAddress)
'    End If
'
'    'Send download complete message
'
'    Exit Function
'
'ErrHandler:
'
'    Call ViewLog.Log(LogMsgTypes.ErrorMsg, "Error reading firmware file.")
'    Call MsgBox("Error reading firmware file.", vbCritical + vbOKOnly, "Error")
'
'End Function

'Private Sub SaxComm1_OnComm()
'
'    Dim sMsg As String
'    Dim intIndex As Integer
'
'    If InCommRoutine = True Then
'        Exit Sub
'    End If
'
'    InCommRoutine = True
'
'    'CommPort.RThreshold = 0
'
'    Select Case SaxComm1.CommEvent
''        Case comEventBreak                          'A Break was received.
''        Case comEventFrame                          'Framing Error
''            PacketErrors += 1
''            Beep()
''        Case comEventOverrun                        'Data Lost
''            PacketErrors += 1
''            Beep()
''        Case comEventRxOver                         'Receive buffer overflow
''            PacketErrors += 1
''            Beep()
''        Case comEventRxParity                       'Parity Error
''            PacketErrors += 1
''            Beep()
''        Case comEventTxFull                         'Transmit buffer full
''            PacketErrors += 1
''            Beep()
''        Case comEventDCB                            'Unexpected error retrieving DCB
'        Case comEvCD                                'Change in the CD line
'            ViewLog.Log DebugMsg, "CD Line Event"
''        Case comEvCTS                               'Change in the CTS line
''        Case comEvDSR                               'Change in the DSR line
''        Case comEvRing                              'Change in the Ring Indicator
'        Case comEvReceive                           'Received RThreshold # of chars
'            If SaxComm1.InBufferCount > 0 Then        'If there are characters in the serial buffer
'
'                Dim sBuf As String
'                sBuf = ""
'                sBuf = SaxComm1.Input
'                CommBuffer = CommBuffer + sBuf
'                If Len(CommBuffer) >= 2 And PacketHeaderReceived = False Then
''                    PacketHeaderReceived = True
''                    CommBufferLen = Asc(Mid(CommBuffer, 1, 1)) + Asc(Mid(CommBuffer, 2, 1))
'                    ' MCS scan for valid lengths
'                    For intIndex = 1 To Len(CommBuffer) - 2
'                        CommBufferLen = Asc(Mid(CommBuffer, intIndex, 1)) + _
'                                                Asc(Mid(CommBuffer, intIndex + 1, 1))
'                        If CommBufferLen = 11 Or CommBufferLen = 34 Or CommBufferLen = 4 Or _
'                                CommBufferLen = 14 Or CommBufferLen = 6 Or CommBufferLen = 8 Then
'                            CommBuffer = Mid$(CommBuffer, intIndex, CommBufferLen)
'                            PacketHeaderReceived = True
'                            Exit For
'                        Else
'                            CommBufferLen = 0
'                        End If
'                    Next intIndex
'                End If
'                If Len(CommBuffer) >= CommBufferLen And CommBufferLen > 0 And PacketHeaderReceived = True Then
'                    Dim sCurMsg As String
'                    ' MCS changed from 1 to 2
'                    sCurMsg = Mid(CommBuffer, 1, CommBufferLen)
'                    Dim iLength As Integer
'                    ' MCS removed -1
'                    iLength = Len(sCurMsg)
'                    If Len(CommBuffer) > CommBufferLen Then
'                        CommBuffer = Mid(CommBuffer, CommBufferLen, Len(CommBuffer))
'                    Else
'                        CommBuffer = ""
'                    End If
'                    PacketHeaderReceived = False
'
'                    '---- Convert string message to hex bytes
'                    'Dim ByteArray(iLength) As Byte
'
'                    Dim ByteArray(256) As Byte
'                    Dim iByte As Integer
'
'                    LastCommMessage = LastCommMessage & "RX:"
'                    For iByte = 0 To iLength - 1
'                        ' MCS Added +1 to ibyte
'                        ByteArray(iByte) = Asc(Mid(sCurMsg, iByte + 1, 1))
'
'                        LastCommMessage = LastCommMessage & ByteArray(iByte) & " "        'RDR NEEDS HEX CONVERSION   x2
'                    Next
'
'                    LastCommMessage = LastCommMessage & "(" & IncomingCheckSum(ByteArray) & ")"   'RDR NEEDS HEX CONVERSION x2
'
'                    '---- Parse the protocol
'                    ' MCS Changed 2 to 3
'                    Select Case (Asc(Mid(sCurMsg, 3, 1)) - (Asc(Mid(sCurMsg, 3, 1)) And &HF))
'                        Case eCommand.Ack                   'Acknowledgement Packet
'                            LastCommMessage = LastCommMessage & "ACK"
'                            PacketsReceived = PacketsReceived + 1
'                        Case eCommand.Cmd                   'Command Packet
'                            Select Case ByteArray(4)            'bDataCmd
'                                ' Diagnostic Message
'                                Case eDataCommand.eDiagnosticMessage
'                                    LastCommMessage = LastCommMessage & "DIAG"
'                                    Dim iTarget As Integer
'                                    iTarget = 1
'                                    Dim iMask As Integer                          'Each Failure Code is 32 bits
'                                    Dim sDiagMsg As String
'                                    sDiagMsg = ""
'                                    MainForm.DiagnosticMessagesText.Text = ""
'
'                                    '---- Loop through diagnostic code bytes
'                                    ' MCS change to 34 vs ubound(bytearray)
'                                    For iByte = 5 To 34 - 2 Step 4   'Data starts at 5th byte
'                                        '---- Compute bitmask
'                                        iMask = ByteArray(iByte) + (256 * ByteArray(iByte + 1))
'                                        '---- Loop through the diagnostic codes and check for match
'                                        sDiagMsg = sDiagMsg + SetupDiagnosticCodes.GetCode(iMask)
'                                        ' Don't Display the Left and Right Tag
'                                        If iTarget < 6 Then
'                                            If sDiagMsg <> "" Then
'                                                MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
'                                                                                        sDiagnosticECU(iTarget) & " - " & sDiagMsg & " Failure " & vbCrLf
'                                            Else
'                                                MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
'                                                                                        sDiagnosticECU(iTarget) & " - Passed " & vbCrLf
'                                            End If
'                                        End If
'                                        sDiagMsg = ""
'                                        iTarget = iTarget + 1
'                                    Next
'                                Case eDataCommand.eCCMStartup
'                                    LastCommMessage = LastCommMessage & "CCM STARTUP"
'                                Case eDataCommand.eCommandError
'                                    LastCommMessage = LastCommMessage & "ERROR"
'                                    PacketErrors = PacketErrors + 1
'                                Case eDataCommand.eGetParameter
'                                    LastCommMessage = LastCommMessage & "GET"
'                                Case eDataCommand.eSetParamater
'                                    LastCommMessage = LastCommMessage & "SET"
'                                    Dim iRow As Integer
'                                    Dim lValue As Currency
'                                    Dim blnSkip As Boolean
'
'                                    blnSkip = False
'                                    ' MCS
'                                    lValue = ByteArray(6) + (256 * CLng(ByteArray(7))) + (65536 * CLng(ByteArray(8))) + (16777216@ * CCur(ByteArray(9)))
'                                    If lValue > 256@ * 256@ * 256@ Then
'                                        lValue = lValue - (256@ * 256@ * 256@ * 256@)
'                                    End If
'
'                                    With MainForm.ParameterSpread
'                                        .Col = 5                                'Parameter codes are in column 5
'                                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
'                                            .Row = iRow                         'Set active row
'                                            If Val(.Text) = ByteArray(5) Then      'If the cell value is same as command byte
'                                                '---- Get the data format
'                                                .Col = 4
'                                                Dim iDecimalPlaces As Integer
'
'                                                iDecimalPlaces = Val(Mid(.Text, InStr(1, .Text, ".") + 1, 1))
'
'                                                Select Case ByteArray(3)           'Byte 3 has the target id
'                                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
'                                                        .Col = 8
'                                                    Case eNodeMask.LEFT_REAR_NODE_MASK
'                                                        .Col = 10
'                                                    Case eNodeMask.LEFT_TAG_NODE_MASK
'                                                        .Col = 12
'                                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
'                                                        .Col = 9
'                                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
'                                                        .Col = 11
'                                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
'                                                        .Col = 13
'                                                    Case eNodeMask.CCM_NODE_MASK
'                                                        .Col = 14
'                                                End Select
'                                                .Text = IIf(iDecimalPlaces > 0, (lValue / (10 ^ iDecimalPlaces)), lValue)
'                                                Exit For
'                                            End If
'                                        Next
'                                    End With
'
'                                    '---- Check firmware table
'                                    With MainForm.ConfigurationSpread
'                                        .Col = 6                                'Parameter codes are in column 5
'                                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
'                                            .Row = iRow                         'Set active row
'                                            If Val(.Text) = ByteArray(5) Then   'If the cell value is same as command byte
'                                                Select Case ByteArray(3)           'Byte 3 has the target id
'                                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
'                                                        .Col = 2
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.LEFT_REAR_NODE_MASK
'                                                        .Col = 4
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
'                                                        .Col = 3
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
'                                                        .Col = 5
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.LEFT_TAG_NODE_MASK
'                                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
'                                                    Case eNodeMask.CCM_NODE_MASK
'                                                    Case Else
'                                                End Select
'                                            End If
'                                        Next
'                                    End With
'
'                                Case eDataCommand.eGetAllParameters
'                                    LastCommMessage = LastCommMessage & "GETALL"
'                                Case eDataCommand.eBroadcastMessage
'
'                                    '---- Graphing parameters are received in Broadcast messages
'                                    LastCommMessage = LastCommMessage & "BCAST"
'
'                                    Dim slot0 As Long, slot1 As Long, slot2 As Long, slot3 As Long
'
'                                    slot0 = "&h" & Trim(Hex(ByteArray(6))) & Trim(Hex(ByteArray(5)))
'                                    slot1 = "&h" & Trim(Hex(ByteArray(8))) & Trim(Hex(ByteArray(7)))
'                                    slot2 = "&h" & Trim(Hex(ByteArray(10))) & Trim(Hex(ByteArray(9)))
'                                    slot3 = "&h" & Trim(Hex(ByteArray(12))) & Trim(Hex(ByteArray(11)))
'
'                                    Call UpdateGraphs(slot0, slot1, slot2, slot3)
'
'
'                                '---- FIRMWARE MESSAGES:
'                                Case eDataCommand.eDownloadData
'                                    LastCommMessage = LastCommMessage & "DOWNLOADDATA"
'
'                                Case eDataCommand.eDownloadStatus
'                                    LastCommMessage = LastCommMessage & "DOWNLOADSTATUS"
'
'                                Case eDataCommand.eDownloadEraseStatus
'                                    LastCommMessage = LastCommMessage & "DOWNLOADERASESTATUS"
'                                    Select Case ByteArray(5)
'                                        Case 0  ' OK
'                                            FirmwareDownload.DownLoadMode = eDownLoadMode.StartDownload
'                                        Case 1  ' Config Error
'                                            FirmwareDownload.DownLoadStatus = CONFIG_ERR
'                                        Case 3  ' Erase Command Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_CMD_ERR
'                                        Case 4  ' Erase Start Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_START_ERR
'                                        Case 5  ' Erase Timeout Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_TIMEOUT_ERR
'                                        Case 9  ' Memory Error
'                                            FirmwareDownload.DownLoadStatus = MEMORY_ERR
'                                    End Select
'                                Case eDataCommand.eDownloadComplete
'                                    LastCommMessage = LastCommMessage & "DOWNLOADCOMPLETE"
'                                Case eDataCommand.eDownloadVerified
'                                    LastCommMessage = LastCommMessage & "DOWNLOADVERIFIED"
'                                Case eDataCommand.eKonect
'                                    LastCommMessage = LastCommMessage & "E-CONNECT"
'                                Case Else
'                                    LastCommMessage = LastCommMessage & "UNKNOWN"
'                                    PacketErrors = PacketErrors + 1
'                            End Select
'                            PacketsReceived = PacketsReceived + 1
'
'                            '---- Send ACK
'                            Call ComSend(CreateAckPacket(eNodeMask.CCM_NODE_MASK, ByteArray(2)))
'
'                        Case eCommand.Ret                   'Retry Packet
'                            LastCommMessage = LastCommMessage & "Retry"
'                            PacketsReceived = PacketsReceived + 1
'                        Case Else
'                            LastCommMessage = LastCommMessage & "Error"
'                            PacketErrors = PacketErrors + 1
'                    End Select
'                    If InStr(1, UCase(LastCommMessage), "ERROR") <> 0 Or InStr(1, LastCommMessage, "UNKNOWN") Then
'                        Call ViewLog.Log(ErrorMsg, LastCommMessage)
'                    Else
'                        Call ViewLog.Log(DebugMsg, LastCommMessage)
'                    End If
'                    Call MainForm.UpdateStatus
'                    LastCommMessage = ""
'
'                Else
'                    ' Added MCS 1/23/2003 see what's being dumped
''                    Dim sMsg As String
''                    Dim intIndex As Integer
'
'                    sMsg = ""
'                    For intIndex = 1 To Len(sBuf)
'                        sMsg = sMsg & (Asc((Mid(sBuf, intIndex, 1)))) & " "
'                    Next
'                    ViewLog.Log ErrorMsg, "Bad Data " & sMsg
'
'                End If
'                WaitingForData = False
'                CommStatus = "Port Open."
'                CommTimeOut = 0
'            Else
'            End If
'
'            Case comEvSend                              'SThreshold number of characters in the transmit buffer
'            Case comEvEOF                               'An EOF charater was found in the input stream
'            Case Else
'                ViewLog.Log DebugMsg, "Comm event (" & Trim(Str(SaxComm1.CommEvent)) & ") triggered."
'        End Select
'
'    'CommPort.RThreshold = 1
'    InCommRoutine = False
'
'End Sub

'Private Sub SaxComm1_Receive()
'
'    Dim sMsg As String
'    Dim intIndex As Integer
'
'    If InCommRoutine = True Then
'        Exit Sub
'    End If
'
'    InCommRoutine = True
'
'    'CommPort.RThreshold = 0
'
'    Select Case SaxComm1.CommEvent
'        Case comEvCD                                'Change in the CD line
'            ViewLog.Log DebugMsg, "CD Line Event"
'        Case comEvReceive                           'Received RThreshold # of chars
'            If SaxComm1.InBufferCount > 0 Then        'If there are characters in the serial buffer
'
'                Dim sBuf As String
'                sBuf = ""
'                sBuf = SaxComm1.Input
'                CommBuffer = CommBuffer + sBuf
'                If Len(CommBuffer) >= 2 And PacketHeaderReceived = False Then
''                    PacketHeaderReceived = True
''                    CommBufferLen = Asc(Mid(CommBuffer, 1, 1)) + Asc(Mid(CommBuffer, 2, 1))
'                    ' MCS scan for valid lengths
'                    For intIndex = 1 To Len(CommBuffer) - 2
'                        CommBufferLen = Asc(Mid(CommBuffer, intIndex, 1)) + _
'                                                Asc(Mid(CommBuffer, intIndex + 1, 1))
'                        If CommBufferLen = 11 Or CommBufferLen = 34 Or CommBufferLen = 4 Or _
'                                CommBufferLen = 14 Or CommBufferLen = 6 Or CommBufferLen = 8 Then
'                            CommBuffer = Mid$(CommBuffer, intIndex, CommBufferLen)
'                            PacketHeaderReceived = True
'                            Exit For
'                        Else
'                            CommBufferLen = 0
'                        End If
'                    Next intIndex
'                End If
'                If Len(CommBuffer) >= CommBufferLen And CommBufferLen > 0 And PacketHeaderReceived = True Then
'                    Dim sCurMsg As String
'                    ' MCS changed from 1 to 2
'                    sCurMsg = Mid(CommBuffer, 1, CommBufferLen)
'                    Dim iLength As Integer
'                    ' MCS removed -1
'                    iLength = Len(sCurMsg)
'                    If Len(CommBuffer) > CommBufferLen Then
'                        CommBuffer = Mid(CommBuffer, CommBufferLen, Len(CommBuffer))
'                    Else
'                        CommBuffer = ""
'                    End If
'                    PacketHeaderReceived = False
'
'                    '---- Convert string message to hex bytes
'                    'Dim ByteArray(iLength) As Byte
'
'                    Dim ByteArray(256) As Byte
'                    Dim iByte As Integer
'
'                    LastCommMessage = LastCommMessage & "RX:"
'                    For iByte = 0 To iLength - 1
'                        ' MCS Added +1 to ibyte
'                        ByteArray(iByte) = Asc(Mid(sCurMsg, iByte + 1, 1))
'
'                        LastCommMessage = LastCommMessage & ByteArray(iByte) & " "        'RDR NEEDS HEX CONVERSION   x2
'                    Next
'
'                    LastCommMessage = LastCommMessage & "(" & IncomingCheckSum(ByteArray) & ")"   'RDR NEEDS HEX CONVERSION x2
'
'                    '---- Parse the protocol
'                    ' MCS Changed 2 to 3
'                    Select Case (Asc(Mid(sCurMsg, 3, 1)) - (Asc(Mid(sCurMsg, 3, 1)) And &HF))
'                        Case eCommand.Ack                   'Acknowledgement Packet
'                            LastCommMessage = LastCommMessage & "ACK"
'                            PacketsReceived = PacketsReceived + 1
'                        Case eCommand.Cmd                   'Command Packet
'                            Select Case ByteArray(4)            'bDataCmd
'                                ' Diagnostic Message
'                                Case eDataCommand.eDiagnosticMessage
'                                    LastCommMessage = LastCommMessage & "DIAG"
'                                    Dim iTarget As Integer
'                                    iTarget = 1
'                                    Dim iMask As Integer                          'Each Failure Code is 32 bits
'                                    Dim sDiagMsg As String
'                                    sDiagMsg = ""
'                                    MainForm.DiagnosticMessagesText.Text = ""
'
'                                    '---- Loop through diagnostic code bytes
'                                    ' MCS change to 34 vs ubound(bytearray)
'                                    For iByte = 5 To 34 - 2 Step 4   'Data starts at 5th byte
'                                        '---- Compute bitmask
'                                        iMask = ByteArray(iByte) + (256 * ByteArray(iByte + 1))
'                                        '---- Loop through the diagnostic codes and check for match
'                                        sDiagMsg = sDiagMsg + SetupDiagnosticCodes.GetCode(iMask)
'                                        ' Don't Display the Left and Right Tag
'                                        If iTarget < 6 Then
'                                            If sDiagMsg <> "" Then
'                                                MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
'                                                                                        sDiagnosticECU(iTarget) & " - " & sDiagMsg & " Failure " & vbCrLf
'                                            Else
'                                                MainForm.DiagnosticMessagesText.Text = MainForm.DiagnosticMessagesText.Text & _
'                                                                                        sDiagnosticECU(iTarget) & " - Passed " & vbCrLf
'                                            End If
'                                        End If
'                                        sDiagMsg = ""
'                                        iTarget = iTarget + 1
'                                    Next
'                                Case eDataCommand.eCCMStartup
'                                    LastCommMessage = LastCommMessage & "CCM STARTUP"
'                                Case eDataCommand.eCommandError
'                                    LastCommMessage = LastCommMessage & "ERROR"
'                                    PacketErrors = PacketErrors + 1
'                                Case eDataCommand.eGetParameter
'                                    LastCommMessage = LastCommMessage & "GET"
'                                Case eDataCommand.eSetParamater
'                                    LastCommMessage = LastCommMessage & "SET"
'                                    Dim iRow As Integer
'                                    Dim lValue As Currency
'                                    Dim blnSkip As Boolean
'
'                                    blnSkip = False
'                                    ' MCS
'                                    lValue = ByteArray(6) + (256 * CLng(ByteArray(7))) + (65536 * CLng(ByteArray(8))) + (16777216@ * CCur(ByteArray(9)))
'                                    If lValue > 256@ * 256@ * 256@ Then
'                                        lValue = lValue - (256@ * 256@ * 256@ * 256@)
'                                    End If
'
'                                    With MainForm.ParameterSpread
'                                        .Col = 5                                'Parameter codes are in column 5
'                                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
'                                            .Row = iRow                         'Set active row
'                                            If Val(.Text) = ByteArray(5) Then      'If the cell value is same as command byte
'                                                '---- Get the data format
'                                                .Col = 4
'                                                Dim iDecimalPlaces As Integer
'
'                                                iDecimalPlaces = Val(Mid(.Text, InStr(1, .Text, ".") + 1, 1))
'
'                                                Select Case ByteArray(3)           'Byte 3 has the target id
'                                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
'                                                        .Col = 8
'                                                    Case eNodeMask.LEFT_REAR_NODE_MASK
'                                                        .Col = 10
'                                                    Case eNodeMask.LEFT_TAG_NODE_MASK
'                                                        .Col = 12
'                                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
'                                                        .Col = 9
'                                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
'                                                        .Col = 11
'                                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
'                                                        .Col = 13
'                                                    Case eNodeMask.CCM_NODE_MASK
'                                                        .Col = 14
'                                                End Select
'                                                .Text = IIf(iDecimalPlaces > 0, (lValue / (10 ^ iDecimalPlaces)), lValue)
'                                                Exit For
'                                            End If
'                                        Next
'                                    End With
'
'                                    '---- Check firmware table
'                                    With MainForm.ConfigurationSpread
'                                        .Col = 6                                'Parameter codes are in column 5
'                                        For iRow = 0 To .MaxRows                'Loop for each row in the spreadsheet
'                                            .Row = iRow                         'Set active row
'                                            If Val(.Text) = ByteArray(5) Then   'If the cell value is same as command byte
'                                                Select Case ByteArray(3)           'Byte 3 has the target id
'                                                    Case eNodeMask.LEFT_FRONT_NODE_MASK
'                                                        .Col = 2
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.LEFT_REAR_NODE_MASK
'                                                        .Col = 4
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.RIGHT_FRONT_NODE_MASK
'                                                        .Col = 3
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.RIGHT_REAR_NODE_MASK
'                                                        .Col = 5
'                                                        .Text = lValue
'                                                        Exit For
'                                                    Case eNodeMask.LEFT_TAG_NODE_MASK
'                                                    Case eNodeMask.RIGHT_TAG_NODE_MASK
'                                                    Case eNodeMask.CCM_NODE_MASK
'                                                    Case Else
'                                                End Select
'                                            End If
'                                        Next
'                                    End With
'
'                                Case eDataCommand.eGetAllParameters
'                                    LastCommMessage = LastCommMessage & "GETALL"
'                                Case eDataCommand.eBroadcastMessage
'
'                                    '---- Graphing parameters are received in Broadcast messages
'                                    LastCommMessage = LastCommMessage & "BCAST"
'
'                                    Dim slot0 As Long, slot1 As Long, slot2 As Long, slot3 As Long
'
'                                    slot0 = "&h" & Trim(Hex(ByteArray(6))) & Trim(Hex(ByteArray(5)))
'                                    slot1 = "&h" & Trim(Hex(ByteArray(8))) & Trim(Hex(ByteArray(7)))
'                                    slot2 = "&h" & Trim(Hex(ByteArray(10))) & Trim(Hex(ByteArray(9)))
'                                    slot3 = "&h" & Trim(Hex(ByteArray(12))) & Trim(Hex(ByteArray(11)))
'
'                                    Call UpdateGraphs(slot0, slot1, slot2, slot3)
'
'
'                                '---- FIRMWARE MESSAGES:
'                                Case eDataCommand.eDownloadData
'                                    LastCommMessage = LastCommMessage & "DOWNLOADDATA"
'
'                                Case eDataCommand.eDownloadStatus
'                                    LastCommMessage = LastCommMessage & "DOWNLOADSTATUS"
'
'                                Case eDataCommand.eDownloadEraseStatus
'                                    LastCommMessage = LastCommMessage & "DOWNLOADERASESTATUS"
'                                    Select Case ByteArray(5)
'                                        Case 0  ' OK
'                                            FirmwareDownload.DownLoadMode = eDownLoadMode.StartDownload
'                                        Case 1  ' Config Error
'                                            FirmwareDownload.DownLoadStatus = CONFIG_ERR
'                                        Case 3  ' Erase Command Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_CMD_ERR
'                                        Case 4  ' Erase Start Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_START_ERR
'                                        Case 5  ' Erase Timeout Error
'                                            FirmwareDownload.DownLoadStatus = ERASE_TIMEOUT_ERR
'                                        Case 9  ' Memory Error
'                                            FirmwareDownload.DownLoadStatus = MEMORY_ERR
'                                    End Select
'                                Case eDataCommand.eDownloadComplete
'                                    LastCommMessage = LastCommMessage & "DOWNLOADCOMPLETE"
'                                Case eDataCommand.eDownloadVerified
'                                    LastCommMessage = LastCommMessage & "DOWNLOADVERIFIED"
'                                Case eDataCommand.eKonect
'                                    LastCommMessage = LastCommMessage & "E-CONNECT"
'                                Case Else
'                                    LastCommMessage = LastCommMessage & "UNKNOWN"
'                                    PacketErrors = PacketErrors + 1
'                            End Select
'                            PacketsReceived = PacketsReceived + 1
'
'                            '---- Send ACK
'                            Call ComSend(CreateAckPacket(eNodeMask.CCM_NODE_MASK, ByteArray(2)))
'
'                        Case eCommand.Ret                   'Retry Packet
'                            LastCommMessage = LastCommMessage & "Retry"
'                            PacketsReceived = PacketsReceived + 1
'                        Case Else
'                            LastCommMessage = LastCommMessage & "Error"
'                            PacketErrors = PacketErrors + 1
'                    End Select
'                    If InStr(1, UCase(LastCommMessage), "ERROR") <> 0 Or InStr(1, LastCommMessage, "UNKNOWN") Then
'                        Call ViewLog.Log(ErrorMsg, LastCommMessage)
'                    Else
'                        Call ViewLog.Log(DebugMsg, LastCommMessage)
'                    End If
'                    Call MainForm.UpdateStatus
'                    LastCommMessage = ""
'
'                Else
'                    ' Added MCS 1/23/2003 see what's being dumped
''                    Dim sMsg As String
''                    Dim intIndex As Integer
'
'                    sMsg = ""
'                    For intIndex = 1 To Len(sBuf)
'                        sMsg = sMsg & (Asc((Mid(sBuf, intIndex, 1)))) & " "
'                    Next
'                    ViewLog.Log ErrorMsg, "Bad Data " & sMsg
'
'                End If
'                WaitingForData = False
'                CommStatus = "Port Open."
'                CommTimeOut = 0
'            Else
'            End If
'
'            Case comEvSend                              'SThreshold number of characters in the transmit buffer
'            Case comEvEOF                               'An EOF charater was found in the input stream
'            Case Else
'                ViewLog.Log DebugMsg, "Comm event (" & Trim(Str(SaxComm1.CommEvent)) & ") triggered."
'        End Select
'
'    'CommPort.RThreshold = 1
'    InCommRoutine = False
'
'End Sub


