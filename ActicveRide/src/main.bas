Attribute VB_Name = "Startup"
'****************************************************************************
'**                                                                        **
'** Project....: Spartan Motors ActiveRide                                 **
'**                                                                        **
'** Module.....: Main.bas - The application main module (startup)          **
'**                                                                        **
'** Description: This is the main application module, it loads all of the  **
'**              global forms.                                             **
'**                                                                        **
'** History....:                                                           **
'**    11/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

'Public Declare Sub Sleep Lib "kernel32" (ByVal Mills As Long)

' MCS Graph Enabled
Public g_blnGraphEnable(3) As Boolean
Public g_bytGraphCmd(3) As Byte
Public g_intGraphSource(3) As Integer
Public g_intGraphDivisor(3) As Integer
Public g_adblMaxValue(3) As Double
Public g_adblMinValue(3) As Double

Public Enum SecurityLevels
    eNone = 0
    eView = 1
    eEdit = 2
End Enum

Public Const sFileExt As String = ".csv"
Public Const sLogExt As String = ".log"

' Parameters Grid Column Locations
Public Const m_cintParameterParam As Integer = 1
Public Const m_cintUnitsParam As Integer = 2
Public Const m_cintFormatParam As Integer = 3
Public Const m_cintIDParam As Integer = 4
Public Const m_cintMinParam As Integer = 5
Public Const m_cintMaxParam As Integer = 6
Public Const m_cintLeftFrontParam As Integer = 7
Public Const m_cintRightFrontParam As Integer = 8
Public Const m_cintLeftRearParam As Integer = 9
Public Const m_cintRightRearParam As Integer = 10
Public Const m_cintLeftTagParam As Integer = 11
Public Const m_cintRightTagParam As Integer = 12
Public Const m_cintCentralParam As Integer = 13

' Store Modules being programmed
Public g_abytNodes(7) As Byte
' Download Node Status
Public g_aintNodeStatus(7) As Integer

Public DownLoadMode(7) As eDownLoadMode
Public DownLoadStatus(7) As eDownloadStatus

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Main                                                  **
'**                                                                        **
'**  Description..:  This is the main application procedure (called from   **
'**                  Windows).  It loads the application's forms.          **
'**                                                                        **
'****************************************************************************
Public Sub Main()
    
    On Error GoTo ErrorHandler                              'Standard Error Handler
        
    '---- Check for a previous instance of the application (prevent multiple copies from running)
    If App.PrevInstance = True Then                         'If the application is already running
        MsgBox "ActiveRide is already running.", vbSystemModal + vbCritical + vbOKOnly, "Error"
        End
    End If
    
    Load UsbKeyDiagnostics
    
    If UsbKeyDiagnostics.GetSecurity(&H1F5) = False Then
        MsgBox "Security key not found, application will not start.  Please contact technical support.", vbApplicationModal + vbCritical + vbOKOnly, "ERROR"
        End
    End If
    
    ChDir App.Path
    
    
    '---- Load the Application forms
    Splash.Show vbModal
    Load ViewLog                                            'Load the application log form (must be the first app form)
    Load ErrorForm                                          'Load the application error handler form
    Load SetupSerialPort
    Load SetupUsers
    Load SetupDiagnosticCodes
'    Load TimingForm                                         'Load the application timing form
'    Call Sleep(200)                                          'Slight delay for splash screen
    Load PrintPreview
    MainForm.Show vbModeless                                'Show the main form (MDI Parent Form)
    Splash.Hide                                             'Hide the splash window
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "Main", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
    End
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  EndTheProgram                                         **
'**                                                                        **
'**  Description..:  This routine terminates the program cleanly.          **
'**                                                                        **
'****************************************************************************
Public Function EndTheProgram() As Boolean
    On Error Resume Next                                    'There is no stopping this routine!
    
    'Removed Exit Program Question - Per Scott!
    'If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo, "Exit ActiveRide") = vbYes Then
        Dim intIndex As Integer
        ' Turn off the Broadcast Message
        For intIndex = 0 To 3
            If g_blnGraphEnable(intIndex) Then
                Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket(eDataCommand.eStopBroadcast, g_intGraphSource(intIndex), g_bytGraphCmd(intIndex), intIndex))
            End If
        Next
                    
        Unload UsbKeyDiagnostics
        Unload Splash
'            Unload TimingForm
        Unload PrintPreview
        Unload SetupGraphCodes
        Unload FirmwareDownload
        Unload ErrorForm
        Unload SetupUsers
        Unload SetupDiagnosticCodes
'            Unload MainForm
        Unload SetupSerialPort
        Unload ViewLog                                  'MUST be the last form unloaded!
        EndTheProgram = True
'            End
    'Else
    '    EndTheProgram = False
    'End If

End Function

