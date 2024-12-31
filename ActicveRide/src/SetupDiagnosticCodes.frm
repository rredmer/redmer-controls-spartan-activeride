VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form SetupDiagnosticCodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Diagnostic Codes"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1920
      Top             =   7410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupDiagnosticCodes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupDiagnosticCodes.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupDiagnosticCodes.frx":0644
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   840
      Left            =   90
      TabIndex        =   0
      Top             =   7320
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread DiagnosticCodesSpread 
      Height          =   7185
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   13485
      _Version        =   393216
      _ExtentX        =   23786
      _ExtentY        =   12674
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
      SpreadDesigner  =   "SetupDiagnosticCodes.frx":0966
   End
End
Attribute VB_Name = "SetupDiagnosticCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sFileName As String = "DiagnosticCodes"

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
100     OpenDiagnosticCodes
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupDiagnosticCodes.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub OpenDiagnosticCodes()
        '<EhHeader>
        On Error GoTo OpenDiagnosticCodes_Err
        '</EhHeader>
    
100     With DiagnosticCodesSpread
102         .LoadTextFile sFileName & sFileExt, ",", ",", vbCrLf, LoadTextFileColHeaders, sFileName & sLogExt
104         .ColWidth(2) = 50
106         .LockBackColor = vbCyan
108         If Splash.iCurrentLevel = SecurityLevels.eView Then
110             .Lock = True
            Else
112             .Lock = False
            End If
        End With

        '<EhFooter>
        Exit Sub

OpenDiagnosticCodes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupDiagnosticCodes.OpenDiagnosticCodes " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo ToolBar_ButtonClick_Err
        '</EhHeader>
        
100     With DiagnosticCodesSpread
102         Select Case Button.index
                Case 1
104                 .ExportToTextFile sFileName & sFileExt, ",", ",", vbCrLf, ExportToTextFileColHeaders, sFileName & sLogExt
106                 Me.Hide
108             Case 2
110                 .MaxRows = .MaxRows + 1
112                 .SetActiveCell 1, .MaxRows
114             Case 3
116                 .Row = .ActiveRow
118                 .Col = 2
120                 If MsgBox("Delete " & Trim$(Str$(.Text)) & "?", vbApplicationModal + vbYesNo + vbQuestion + vbDefaultButton2, "Are you sure?") = vbYes Then
122                     .DeleteRows .ActiveRow, 1
124                     .MaxRows = .MaxRows - 1
126                     .SetActiveCell 1, 1
                    End If
            End Select
        End With
        '<EhFooter>
        Exit Sub

ToolBar_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupDiagnosticCodes.ToolBar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function GetCode(ByVal curMask As Currency) As String
        '<EhHeader>
        On Error GoTo GetCode_Err
        '</EhHeader>
    
        Dim sReturn As String
        Dim iRow As Integer
    
100     sReturn = ""
102     iRow = 0
    
104     With DiagnosticCodesSpread
106         For iRow = 1 To .MaxRows
108             .Row = iRow
110             .Col = 1
112             If curMask And 2 ^ Val(.Text) Then
114                 .Col = 2
116                 sReturn = sReturn + .Text & " "
                End If
            Next
        End With
    
118     If Len(sReturn) = 0 And curMask <> 0 Then
120         sReturn = curMask
        End If
    
122     GetCode = sReturn

        '<EhFooter>
        Exit Function

GetCode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupDiagnosticCodes.GetCode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

