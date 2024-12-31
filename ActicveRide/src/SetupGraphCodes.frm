VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form SetupGraphCodes 
   Caption         =   "Setup Graph Codes"
   ClientHeight    =   8190
   ClientLeft      =   885
   ClientTop       =   2295
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1920
      Top             =   7380
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
            Picture         =   "SetupGraphCodes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupGraphCodes.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupGraphCodes.frx":0644
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   840
      Left            =   90
      TabIndex        =   0
      Top             =   7290
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
   Begin FPSpread.vaSpread GraphCodesSpread 
      Height          =   7185
      Left            =   60
      TabIndex        =   1
      Top             =   60
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
      SpreadDesigner  =   "SetupGraphCodes.frx":0966
   End
End
Attribute VB_Name = "SetupGraphCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_cintView As Integer = 1
Private Const m_cintParameter As Integer = 2
Private Const m_cintUnits As Integer = 3
Private Const m_cintFormat As Integer = 4
Private Const m_cintID As Integer = 5
Private Const m_cintMin As Integer = 6
Private Const m_cintMax As Integer = 7
Private Const m_cintLeftFront As Integer = 8
Private Const m_cintRightFront As Integer = 9
Private Const m_cintLeftRear As Integer = 10
Private Const m_cintRightRear As Integer = 11
Private Const m_cintLeftTag As Integer = 12
Private Const m_cintRightTag As Integer = 13
Private Const m_cintCentral As Integer = 14

Private Const sFileName As String = "Graphing"
Private m_intParameterID As Integer
Private m_intParameterRow As Integer
Private m_intDecimalPlaces As Integer

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       DecimalPlaces
' Description:
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       10/08/2003-11:28:06
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get DecimalPlaces() As Integer
        
    With GraphCodesSpread
        ' Set the Parameter Row
        .Row = m_intParameterRow
        ' Set the format column
        .Col = m_cintFormat
        ' Get the Text
        m_intDecimalPlaces = Val(Mid$(.Text, InStr(1, .Text, ".") + 1, 1))
    End With
    
    DecimalPlaces = m_intDecimalPlaces

End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       ParameterID
' Description:       [type_description_here]
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       10/08/2003-11:22:05
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Get ParameterID() As Integer
    
    ParameterID = m_intParameterID

End Property

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       ParameterID
' Description:
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       10/08/2003-11:22:05
'
' Parameters :       intValue (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Public Property Let ParameterID(ByVal intValue As Integer)
    
    Dim intIndex As Integer
    
    m_intParameterID = intValue
    
    ' find the Parameter in the Grid
    With GraphCodesSpread
        .Col = m_cintParameter
        For intIndex = 1 To .MaxRows
            .Row = intIndex
            If Val(.Text) = intValue Then
                m_intParameterRow = intIndex
                Exit For
            End If
        Next
    End With

End Property

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
100     OpenGraphCodes
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupGraphCodes.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub OpenGraphCodes()
        '<EhHeader>
        On Error GoTo OpenGraphCodes_Err
        '</EhHeader>
    
    '    Dim iRow As Integer
    
100     With GraphCodesSpread
102         .LoadTextFile sFileName & sFileExt, ",", ",", vbCrLf, LoadTextFileColHeaders, sFileName & sLogExt
104         .EditMode = False
106         .ColWidth(2) = 20
108         .LockBackColor = vbCyan
110         If Splash.iCurrentLevel = SecurityLevels.eView Then
112             .Lock = True
            Else
114             .Lock = False
            End If
    '         Started play with turning on and off the graph parameter they
    '         want to view
116         .Col = 1
118         .TypeHAlign = TypeHAlignCenter
    '        For iRow = 1 To .MaxRows
    '            .Col = 1
    '            .Row = iRow
    '            .CellType = CellTypeCheckBox
    '            .TypeCheckType = TypeCheckTypeNormal
    '            .TypeCheckCenter = True
    '            .Lock = False
    ''            .Col = 2
    ''            .Lock = True
    '        Next
        End With

        '<EhFooter>
        Exit Sub

OpenGraphCodes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupGraphCodes.OpenGraphCodes " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


' Added MCS
Private Sub Form_Resize()

    On Error Resume Next
    
    With Toolbar
        .Top = Me.ScaleHeight - .Left - .Height
    End With

    With GraphCodesSpread
        .Width = Me.ScaleWidth - (.Left * 2)
        .Height = Toolbar.Top - (.Top * 2)
    End With

    On Error GoTo 0
    
End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo ToolBar_ButtonClick_Err
        '</EhHeader>
100     With GraphCodesSpread
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
               "in ActiveRide.SetupGraphCodes.ToolBar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function GetCode(ByVal iMask As Integer) As String

    Dim sReturn As String
    Dim iRow As Integer
    
    sReturn = ""
    iRow = 0
    With GraphCodesSpread
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 1
            If iMask And 2 ^ Val(.Text) Then
                .Col = 2
                sReturn = sReturn + .Text & " "
            End If
        Next
    End With
    If Len(sReturn) = 0 And iMask <> 0 Then
        sReturn = iMask
    End If
    GetCode = sReturn

End Function


