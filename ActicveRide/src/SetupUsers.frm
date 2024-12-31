VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form SetupUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Users"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1920
      Top             =   7440
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
            Picture         =   "SetupUsers.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupUsers.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SetupUsers.frx":0644
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   840
      Left            =   90
      TabIndex        =   3
      Top             =   7350
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
   Begin FPSpread.vaSpread UserSpread 
      Height          =   7185
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      _Version        =   393216
      _ExtentX        =   10610
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
      SpreadDesigner  =   "SetupUsers.frx":0966
   End
   Begin VB.Frame SecurityLevelFrame 
      Caption         =   "Security Levels"
      Height          =   4245
      Left            =   6090
      TabIndex        =   0
      Top             =   60
      Width           =   5115
      Begin FPSpread.vaSpread SecurityLevelsSpread 
         Height          =   3885
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4905
         _Version        =   393216
         _ExtentX        =   8652
         _ExtentY        =   6853
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
         SpreadDesigner  =   "SetupUsers.frx":0B3A
      End
   End
End
Attribute VB_Name = "SetupUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     With Me.SecurityLevelsSpread
102         .LoadTextFile "SecurityLevels.csv", ",", ",", vbCrLf, LoadTextFileColHeaders + LoadTextFileRowHeaders, "Securitylevels.log"
104         .LockBackColor = vbCyan
106         .Lock = True
108         .ColWidth(1) = 25
        End With

110     With Me.UserSpread
112         .LoadTextFile "Users.csv", ",", ",", vbCrLf, LoadTextFileColHeaders, "Users.log"
114         .ColWidth(1) = 25
116         .ColWidth(2) = 10
        End With

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupUsers.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo ToolBar_ButtonClick_Err
        '</EhHeader>

100     With Me.UserSpread
102         Select Case Button.index
                Case 1
104                 .ExportToTextFile "Users.csv", ",", ",", vbCrLf, ExportToTextFileColHeaders, "Users.log"
106                 Me.Hide
108             Case 2
110                 .MaxRows = .MaxRows + 1
112                 .Row = .MaxRows - 1
114             Case 3
116                 .DeleteRows .ActiveRow, 1
118                 .MaxRows = .MaxRows - 1
            End Select
        End With

        '<EhFooter>
        Exit Sub

ToolBar_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupUsers.ToolBar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function GetUserSecurityLevel(ByVal sUserName As String) As Integer
        '<EhHeader>
        On Error GoTo GetUserSecurityLevel_Err
        '</EhHeader>
        
        Dim RowNum As Integer
        Dim SecurityLevel As Integer
        
100     SecurityLevel = -1
102     With UserSpread
104         .Col = 1
106         For RowNum = 1 To .MaxRows
108             .Row = RowNum
110             If Trim$(UCase$(.Text)) = Trim$(UCase$(sUserName)) Then
112                 .Col = 2
114                 SecurityLevel = Val(.Text)
                    Exit For
                End If
            Next
        End With
116     GetUserSecurityLevel = SecurityLevel
        '<EhFooter>
        Exit Function

GetUserSecurityLevel_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.SetupUsers.GetUserSecurityLevel " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

