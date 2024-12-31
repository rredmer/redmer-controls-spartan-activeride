VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Begin VB.Form MainForm 
   Caption         =   "Spartan Motors ActiveRide"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   -1770
   ClientWidth     =   11880
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab MainTab 
      Height          =   9795
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   17277
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Status"
      TabPicture(0)   =   "MainForm.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ConfigurationSpread"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MessagesFrame"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Parameters"
      TabPicture(1)   =   "MainForm.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ParameterSpread"
      Tab(1).Control(1)=   "ParametersToolbar"
      Tab(1).Control(2)=   "CommonDialog"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Graphs"
      TabPicture(2)   =   "MainForm.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GraphFrame(1)"
      Tab(2).Control(1)=   "GraphFrame(3)"
      Tab(2).Control(2)=   "GraphFrame(2)"
      Tab(2).Control(3)=   "GraphFrame(0)"
      Tab(2).ControlCount=   4
      Begin VB.Frame GraphFrame 
         Caption         =   "Graph"
         Height          =   4755
         Index           =   1
         Left            =   -67350
         TabIndex        =   14
         Top             =   360
         Width           =   7545
         Begin VB.Frame fraControlFrame 
            BorderStyle     =   0  'None
            Height          =   1035
            Index           =   1
            Left            =   30
            TabIndex        =   26
            Top             =   3600
            Width           =   7395
            Begin VB.ComboBox GraphParameterCombo 
               Height          =   315
               Index           =   1
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   660
               Width           =   1845
            End
            Begin VB.ComboBox GraphSourceCombo 
               Height          =   315
               Index           =   1
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   1815
            End
            Begin VB.ComboBox GraphScaleCombo 
               Height          =   315
               Index           =   1
               ItemData        =   "MainForm.frx":0496
               Left            =   4500
               List            =   "MainForm.frx":04AC
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   240
               Width           =   915
            End
            Begin VB.CommandButton GraphStartButton 
               Caption         =   "Start"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lblSource 
               AutoSize        =   -1  'True
               Caption         =   "Source"
               Height          =   195
               Index           =   1
               Left            =   930
               TabIndex        =   34
               Top             =   300
               Width           =   510
            End
            Begin VB.Label lblParameter 
               AutoSize        =   -1  'True
               Caption         =   "Parameter"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   33
               Top             =   720
               Width           =   720
            End
            Begin VB.Label GraphValueLabel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.0"
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   1
               Left            =   4440
               TabIndex        =   32
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblScale 
               AutoSize        =   -1  'True
               Caption         =   "Scale"
               Height          =   195
               Index           =   1
               Left            =   4020
               TabIndex        =   31
               Top             =   300
               Width           =   405
            End
         End
         Begin ChartfxLibCtl.ChartFX GraphChart 
            Height          =   3675
            Index           =   1
            Left            =   30
            TabIndex        =   15
            Top             =   210
            Width           =   7455
            _cx             =   13150
            _cy             =   6482
            Build           =   20
            TypeMask        =   109576193
            MarkerShape     =   0
            Axis(2).Min     =   0
            Axis(2).Max     =   100
            nColors         =   16
            Colors          =   "MainForm.frx":04CD
            nSer            =   1
            NumSer          =   1
            BorderS         =   0
            _Data_          =   "MainForm.frx":056D
         End
      End
      Begin VB.Frame GraphFrame 
         Caption         =   "Graph"
         Height          =   4635
         Index           =   3
         Left            =   -67320
         TabIndex        =   12
         Top             =   4920
         Width           =   7545
         Begin VB.Frame fraControlFrame 
            BorderStyle     =   0  'None
            Height          =   1035
            Index           =   3
            Left            =   30
            TabIndex        =   44
            Top             =   3480
            Width           =   7395
            Begin VB.ComboBox GraphParameterCombo 
               Height          =   315
               Index           =   3
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   660
               Width           =   1845
            End
            Begin VB.ComboBox GraphSourceCombo 
               Height          =   315
               Index           =   3
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   240
               Width           =   1815
            End
            Begin VB.ComboBox GraphScaleCombo 
               Height          =   315
               Index           =   3
               ItemData        =   "MainForm.frx":05C5
               Left            =   4500
               List            =   "MainForm.frx":05DB
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   240
               Width           =   915
            End
            Begin VB.CommandButton GraphStartButton 
               Caption         =   "Start"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   45
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lblSource 
               AutoSize        =   -1  'True
               Caption         =   "Source"
               Height          =   195
               Index           =   3
               Left            =   930
               TabIndex        =   52
               Top             =   300
               Width           =   510
            End
            Begin VB.Label lblParameter 
               AutoSize        =   -1  'True
               Caption         =   "Parameter"
               Height          =   195
               Index           =   3
               Left            =   720
               TabIndex        =   51
               Top             =   720
               Width           =   720
            End
            Begin VB.Label GraphValueLabel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.0"
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   3
               Left            =   4440
               TabIndex        =   50
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblScale 
               AutoSize        =   -1  'True
               Caption         =   "Scale"
               Height          =   195
               Index           =   3
               Left            =   4020
               TabIndex        =   49
               Top             =   300
               Width           =   405
            End
         End
         Begin ChartfxLibCtl.ChartFX GraphChart 
            Height          =   3675
            Index           =   3
            Left            =   30
            TabIndex        =   13
            Top             =   210
            Width           =   7455
            _cx             =   13150
            _cy             =   6482
            Build           =   20
            TypeMask        =   109576193
            MarkerShape     =   0
            Axis(2).Min     =   0
            Axis(2).Max     =   100
            nColors         =   16
            Colors          =   "MainForm.frx":05FC
            nSer            =   1
            NumSer          =   1
            BorderS         =   0
            _Data_          =   "MainForm.frx":069C
         End
      End
      Begin VB.Frame GraphFrame 
         Caption         =   "Graph"
         Height          =   4335
         Index           =   2
         Left            =   -74940
         TabIndex        =   10
         Top             =   5100
         Width           =   7545
         Begin VB.Frame fraControlFrame 
            BorderStyle     =   0  'None
            Height          =   1035
            Index           =   2
            Left            =   30
            TabIndex        =   35
            Top             =   3240
            Width           =   7395
            Begin VB.ComboBox GraphParameterCombo 
               Height          =   315
               Index           =   2
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   660
               Width           =   1845
            End
            Begin VB.ComboBox GraphSourceCombo 
               Height          =   315
               Index           =   2
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   240
               Width           =   1815
            End
            Begin VB.ComboBox GraphScaleCombo 
               Height          =   315
               Index           =   2
               ItemData        =   "MainForm.frx":06F4
               Left            =   4500
               List            =   "MainForm.frx":070A
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   240
               Width           =   915
            End
            Begin VB.CommandButton GraphStartButton 
               Caption         =   "Start"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   36
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lblSource 
               AutoSize        =   -1  'True
               Caption         =   "Source"
               Height          =   195
               Index           =   2
               Left            =   930
               TabIndex        =   43
               Top             =   300
               Width           =   510
            End
            Begin VB.Label lblParameter 
               AutoSize        =   -1  'True
               Caption         =   "Parameter"
               Height          =   195
               Index           =   2
               Left            =   720
               TabIndex        =   42
               Top             =   720
               Width           =   720
            End
            Begin VB.Label GraphValueLabel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.0"
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   2
               Left            =   4440
               TabIndex        =   41
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblScale 
               AutoSize        =   -1  'True
               Caption         =   "Scale"
               Height          =   195
               Index           =   2
               Left            =   4020
               TabIndex        =   40
               Top             =   300
               Width           =   405
            End
         End
         Begin ChartfxLibCtl.ChartFX GraphChart 
            Height          =   3195
            Index           =   2
            Left            =   30
            TabIndex        =   11
            Top             =   210
            Width           =   7455
            _cx             =   13150
            _cy             =   5636
            Build           =   20
            TypeMask        =   109576193
            MarkerShape     =   0
            Axis(2).Min     =   0
            Axis(2).Max     =   100
            nColors         =   16
            Colors          =   "MainForm.frx":072B
            nSer            =   1
            NumSer          =   1
            BorderS         =   0
            _Data_          =   "MainForm.frx":07CB
         End
      End
      Begin VB.Frame GraphFrame 
         Caption         =   "Graph"
         Height          =   4755
         Index           =   0
         Left            =   -74940
         TabIndex        =   8
         Top             =   360
         Width           =   7545
         Begin VB.Frame fraControlFrame 
            BorderStyle     =   0  'None
            Height          =   1035
            Index           =   0
            Left            =   30
            TabIndex        =   17
            Top             =   3600
            Width           =   7395
            Begin VB.CommandButton GraphStartButton 
               Caption         =   "Start"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   270
               Width           =   615
            End
            Begin VB.ComboBox GraphScaleCombo 
               Height          =   315
               Index           =   0
               ItemData        =   "MainForm.frx":0823
               Left            =   4500
               List            =   "MainForm.frx":0839
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   240
               Width           =   915
            End
            Begin VB.ComboBox GraphSourceCombo 
               Height          =   315
               Index           =   0
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   240
               Width           =   1815
            End
            Begin VB.ComboBox GraphParameterCombo 
               Height          =   315
               Index           =   0
               Left            =   1500
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   660
               Width           =   1845
            End
            Begin VB.Label lblScale 
               AutoSize        =   -1  'True
               Caption         =   "Scale"
               Height          =   195
               Index           =   0
               Left            =   4020
               TabIndex        =   25
               Top             =   300
               Width           =   405
            End
            Begin VB.Label GraphValueLabel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.0"
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   0
               Left            =   4440
               TabIndex        =   24
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblParameter 
               AutoSize        =   -1  'True
               Caption         =   "Parameter"
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   23
               Top             =   720
               Width           =   720
            End
            Begin VB.Label lblSource 
               AutoSize        =   -1  'True
               Caption         =   "Source"
               Height          =   195
               Index           =   0
               Left            =   930
               TabIndex        =   22
               Top             =   300
               Width           =   510
            End
         End
         Begin ChartfxLibCtl.ChartFX GraphChart 
            Height          =   3735
            Index           =   0
            Left            =   30
            TabIndex        =   9
            Top             =   180
            Width           =   7455
            _cx             =   13150
            _cy             =   6588
            Build           =   20
            TypeMask        =   109576193
            MarkerShape     =   0
            Axis(2).Min     =   0
            Axis(2).Max     =   100
            nColors         =   16
            Colors          =   "MainForm.frx":085A
            nSer            =   1
            NumSer          =   1
            BorderS         =   0
            _Data_          =   "MainForm.frx":08FA
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   -73740
         Top             =   9180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame MessagesFrame 
         Caption         =   "Diagnostic Messages"
         Height          =   9765
         Left            =   7680
         TabIndex        =   4
         Top             =   360
         Width           =   7515
         Begin MSComctlLib.Toolbar DiagnosticCodesToolbar 
            Height          =   705
            Left            =   60
            TabIndex        =   6
            Top             =   8700
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   1244
            ButtonWidth     =   1323
            ButtonHeight    =   1138
            Wrappable       =   0   'False
            ImageList       =   "ImageList"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Check"
                  Description     =   "Check for Errors"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Clear"
                  Description     =   "Clear the Errors"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Get All"
                  Description     =   "Get All the Parameters"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Put All"
                  Description     =   "Send all the Parameters"
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Calibrate"
                  Description     =   "Calibrate Unit"
                  ImageIndex      =   10
               EndProperty
            EndProperty
         End
         Begin VB.TextBox DiagnosticMessagesText 
            Height          =   8415
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   300
            Width           =   7335
         End
      End
      Begin FPSpread.vaSpread ConfigurationSpread 
         Height          =   9285
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   7545
         _Version        =   393216
         _ExtentX        =   13309
         _ExtentY        =   16378
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
         SpreadDesigner  =   "MainForm.frx":0A07
      End
      Begin MSComctlLib.Toolbar ParametersToolbar 
         Height          =   600
         Left            =   -74940
         TabIndex        =   7
         Top             =   9060
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1058
         ButtonWidth     =   1005
         ButtonHeight    =   953
         Wrappable       =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Get All"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Put All"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread ParameterSpread 
         Height          =   8655
         Left            =   -74940
         TabIndex        =   16
         Top             =   360
         Width           =   15045
         _Version        =   393216
         _ExtentX        =   26538
         _ExtentY        =   15266
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
         SpreadDesigner  =   "MainForm.frx":0BDB
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   14760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1429
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1AA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":211D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2797
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2AB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2DCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":30ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3767
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3DE1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   926
      ButtonWidth     =   820
      ButtonHeight    =   767
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu FileMenu 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu FileMenu 
         Caption         =   "Save &As"
         Index           =   3
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Print"
         Index           =   5
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu FileMenu 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&Tools"
      Index           =   1
      Begin VB.Menu ToolsMenu 
         Caption         =   "Display &Application Log File"
         Index           =   0
      End
      Begin VB.Menu ToolsMenu 
         Caption         =   "Setup &Diagnostic Codes"
         Index           =   1
      End
      Begin VB.Menu ToolsMenu 
         Caption         =   "Setup &Graph Codes"
         Index           =   2
      End
      Begin VB.Menu ToolsMenu 
         Caption         =   "Setup &Serial Port"
         Index           =   3
      End
      Begin VB.Menu ToolsMenu 
         Caption         =   "Setup &Users"
         Index           =   4
      End
      Begin VB.Menu ToolsMenu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu ToolsMenu 
         Caption         =   "Update &Firmware"
         Index           =   6
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&Help"
      Index           =   2
      Begin VB.Menu HelpMenu 
         Caption         =   "About ActiveRide"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: RC ActiveRide                                             **
'**                                                                        **
'** Module.....: frmMain - Main application form.                          **
'**                                                                        **
'** Description: This form provides the main application interface.        **
'**                                                                        **
'** History....:                                                           **
'**    08/10/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit                                             'Require explicit variable declaration

Private intPreviousIndex As Integer


'Private bFileLoaded As Boolean                              'Set TRUE when file is loaded
'Public bGraph1Enable As Boolean                             'Graph 1 Mode (TRUE=ON)
'Public bGraph2Enable As Boolean                             'Graph 2 Mode (TRUE=ON)
'Public bGraph3Enable As Boolean                             'Graph 1 Mode (TRUE=ON)
'Public bGraph4Enable As Boolean                             'Graph 2 Mode (TRUE=ON)

'Public iGraph1Divisor As Integer
'Public iGraph2Divisor As Integer
'Public iGraph3Divisor As Integer
'Public iGraph4Divisor As Integer
'Private iGraph1RowNum As Integer
''Private iGraph2RowNum As Integer
'Private iGraph3RowNum As Integer
'Private iGraph4RowNum As Integer
'Private iGraph1Cmd As Byte                                  'Command of interest in Graph 1
'Private iGraph2Cmd As Byte                                  'Command of interest in Graph 2
'Private iGraph3Cmd As Byte                                  'Command of interest in Graph 1
'Private iGraph4Cmd As Byte                                  'Command of interest in Graph 2
'Private iGraph1Source As Integer                              'Source of graph command
'Private iGraph2Source As Integer                              'Source of graph command
'Private iGraph3Source As Integer                              'Source of graph command
'Private iGraph4Source As Integer                              'Source of graph command

'Private g_adblMaxValue(3) As Double
'Private g_adblMinValue(3) As Double
Private m_astrUnits(3) As String

Private blnLoading As Boolean

Private Sub ConfigurationSpread_Change(ByVal Col As Long, ByVal Row As Long)
        '<EhHeader>
        On Error GoTo ConfigurationSpread_Change_Err
        '</EhHeader>

100     If Col >= 2 And Col <= 5 Then
102         Call Configuration(Row, Col)
        End If

        '<EhFooter>
        Exit Sub

ConfigurationSpread_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.ConfigurationSpread_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DiagnosticCodesToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo DiagnosticCodesToolbar_ButtonClick_Err
        '</EhHeader>

100     Select Case Button.index
            ' Check for Errors
            Case 1
102             Call SetupSerialPort.GetDiagnosticMessages
            ' Clear Errors
104         Case 2
106             Call SetupSerialPort.ClearDiagnosticMessages
            ' Get Parameters
108         Case 4
110             SetupSerialPort.GetAllParameters
            ' Put Parameters
112         Case 5
114             Call PutAllParameters
            ' Calibrate
116         Case 7
118             Call Calibrate
        End Select

        '<EhFooter>
        Exit Sub

DiagnosticCodesToolbar_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.DiagnosticCodesToolbar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     blnLoading = True
    
102     OpenTextFile "Parameters.csv"
        ClearParameters                 'RDR Added to Clear Parameters
        
104     OpenConfiguration
    
        ' Set the first tab to status
106     MainTab.Tab = 0
    
        ' Security Levels
108     If Splash.iCurrentLevel = 1 Then
    '        MainTab.TabEnabled(1) = False
            ' MCS 2/24/2003
            ' Disable User Setup
110         ToolsMenu(4).Enabled = False
            ' Disable the Firmware Upload
112         ToolsMenu(6).Enabled = False
            ' Disable Graph Codes
114         ToolsMenu(2).Enabled = False
            ' Disable
116         ToolsMenu(1).Enabled = False
        End If
    
        '---- Configure the Graph Controls
        Dim GraphNum As Integer
    
118     For GraphNum = 0 To 3
120         g_blnGraphEnable(GraphNum) = False
122         GraphSourceCombo(GraphNum).Clear
124         GraphParameterCombo(GraphNum).Clear
126         With GraphChart(GraphNum)
128             .Gallery = LINES
130             .ClearData CD_DATA
132             .TypeMask = CT_EVENSPACING
134             .LineWidth = 2
136             .MaxValues = 100
138             .RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
140             .MenuBar = False
142             .Toolbar = False
144             .ShowTips = False
146             .Axis(AXIS_Y).Grid = True
                Dim iValue As Long
148             iValue = 0
150             .OpenDataEx COD_VALUES Or COD_ADDPOINTS, 1, 100
152             For iValue = 1 To .MaxValues
154                 .Value(0) = 0
                Next
156             .CloseData COD_VALUES Or COD_REALTIME
            End With
        Next
158     With SetupGraphCodes.GraphCodesSpread
            Dim ColNum As Integer
160         .Row = 0
162         For ColNum = 8 To .MaxCols
164             .Col = ColNum
166             For GraphNum = 0 To 3
                    ' Add Source to the ComboBox
168                 GraphSourceCombo(GraphNum).AddItem .Text
                    ' Store the Column Number in the Item Data
170                 GraphSourceCombo(GraphNum).ItemData(GraphSourceCombo(GraphNum).NewIndex) = ColNum
172             Next GraphNum
            Next
        End With
        
        intPreviousIndex = 1
174     For GraphNum = 0 To 3
176         GraphScaleCombo(GraphNum).ListIndex = 1
178         GraphSourceCombo(GraphNum).ListIndex = 0
    '        GraphSourceCombo_LostFocus GraphNum
180         GraphParameterCombo(GraphNum).ListIndex = 1

182         GraphScaleCombo(GraphNum).ListIndex = 1
        Next
        
184     blnLoading = False

        ' Get a parameters
186     Call SetupSerialPort.GetAllParameters
        ' Check for Diagnostic Messages
188     Call SetupSerialPort.GetDiagnosticMessages

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub FileMenu_Click(index As Integer)
        '<EhHeader>
        On Error GoTo FileMenu_Click_Err
        '</EhHeader>
    
100     Select Case index
            Case 0
                'New
102         Case 1
104             FileOpen
106         Case 2
108             FileSave
110         Case 3
                'Save As
112         Case 4
                'Line
114         Case 5
116             PrintPreview.PrepareToPreview ParameterSpread
118         Case 6
                'Line
120         Case 7
                'Exit
122             Unload Me
        End Select
    
        '<EhFooter>
        Exit Sub

FileMenu_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.FileMenu_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>
    
        ' Check to make sure the user wants to quit
100     If Not EndTheProgram Then
            ' Set Cancel Flag to stop the Unload
102         Cancel = 2
        End If

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.Form_QueryUnload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' Added 6/18/2003
Private Sub Form_Resize()

    Dim lngTabWidth As Long
    Dim lngTabHeight As Long
    Dim lngSpreadWidth As Long
    
    On Error Resume Next
    
    With MainTab
        .Width = Me.ScaleWidth - 2 * .Left
        lngTabWidth = .Width
        .Height = Me.ScaleHeight - 2 * .Top
        lngTabHeight = .Height
    End With

    With ConfigurationSpread
        .Width = lngTabWidth / 2 - 2 * 60
        lngSpreadWidth = .Width
        .Height = lngTabHeight - 500
    End With
    
    With MessagesFrame
        .Height = lngTabHeight - 500
        .Left = lngSpreadWidth + 100
        .Width = ConfigurationSpread.Width
    End With
    
    With DiagnosticCodesToolbar
        .Top = MessagesFrame.Height - .Height - 50
    End With
    
    With DiagnosticMessagesText
        .Height = MessagesFrame.Height - 1000
        .Width = MessagesFrame.Width - DiagnosticMessagesText.Left * 2
    End With

    With ParametersToolbar
        .Top = lngTabHeight - 1000
    End With
    
    With ParameterSpread
        .Height = lngTabHeight - 1500
        .Width = lngTabWidth - (60 * 2)
    End With
    
    Dim intIndex As Integer
    
    With GraphFrame
        For intIndex = 0 To 3
            .Item(intIndex).Width = lngTabWidth / 2 - 50
            If intIndex = 1 Or intIndex = 3 Then
                .Item(intIndex).Left = .Item(intIndex - 1).Left + .Item(intIndex - 1).Width
            End If
            If intIndex = 2 Or intIndex = 3 Then
                .Item(intIndex).Top = .Item(0).Top + .Item(0).Height
            End If
            .Item(intIndex).Height = lngTabHeight / 2 - 250
        Next
    End With

    With GraphChart
        For intIndex = 0 To 3
            .Item(intIndex).Width = GraphFrame(0).Width - .Item(intIndex).Left * 2
        Next
        For intIndex = 0 To 3
            .Item(intIndex).Height = GraphFrame(0).Height - fraControlFrame(intIndex).Height - 100
        Next
    End With

    With fraControlFrame
        For intIndex = 0 To 3
            .Item(intIndex).Top = GraphFrame(0).Height - .Item(intIndex).Height - 50
        Next
        For intIndex = 0 To 3
            .Item(intIndex).Width = GraphChart(0).Width
        Next
    End With
    On Error GoTo 0
    
End Sub

Private Sub GraphParameterCombo_Click(index As Integer)
        '<EhHeader>
        On Error GoTo GraphParameterCombo_Click_Err
        '</EhHeader>

        Dim iCol As Long
    '    Dim iMask As eNodeMask
    '    Dim dMax As Double
    '    Dim iRow As Long

100     With SetupGraphCodes.GraphCodesSpread
            '---- Determine source node column and comm mask
    '        iRow = 0
102         .Row = 0
            ' MCS
            ' Get the Source Column
104         iCol = GraphSourceCombo(index).ItemData(GraphSourceCombo(index).ListIndex)
            ' Find the Mask value for the Col
    '        Select Case iCol
    '            Case 8
    '                iMask = eNodeMask.LEFT_FRONT_NODE_MASK
    '            Case 9
    '                iMask = eNodeMask.RIGHT_FRONT_NODE_MASK
    '            Case 10
    '                iMask = eNodeMask.LEFT_REAR_NODE_MASK
    '            Case 11
    '                iMask = eNodeMask.RIGHT_REAR_NODE_MASK
    '            Case 12
    '                iMask = eNodeMask.LEFT_TAG_NODE_MASK
    '            Case 13
    '                iMask = eNodeMask.RIGHT_TAG_NODE_MASK
    '            Case 14
    '                iMask = eNodeMask.CCM_NODE_MASK
    '        End Select
        
            ' Get the Row the Parameter is located
106         .Row = GraphParameterCombo(index).ItemData(GraphParameterCombo(index).ListIndex)
108         .Col = 4
110         g_intGraphDivisor(index) = Val(Mid$(.Text, InStr(1, .Text, ".") + 1, 1))
112         .Col = 5                                    'Command # in column 5
114         g_bytGraphCmd(index) = CByte(Val(.Text))
116         .Col = 6
118         g_adblMinValue(index) = Val(.Text)
120         .Col = 7
122         g_adblMaxValue(index) = Val(.Text)
            ' MCS Units 1/20/2003
124         .Col = 3
126         m_astrUnits(index) = .Text
128         GraphChart(index).Axis(AXIS_Y).Title = m_astrUnits(index)
        End With

130     With GraphChart(index)
132         If GraphScaleCombo(index).Text = "Auto" Then
134             .Axis(AXIS_Y).AutoScale = True
136             .RecalcScale
138             .Axis(AXIS_Y).STEP = (.Axis(AXIS_Y).Max / 3)
            Else
140             .Axis(AXIS_Y).AutoScale = False
                ' Set the Graph Minimum Value
142             If g_adblMinValue(index) = 0 Then
144                 .Axis(AXIS_Y).Min = (g_adblMinValue(index))
                Else
146                 .Axis(AXIS_Y).Min = (g_adblMinValue(index) / Val(GraphScaleCombo(index).Text))
                End If
                ' Set the Graph Maximum Value
148             If g_adblMaxValue(index) = 0 Then
150                 .Axis(AXIS_Y).Max = (g_adblMaxValue(index))
                Else
152                 .Axis(AXIS_Y).Max = (g_adblMaxValue(index) / Val(GraphScaleCombo(index).Text))
                End If
                ' Set the Steps
154             If g_adblMaxValue(index) = 0 Then
156                 .Axis(AXIS_Y).STEP = (g_adblMinValue(index) / Val(GraphScaleCombo(index).Text)) / 3
                Else
158                 .Axis(AXIS_Y).STEP = (g_adblMaxValue(index) / Val(GraphScaleCombo(index).Text)) / 3
                End If
            End If
        End With
    
        '<EhFooter>
        Exit Sub

GraphParameterCombo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.GraphParameterCombo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' MCS
Private Sub GraphScaleCombo_Click(index As Integer)
        '<EhHeader>
        On Error GoTo GraphScaleCombo_Click_Err
        '</EhHeader>

    '    Dim iCol As Long
    '    Dim iMask As eNodeMask
    '    Dim dMax As Double

100     If Not blnLoading Then
102         With SetupGraphCodes.GraphCodesSpread
                '---- Determine source node column and comm mask
    '            Dim iRow As Long
    '            iRow = 0
104             .Row = 0
                ' MCS
    '            iCol = GraphSourceCombo(index).ItemData(GraphSourceCombo(index).ListIndex)
    '            For iCol = 0 To .MaxCols
    '                .Col = iCol
    '                If UCase$(.Text) = UCase$(GraphSourceCombo(Index).Text) Then
    '                    g_intGraphSource(Index) = iCol                'Source column in graph grid
    '                    Exit For
    '                End If
    '            Next
    '            Select Case iCol
    '                Case 7
    '                    iMask = eNodeMask.LEFT_FRONT_NODE_MASK
    '                Case 8
    '                    iMask = eNodeMask.RIGHT_FRONT_NODE_MASK
    '                Case 9
    '                    iMask = eNodeMask.LEFT_REAR_NODE_MASK
    '                Case 10
    '                    iMask = eNodeMask.RIGHT_REAR_NODE_MASK
    '                Case 11
    '                    iMask = eNodeMask.LEFT_TAG_NODE_MASK
    '                Case 12
    '                    iMask = eNodeMask.RIGHT_TAG_NODE_MASK
    '                Case 13
    '                    iMask = eNodeMask.CCM_NODE_MASK
    '            End Select

                '---- Determine parameter row
    '            For iRow = 1 To .MaxRows
    '                .Row = iRow
    '                .Col = 1
    '                If UCase$(Trim$(.Text)) = UCase$(Trim$(GraphParameterCombo(Index).Text)) Then
    '                    iGraph1RowNum = .Row
    '                    Exit For
    '                End If
    '            Next
106             .Row = GraphParameterCombo(index).ItemData(GraphParameterCombo(index).ListIndex)
108             .Col = 4
110             g_intGraphDivisor(index) = Val(Mid$(.Text, InStr(1, .Text, ".") + 1, 1))
112             .Col = 5                                    'Command # in column 5
114             g_bytGraphCmd(index) = CByte(Val(.Text))
    '            .Col = 5
    '            dMin = Val(.Text)
    '            .Col = 7
    '            dMax = Val(.Text)
                ' MCS Units 1/20/2003
116             .Col = 3
118             GraphChart(index).Axis(AXIS_Y).Title = .Text
            End With

120         With GraphChart(index)
122             If GraphScaleCombo(index).Text = "Auto" Then
124                 .Axis(AXIS_Y).AutoScale = True
126                 .RecalcScale
128                 .Axis(AXIS_Y).STEP = (.Axis(AXIS_Y).Max / 3)
                Else
130                 .Axis(AXIS_Y).AutoScale = False
                    ' Set the Graph Minimum Value
132                 .Axis(AXIS_Y).Min = (g_adblMinValue(index) / Val(GraphScaleCombo(index).Text))
                    ' Set the Graph Maximum Value
134                 .Axis(AXIS_Y).Max = (g_adblMaxValue(index) / Val(GraphScaleCombo(index).Text))
                    ' Set the Steps
136                 .Axis(AXIS_Y).STEP = (g_adblMaxValue(index) / Val(GraphScaleCombo(index).Text)) / 3
                End If
            End With
        End If
    
        '<EhFooter>
        Exit Sub

GraphScaleCombo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.GraphScaleCombo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GraphSourceCombo_Click(index As Integer)
        '<EhHeader>
        On Error GoTo GraphSourceCombo_Click_Err
        '</EhHeader>

        Dim RowNum As Integer
        Dim ColNum As Integer
        Dim myIndex As Integer
        
        myIndex = GraphParameterCombo(index).ListIndex
        
100     GraphParameterCombo(index).Clear
        ' Prevent the user from clicking on the parameter while fill it
102     GraphParameterCombo(index).Enabled = False      'MCS
    
104     With SetupGraphCodes.GraphCodesSpread
106         .Row = 0
            ' Get the column number form the Source ComboBox
108         ColNum = GraphSourceCombo(index).ItemData(GraphSourceCombo(index).ListIndex)
            ' Store the Source
110         Select Case ColNum
                Case 8
112                 g_intGraphSource(index) = eNodeMask.LEFT_FRONT_NODE_MASK
114             Case 9
116                 g_intGraphSource(index) = eNodeMask.RIGHT_FRONT_NODE_MASK
118             Case 10
120                 g_intGraphSource(index) = eNodeMask.LEFT_REAR_NODE_MASK
122             Case 11
124                 g_intGraphSource(index) = eNodeMask.RIGHT_REAR_NODE_MASK
126             Case 12
128                 g_intGraphSource(index) = eNodeMask.LEFT_TAG_NODE_MASK
130             Case 13
132                 g_intGraphSource(index) = eNodeMask.RIGHT_TAG_NODE_MASK
134             Case 14
136                 g_intGraphSource(index) = eNodeMask.CCM_NODE_MASK
            End Select

            ' Fill in the Parameter ComboBox
138         For RowNum = 1 To .MaxRows                      'For each row in the graphing spreadsheet
140             .Row = RowNum                               'Set row
142             .Col = 1
144             If .Text = 1 Then
146                 .Col = ColNum                               'Set column to currently selected column
148                 If Trim$(UCase$(.Text)) <> "XXX" Then       'If the graph parameter is available
150                     .Col = 2                                'Go to label column
152                     GraphParameterCombo(index).AddItem .Text
                        ' MCS Put the row number in the Item Data
154                     GraphParameterCombo(index).ItemData(GraphParameterCombo(index).NewIndex) = .Row
                    End If
                End If
                ' RDR May cause Problems
    '            DoEvents 'mcs
            Next
        End With
156     GraphParameterCombo(index).Enabled = True       'MCS
158     GraphParameterCombo(index).ListIndex = 0
    
    
        'RDR
        If myIndex > 0 Then
            If GraphParameterCombo(index).ListIndex <= GraphParameterCombo(index).ListCount Then
                GraphParameterCombo(index).ListIndex = myIndex
            End If
        End If
        
        ' Remember last item clicked
    '    If ColNum <> 13 And intPreviousIndex <> -1 Then
    '        GraphParameterCombo(Index).ListIndex = intPreviousIndex
    '        Call GraphParameterCombo_Click(Index)
    '    Else
    '        If GraphParameterCombo(Index).ListCount <> 0 Then
    '            GraphParameterCombo(Index).ListIndex = 0
    '            Call GraphParameterCombo_Click(Index)
    '        End If
    '    End If
    
        '<EhFooter>
        Exit Sub

GraphSourceCombo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.GraphSourceCombo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GraphStartButton_Click(index As Integer)
        '<EhHeader>
        On Error GoTo GraphStartButton_Click_Err
        '</EhHeader>

    '    Dim iCol As Long
    '    Dim iRow As Long
    '    Dim intIndex As Integer

100     If g_blnGraphEnable(index) = False Then
102         GraphStartButton(index).Caption = "Stop"
            ' MCS 1/19/2003
            ' Disable the Controls for the Graph
    '                GraphScaleCombo(Index).Enabled = False
104         GraphParameterCombo(index).Enabled = False
106         GraphSourceCombo(index).Enabled = False

108         With GraphChart(index)
    '            .Gallery = LINES
    '            .ClearData CD_DATA
    '            .TypeMask = CT_EVENSPACING
    '            .LineWidth = 2
    '            .MaxValues = 100
    '            .RealTimeStyle = CRT_LOOPPOS Or CRT_NOWAITARROW
    '            .MenuBar = False
    '            .Toolbar = False
    '            .ShowTips = False
    '            .Axis(AXIS_Y).Grid = True

                ' MCS Added Autoscale Option 1/20/2003
110             If GraphScaleCombo(index).Text = "Auto" Then
112                 .Axis(AXIS_Y).AutoScale = True
114                 .RecalcScale
116                 .Axis(AXIS_Y).STEP = (.Axis(AXIS_Y).Max / 3)
                Else
118                 .Axis(AXIS_Y).AutoScale = False
                    ' Set the Graph Minimum Value
120                 .Axis(AXIS_Y).Min = (g_adblMinValue(index) / Val(GraphScaleCombo(index).Text))
                    ' Set the Graph Maximum Value
122                 .Axis(AXIS_Y).Max = (g_adblMaxValue(index) / Val(GraphScaleCombo(index).Text))
                    ' Set the Steps
124                 .Axis(AXIS_Y).STEP = (g_adblMaxValue(index) / Val(GraphScaleCombo(index).Text)) / 3
                End If

126             .MarkerShape = MK_NONE
    '            .RecalcScale
                Dim iValue As Long
128             iValue = 0
130             .OpenDataEx COD_VALUES Or COD_ADDPOINTS, 1, 100
132             For iValue = 1 To .MaxValues
134                 .Value(0) = 0
                Next
136             GraphChart(index).CloseData COD_VALUES Or COD_REALTIME
            End With
        
    '        ViewLog.Log DebugMsg, "Start Broadcast " & GraphSourceCombo(Index).Text & " " & _
    '                        GraphParameterCombo(Index).Text
    '        ' MCS Stop All Broadcasts
    '        For intIndex = 0 To 3
    '            If g_blnGraphEnable(intIndex) Then
    '                Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket( _
    '                        eDataCommand.eStopBroadcast, g_intGraphSource(intIndex), _
    '                        g_bytGraphCmd(intIndex), intIndex))
    '                ' RDR Removed May Cause Problems
    '                'DoEvents
    '            End If
    '        Next intIndex
            ' Start the Channel
138         g_blnGraphEnable(index) = True
            ' MCS ReStart Broadcasts
140         Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket( _
                    eDataCommand.eStartBroadcast, g_intGraphSource(index), _
                    g_bytGraphCmd(index), index))
    '
    '        For intIndex = 0 To 3
    '            If g_blnGraphEnable(intIndex) Then
    '                Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket( _
    '                        eDataCommand.eStartBroadcast, g_intGraphSource(intIndex), _
    '                        g_bytGraphCmd(intIndex), intIndex))
    '                ' RDR Removed May Cause Problems
    '                'DoEvents
    '            End If
    '        Next intIndex
        Else
142         GraphStartButton(index).Caption = "Start"
    '        ViewLog.Log DebugMsg, "Stop Broadcast " & GraphSourceCombo(Index).Text & " " & _
    '                        GraphParameterCombo(Index).Text
    '        ' MCS Stop All Broadcasts
    '        For intIndex = 0 To 3
    '            If g_blnGraphEnable(intIndex) Then
    '                Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket( _
    '                            eDataCommand.eStopBroadcast, g_intGraphSource(intIndex), _
    '                            g_bytGraphCmd(intIndex), intIndex))
    '                ' RDR Removed May Cause Problems
    '                'DoEvents
    '            End If
    '        Next intIndex
    '        ' Stop the Broadcast
144         g_blnGraphEnable(index) = False
146         Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket( _
                        eDataCommand.eStopBroadcast, g_intGraphSource(index), _
                        g_bytGraphCmd(index), index))

    '        ' MCS ReStart Broadcasts
    '        For intIndex = 0 To 3
    '            If g_blnGraphEnable(intIndex) Then
    '                Call SetupSerialPort.ComSend(SetupSerialPort.CreateBroadcastPacket( _
    '                        eDataCommand.eStartBroadcast, g_intGraphSource(intIndex), _
    '                        g_bytGraphCmd(intIndex), intIndex))
    '                ' RDR Removed May Cause Problems
    '                'DoEvents
    '            End If
    '        Next intIndex
            ' MCS 1/19/2003
            ' Disable the Controls for the Graph
148         GraphScaleCombo(index).Enabled = True
150         GraphParameterCombo(index).Enabled = True
152         GraphSourceCombo(index).Enabled = True
        End If

        '<EhFooter>
        Exit Sub

GraphStartButton_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.GraphStartButton_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub MainTab_Click(PreviousTab As Integer)
        '<EhHeader>
        On Error GoTo MainTab_Click_Err
        '</EhHeader>
    
100     Select Case PreviousTab
            Case 0
102             ConfigurationSpread.Visible = False
104             MessagesFrame.Visible = False
106         Case 1
108             ParameterSpread.Visible = False
        End Select
    
110     Select Case MainTab.Tab
            Case 0
112             ConfigurationSpread.Visible = True
114             MessagesFrame.Visible = True
116         Case 1
118             ParameterSpread.Visible = True
        End Select
    
        '<EhFooter>
        Exit Sub

MainTab_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.MainTab_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' MCS Added to Send Changed Parameter on the Parameter Spreadsheet
' 2/12/2003
Private Sub ParameterSpread_Change(ByVal Col As Long, ByVal Row As Long)
        '<EhHeader>
        On Error GoTo ParameterSpread_Change_Err
        '</EhHeader>
        Dim Min As Double, Max As Double

        With ParameterSpread
            .Row = Row
            .Col = 5
            Min = Val(.Text)
            .Col = 6
            Max = Val(.Text)
            .Col = Col
            If Val(.Text) < Min Or Val(.Text) > Max Then
                MsgBox "Parameter out of range."
                .Text = ""
            Else
100             Call SetParameterOnCCM(Row, Col)
            End If
        End With



    
        '<EhFooter>
        Exit Sub

ParameterSpread_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.ParameterSpread_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ParametersToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo ParametersToolbar_ButtonClick_Err
        '</EhHeader>

100     Select Case Button.index
            ' Get All
            Case 1
102             SetupSerialPort.GetAllParameters
            ' Put All
104         Case 2

106             Call PutAllParameters
    '            ViewLog.Log DebugMsg, "eKonect Test Communications"
    '            Call SetupSerialPort.ComSend(SetupSerialPort.CreateCommandPacket(eDataCommand.eKonect, eNodeMask.CCM_NODE_MASK, 0, 0))
        End Select

        '<EhFooter>
        Exit Sub

ParametersToolbar_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.ParametersToolbar_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar1_ButtonClick_Err
        '</EhHeader>

100     With Button
102         Select Case .index
                ' New
                Case 1
            
                ' Open
104             Case 2
106                 Call FileOpen
                ' Save
108             Case 3
110                 Call FileSave
                ' Print
112             Case 4
114                 PrintPreview.PrepareToPreview ParameterSpread
                '??
116             Case 5
            
            End Select
        End With

        '<EhFooter>
        Exit Sub

Toolbar1_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.Toolbar1_ButtonClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ToolsMenu_Click(index As Integer)
        '<EhHeader>
        On Error GoTo ToolsMenu_Click_Err
        '</EhHeader>
    
100     Select Case index
            Case 0
102             ViewLog.Show vbModal
104         Case 1
106             SetupDiagnosticCodes.Show vbModal
108         Case 2
                'graph Codes
110             SetupGraphCodes.Show vbModal
112         Case 3
114             SetupSerialPort.Show vbModal
116         Case 4
118             SetupUsers.Show vbModal
120         Case 6 ' Download Firmware
122             FirmwareDownload.Show vbModal
        End Select

        '<EhFooter>
        Exit Sub

ToolsMenu_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.ToolsMenu_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OpenConfiguration()
        '<EhHeader>
        On Error GoTo OpenConfiguration_Err
        '</EhHeader>

        Dim iRow As Integer
    
100     With ConfigurationSpread
102         .LoadTextFile "Configuration.csv", ",", ",", vbCrLf, LoadTextFileColHeaders, "ConfigurationErrors.txt"
104         .ColWidth(1) = 10
106         .ColWidth(2) = 5
108         .ColWidth(3) = 5
110         .ColWidth(4) = 5
112         .ColWidth(5) = 5
114         .ColWidth(6) = 5
116         .ColWidth(7) = 20
118         .LockBackColor = vbCyan
120         For iRow = 1 To .MaxRows
122             .Row = iRow
124             .Col = 1
126             .Lock = True
128             .Col = 2
130             .CellType = CellTypeCheckBox
132             .TypeCheckType = TypeCheckTypeNormal
134             .TypeCheckCenter = True
136             If Splash.iCurrentLevel = 2 Then
138                 .Lock = False
                Else
140                 .Lock = True
                End If
142             .Col = 3
144             .CellType = CellTypeCheckBox
146             .TypeCheckType = TypeCheckTypeNormal
148             .TypeCheckCenter = True
150             If Splash.iCurrentLevel = 2 Then
152                 .Lock = False
                Else
154                 .Lock = True
                End If
156             .Col = 4
158             .CellType = CellTypeCheckBox
160             .TypeCheckType = TypeCheckTypeNormal
162             .TypeCheckCenter = True
164             If Splash.iCurrentLevel = 2 Then
166                 .Lock = False
                Else
168                 .Lock = True
                End If
170             .Col = 5
172             .CellType = CellTypeCheckBox
174             .TypeCheckType = TypeCheckTypeNormal
176             .TypeCheckCenter = True
178             If Splash.iCurrentLevel = 2 Then
180                 .Lock = False
                Else
182                 .Lock = True
                End If
184             .Col = 6
186             .Lock = True
188             .Col = 7
190             .Lock = True
            Next
        End With
    
        '<EhFooter>
        Exit Sub

OpenConfiguration_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.OpenConfiguration " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub FileOpen()

    On Error GoTo ErrorHandler
    
'    Dim RecordNum As Integer                                            'Counter for Packages & Frames
'    Dim ColumnNum As Integer                                            'Counter for spreadsheet column
    
    ViewLog.Log DebugMsg, Me.Name & ":OpenFile()"
    With CommonDialog
        .DefaultExt = ".CSV"
        .Filter = "Comma Separated Variable(*.csv)|*.csv"
        .FilterIndex = 1
        .CancelError = False
        .DialogTitle = "Open File"
        .ShowOpen
        If .FileName <> "" Then                                         'If the user chose a file to open
            OpenTextFile .FileName
        End If
    End With
    ViewLog.Log DebugMsg, Me.Name & ":OpenFile Exiting"
    
    Exit Sub
    
ErrorHandler:

    ErrorForm.ReportError Me.Name & ":OpenFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
    
End Sub

Private Sub OpenTextFile(ByVal sName As String)
        '<EhHeader>
        On Error GoTo OpenTextFile_Err
        '</EhHeader>
    
        Dim iCol As Integer, iRow As Integer, DecPlaces As Integer, Pos As Integer, MinValue As Double, MaxValue As Double
    
100     With ParameterSpread
102         .LoadTextFile sName, ",", ",", vbCrLf, LoadTextFileColHeaders, "test.txt"
104         For iRow = 1 To .MaxRows
106             For iCol = 0 To .MaxCols
108                 .Row = iRow
                    .Col = 3
                    Pos = InStr(1, .Text, ".")
                    If Pos > 0 Then
                        DecPlaces = Val(Mid(.Text, Pos + 1, 1))
                    Else
                        DecPlaces = 0
                    End If
110                 .Col = iCol
112                 .EditModeReplace = True
114                 If Trim$(UCase$(.Text)) = "XXX" Or .Col < 7 Then
116                     .BackColor = vbCyan
118                     .Lock = True
                    Else
                        '---- Set numeric cell types
120                     If Splash.iCurrentLevel = 2 Then
122                         .Lock = False
                            
123                         .Col = 5
124                          MinValue = Val(.Text)

125                         .Col = 6
126                          MaxValue = Val(.Text)

127                         .Col = iCol
                            
128                         .CellType = CellTypeNumber
129                         .TypeNumberDecPlaces = DecPlaces
130                         .TypeNumberMin = MinValue
131                         .TypeNumberMax = MaxValue
                            
                        Else
132                      .Lock = True
                        End If
                    End If

                Next
            Next
        
            ' Removed Checkboxes MCS 2/12/2003
    '        .MaxCols = .MaxCols + 1
    '        .InsertCols 2, 1
    '        .Col = 2
    '        .Row = 0
    '        .Text = "Track"
    '        .ColsFrozen = 1
    '        For iRow = 1 To .MaxRows
    '            .Row = iRow
    '            .CellType = CellTypeCheckBox
    '            .TypeCheckType = TypeCheckTypeNormal
    '            .TypeCheckCenter = True
    '            .Text = "0"
    '        Next
133         .ColWidth(0) = 4
134         .ColWidth(1) = 15
    '        .ColWidth(2) = 7
        End With

        '<EhFooter>
        Exit Sub

OpenTextFile_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in ActiveRide.MainForm.OpenTextFile " & _
'               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ClearParameters()
        '---- RDR Clear the parameters data (these will be overwritten by module)
        Dim iCol As Integer, iRow As Integer
    
100     With ParameterSpread
104         For iRow = 1 To .MaxRows                'Do not clear row 0!!!  Row 0 is used for column headers, which are used for NodeMask
                .Row = iRow
                For iCol = 7 To 13
                    .Col = iCol
                    If .Lock = False Then
                        .Value = 0
                    End If
                Next
            Next
        End With
End Sub



Private Sub FileSave()
    On Error GoTo ErrorHandler
    
    ViewLog.Log DebugMsg, Me.Name & ":FileSave()"
    
    With CommonDialog                                               'Windows Common control for Save As
        .DefaultExt = ".CSV"
        .Filter = "Comma Separated Variable(*.csv)|*.csv"
        ' MCS Changed to True 2/24/2003
        .CancelError = True
        .DialogTitle = "Save File"
        .ShowSave                                                   'Show the Save As dialog
        If .FileName <> "" Then                  'If the user chose a file
            SaveTextFile .FileName
        End If
    End With
    
    ViewLog.Log DebugMsg, Me.Name & ":FileSave Exiting"
    
    Exit Sub

ErrorHandler:

    ' MCS 2/24/2003
    If Err.Number = cdlCancel Then
        ' Save Canceled
    Else
        ErrorForm.ReportError Me.Name & ":FileSave", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
    End If
    
End Sub

Private Sub SaveTextFile(ByVal sName As String)
        '<EhHeader>
        On Error GoTo SaveTextFile_Err
        '</EhHeader>
    
        ' Removed 6/19/2003 MCS
    '    ParameterSpread.DeleteCols 2, 1
100     ParameterSpread.ExportToTextFile sName, ",", ",", vbCrLf, ExportToTextFileColHeaders, "ParameterErrors.txt"
102     OpenTextFile sName

        '<EhFooter>
        Exit Sub

SaveTextFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.SaveTextFile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub UpdateStatus()
        '<EhHeader>
        On Error GoTo UpdateStatus_Err
        '</EhHeader>
    
100     With StatusBar
102         .Panels(3).Text = "#Tx:" & SetupSerialPort.PacketsSent
104         .Panels(4).Text = "#Rv:" & SetupSerialPort.PacketsReceived
106         .Panels(5).Text = "#Er:" & SetupSerialPort.PacketErrors
108         .Refresh ' 3/25/2003
        End With

        '<EhFooter>
        Exit Sub

UpdateStatus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.UpdateStatus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetParameter(ByVal intID As Integer, ByVal bytTarget As eNodeMask)
    
    Dim bytCmd As Byte
    Dim iData As Currency
    Dim intDecimalPlaces As Integer
    Dim intRowIndex As Integer
    
'    iDecimalPlaces = 0
    
    ' Set the ID in the SetupGraphCodes Form
'    SetupGraphCodes.ParameterID = intID
'    ' Get the DecimalPlace
'    intDecimalPlaces = SetupGraphCodes.DecimalPlaces
    bytCmd = CByte(intID)
    
'    Call SetupSerialPort.ComSend(SetupSerialPort.CreateGetPacket(eGetParameter, bytTarget, bytCmd))
    Call SetupSerialPort.ComSend(SetupSerialPort.CreateCommandPacket(eGetParameter, bytTarget, bytCmd, 0))

End Sub

' MCS
Private Sub SetParameterOnCCM(ByVal iRow As Long, ByVal iCol As Long)
        '<EhHeader>
        On Error GoTo SetParameterOnCCM_Err
        '</EhHeader>
    
        Dim bTarget As Byte
        Dim bytCmd As Byte
        Dim iData As Currency
        Dim iDecimalPlaces As Long
    
100     iDecimalPlaces = 0
    
    '    SetupSerialPort.CommTimer.Enabled = False

102     With ParameterSpread
104         .Row = iRow
            ' Format
106         .Col = m_cintFormatParam
108         iDecimalPlaces = Val(Mid$(.Text, InStr(1, .Text, ".") + 1, 1))
            ' Column Sent into the command
110         .Col = iCol
112         If UCase$(.Text) = "XXX" Then
                Exit Sub
            End If
114         iData = Val(.Text) * (10 ^ iDecimalPlaces)
            ' Parameter ID
116         .Col = m_cintIDParam
118         bytCmd = CByte(Val(.Text))
            ' Column sent into the command
120         .Col = iCol
122         .Row = 0
            Dim sColumn As String
124         sColumn = .Text
126         Select Case sColumn
                Case "Left Front"
128                 bTarget = eNodeMask.LEFT_FRONT_NODE_MASK
130             Case "Left Rear"
132                 bTarget = eNodeMask.LEFT_REAR_NODE_MASK
134             Case "Left Tag"
136                 bTarget = eNodeMask.LEFT_TAG_NODE_MASK
138             Case "Right Front"
140                 bTarget = eNodeMask.RIGHT_FRONT_NODE_MASK
142             Case "Right Rear"
144                 bTarget = eNodeMask.RIGHT_REAR_NODE_MASK
146             Case "Right Tag"
148                 bTarget = eNodeMask.RIGHT_TAG_NODE_MASK
150             Case "Central"
152                 bTarget = eNodeMask.CCM_NODE_MASK
            End Select
        End With
    
154     Call SetupSerialPort.ComSend(SetupSerialPort.CreateCommandPacket(eDataCommand.eSetParamater, bTarget, bytCmd, iData))
    
    '    SetupSerialPort.CommTimer.Enabled = True
    
        ' RDR May Cause Problems
        'DoEvents
    
        '<EhFooter>
        Exit Sub

SetParameterOnCCM_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.SetParameterOnCCM " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' MCS
Public Sub PutAllParameters()
        '<EhHeader>
        On Error GoTo PutAllParameters_Err
        '</EhHeader>
        
        Dim RowNum As Integer
        Dim ColNum As Integer
    
        '---- RDR System locks up when putting all while graphing... Added code to check status
        Dim GraphNum As Integer
        For GraphNum = 0 To 3
            If g_blnGraphEnable(GraphNum) = True Then
                MsgBox "Please stop graphing before putting all parameters."
                Exit Sub
            End If
        Next
    
    
    
100     If MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Put ALL parameters?") = vbYes Then
            '---- Loop for all rows
102         For RowNum = 1 To ParameterSpread.MaxRows
104             For ColNum = 7 To ParameterSpread.MaxCols
106                 Call SetParameterOnCCM(RowNum, ColNum)
                Next
            Next
108         MsgBox "Completed.", vbApplicationModal + vbOKOnly, "Put All"

        End If

        '<EhFooter>
        Exit Sub

PutAllParameters_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.PutAllParameters " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       ActiveRide
' Procedure  :       Calibrate
' Description:       [type_description_here]
' Created by :       Mark Saur
' Machine    :       SUSPENSION_RIDE
' Date-Time  :       10/08/2003-10:28:25
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub Calibrate()
    '<EhHeader>
    On Error GoTo Calibrate_Err
    '</EhHeader>

    Dim intRowIndex As Integer
        
        Dim RowNum As Long
        Dim ColNum As Integer
    
            '---- Loop for all rows (when calibrate id is found for CCM, set it to 1)
            ParameterSpread.Col = 4
102         For RowNum = 1 To ParameterSpread.MaxRows
                ParameterSpread.Row = RowNum
                If ParameterSpread.Value = 222 Then         'Docs say 110
                    ParameterSpread.Col = 13
                    ParameterSpread.Value = 1
106                 Call SetParameterOnCCM(RowNum, 13)
                    Exit For
                End If
            Next

        
    ' Store the Variables before Calibration
    'GetParameter 3, CCM_NODE_MASK
    'GetParameter 8, CCM_NODE_MASK
    'GetParameter 16, CCM_NODE_MASK
    
    
    MsgBox "Calibration complete."
        
        
'    With ParameterSpread
'        .Col = m_cintParameterParam
'        For intRowIndex = 1 To .MaxRows
'            .Row = intRowIndex
'            If UCase$(.Text) = "CALIBRATE" Then
'                Call SetParameterOnCCM(.Row, 14 )
'                Exit For
'            End If
'        Next
'    End With
    '<EhFooter>
    Exit Sub

Calibrate_Err:
    MsgBox Err.Description & vbCrLf & _
           "in ActiveRide.MainForm.Calibrate " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub Configuration(lngRow As Long, lngCol As Long)
        '<EhHeader>
        On Error GoTo Configuration_Err
        '</EhHeader>

        Dim bytTarget As Byte
        Dim bytCommand As Byte
        Dim intData As Integer
    
100     With ConfigurationSpread
102         .Row = lngRow
            ' Command Column
104         .Col = 6
106         bytCommand = CByte(Val(.Text))
        
            ' Get the Data
108         .Col = lngCol
110         intData = Val(.Text)
        
            ' Get the Node
112         Select Case lngCol
                Case 2
114                 bytTarget = eNodeMask.LEFT_FRONT_NODE_MASK
116             Case 4
118                 bytTarget = eNodeMask.LEFT_REAR_NODE_MASK
120             Case 3
122                 bytTarget = eNodeMask.RIGHT_FRONT_NODE_MASK
124             Case 5
126                 bytTarget = eNodeMask.RIGHT_REAR_NODE_MASK
            End Select
        End With
    
128     Call SetupSerialPort.ComSend(SetupSerialPort.CreateCommandPacket(eDataCommand.eSetParamater, bytTarget, bytCommand, intData))

        '<EhFooter>
        Exit Sub

Configuration_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.MainForm.Configuration " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
