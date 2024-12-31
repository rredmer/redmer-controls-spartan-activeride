VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton LoginButton 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   795
      Index           =   1
      Left            =   4740
      Picture         =   "Splash.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2550
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   9600
      Begin VB.CommandButton LoginButton 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   795
         Index           =   0
         Left            =   3720
         Picture         =   "Splash.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3750
         TabIndex        =   0
         Top             =   1980
         Width           =   1725
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   30
         Picture         =   "Splash.frx":0624
         ScaleHeight     =   1275
         ScaleWidth      =   2325
         TabIndex        =   7
         Top             =   150
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         TabIndex        =   8
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "(c) 1992-2002 Redmer Controls Inc.  All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   3930
         Width           =   6555
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1020
         Width           =   4425
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Windows 2000,XP, && .NET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   660
         Width           =   3480
      End
      Begin VB.Label lblLicenseTo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2460
         TabIndex        =   2
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spartan Motors ActiveRide"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public sCurrentUser As String
Public iCurrentLevel As Integer


Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
100     lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.Splash.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub LoginButton_Click(index As Integer)
        '<EhHeader>
        On Error GoTo LoginButton_Click_Err
        '</EhHeader>
        
100     Select Case index
            Case 0
102             End
104         Case 1
106             iCurrentLevel = SetupUsers.GetUserSecurityLevel(txtUserID.Text)
108             If iCurrentLevel = -1 Then
110                 MsgBox "User Id is not valid.", vbApplicationModal + vbOKOnly + vbExclamation, "Warning!"
112                 txtUserID.Text = ""
114                 txtUserID.SetFocus
                Else
    '                sCurrentUser = Trim(UCase(txtUserID.Text))
116                 Me.Hide
                End If

        End Select
        '<EhFooter>
        Exit Sub

LoginButton_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ActiveRide.Splash.LoginButton_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
