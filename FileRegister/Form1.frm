VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test Form [ QWE Reader ]"
   ClientHeight    =   4920
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReg 
      Caption         =   "Register File"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      Begin VB.CheckBox chkOver 
         Caption         =   "OverWrite if already exists!!!"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtIcon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   16
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Text            =   "QWE Document File"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtAppName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   8
         Text            =   "QWE Reader"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtAppPath 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtExt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   6
         Text            =   ".qwe"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Icon File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXE Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXE Path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Extension"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1170
      End
   End
   Begin VB.Frame frResult 
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   6735
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jim Jose :-))"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Friend,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0D4A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1080
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   6090
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "Form1.frx":0E28
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please complie the App and run the EXE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   3435
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdReg_Click()
On Error GoTo Handle
Dim Rtn As Boolean

    '--Register the file
    Rtn = RegisterFile(txtExt, txtDesc, txtAppName, txtAppPath, txtIcon, chkOver)
    
    
    If Rtn = True Then
        
        '--Create temporary document file
        Dim mFile As String
        If Right$(App.Path, 1) = "\" Then
            mFile = App.Path & "Test" & txtExt
        Else
            mFile = App.Path & "\" & "Test" & txtExt
        End If
        Open mFile For Binary As #1
            Put #1, , "Test File..."
        Close #1
        frResult.ZOrder (0)
        
    End If
    
Exit Sub
Handle:
MsgBox Err.Description, vbCritical, "Unable to register file!!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Right$(App.Path, 1) = "\" Then
        txtAppPath = App.Path & App.EXEName & ".exe"
        txtIcon = App.Path & "Icon.ico"
    Else
        txtAppPath = App.Path & "\" & App.EXEName & ".exe"
        txtIcon = App.Path & "\Icon.ico"
    End If
End Sub
