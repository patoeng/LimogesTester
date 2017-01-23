VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FRM_PARAM 
   Caption         =   "Test Parameter"
   ClientHeight    =   6975
   ClientLeft      =   15
   ClientTop       =   360
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9510
   Begin VB.CommandButton CMD_QUITTER 
      Caption         =   "F7-&QUIT"
      Height          =   495
      Left            =   3945
      TabIndex        =   0
      Top             =   6210
      Width           =   1935
   End
   Begin Threed.SSPanel SSPanel_temp 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   4890
      _Version        =   65536
      _ExtentX        =   8625
      _ExtentY        =   7646
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   6
      Begin VB.TextBox TxT_aff_temp 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   4605
      End
      Begin VB.Label Lbl_temp 
         Alignment       =   2  'Center
         Caption         =   "TEST PARAMETER TEMPERATURE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   4575
      End
   End
   Begin Threed.SSPanel SSPanel_final 
      Height          =   4350
      Left            =   4920
      TabIndex        =   3
      Top             =   1665
      Width           =   4560
      _Version        =   65536
      _ExtentX        =   8043
      _ExtentY        =   7673
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox TxT_aff_final 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4290
      End
      Begin VB.Label Lbl_final 
         Alignment       =   2  'Center
         Caption         =   "TEST PARAMETER FINAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   10
         Top             =   150
         Width           =   4095
      End
   End
   Begin Threed.SSPanel SSPanel_géné 
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4920
      _Version        =   65536
      _ExtentX        =   8678
      _ExtentY        =   2566
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox TxT_aff_géné 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   105
         Width           =   4665
      End
   End
   Begin Threed.SSPanel SSPanel_claq 
      Height          =   1455
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   4515
      _Version        =   65536
      _ExtentX        =   7964
      _ExtentY        =   2566
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   0
      Begin VB.TextBox TxT_aff_claq 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   4290
      End
      Begin VB.Label Lbl_claq 
         Alignment       =   2  'Center
         Caption         =   "TEST PARAMETER HIPOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   120
         Width           =   4245
      End
   End
End
Attribute VB_Name = "FRM_PARAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD_quitter_Click()
    Contexte$ = "STARTUP"
    Unload Me
    FRM_startup.Show
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_F7 Or KeyCode = vbKeyDivide Then
        CMD_quitter_Click
    End If
End Sub

Private Sub Form_Load()
    Call maj_text_param_géné
    Call maj_text_param_claq
    Call maj_text_param_temp
    Call maj_text_param_final
End Sub


