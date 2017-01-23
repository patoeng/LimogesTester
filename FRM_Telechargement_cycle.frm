VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_Telechargement_cycle 
   Caption         =   "Download test cycles"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   1452
      Width           =   3984
      _ExtentX        =   7038
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3984
      _ExtentX        =   7038
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   10
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loading test cycle on progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3852
   End
End
Attribute VB_Name = "FRM_Telechargement_cycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
    SetWindowPos hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Public Sub Set_avancement(Avancement As Integer, Optional Max As Integer = -10)
    
    If Max > 0 Then
        ProgressBar1.Max = Max
    End If
    If Avancement <= ProgressBar1.Max Then
        ProgressBar1.Value = Avancement
    End If
    
    StatusBar1.SimpleText = Avancement & " / " & ProgressBar1.Max

End Sub
