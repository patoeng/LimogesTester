VERSION 5.00
Begin VB.Form FRM_Reprise 
   BorderStyle     =   0  'None
   Caption         =   "FRM_Reprise"
   ClientHeight    =   405
   ClientLeft      =   1395
   ClientTop       =   6225
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "FRM_ARU"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   405
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LBL_REP 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bench Initialization...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "FRM_Reprise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ARU_Affiché = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ARU_Affiché = False
End Sub


