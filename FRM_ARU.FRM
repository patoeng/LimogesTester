VERSION 5.00
Begin VB.Form FRM_ARU 
   BorderStyle     =   0  'None
   Caption         =   "FRM_ARU"
   ClientHeight    =   525
   ClientLeft      =   1380
   ClientTop       =   6165
   ClientWidth     =   5790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "FRM_ARU"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   525
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LBL_ARU 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMERGENCY EXIT Pressed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "FRM_ARU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LBL_ARU_Click()
    symb_com = ""
    FRM_startup!Lbl_code_lu.Caption = ""
End Sub


