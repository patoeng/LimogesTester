VERSION 5.00
Begin VB.Form FRM_Attente 
   BorderStyle     =   0  'None
   Caption         =   "Demande entree C4"
   ClientHeight    =   510
   ClientLeft      =   135
   ClientTop       =   4065
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "FRM_ARU"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   510
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.Label LBL_Attente 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waiting Autorization Sorting Stat C4"
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
      Width           =   6585
   End
End
Attribute VB_Name = "FRM_Attente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
