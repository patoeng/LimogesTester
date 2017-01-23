VERSION 5.00
Begin VB.Form FRM_Affiche_def 
   Caption         =   "Gestion défaut au poste controle final"
   ClientHeight    =   2520
   ClientLeft      =   1368
   ClientTop       =   3048
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2520
   ScaleWidth      =   7560
   Begin VB.CommandButton Cmd_relancer 
      Caption         =   " F10 - &Relancer le            controle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   1560
      Width           =   2130
   End
   Begin VB.CommandButton Cmd_confirmer 
      Caption         =   "F3 - &Vidage poste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   750
      TabIndex        =   1
      Top             =   1560
      Width           =   2160
   End
   Begin VB.Label Lbl_defaut_final 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3360
      TabIndex        =   0
      Top             =   435
      Width           =   735
   End
End
Attribute VB_Name = "FRM_Affiche_def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_confirmer_Click()
'Vidage poste

set_bit MW350(6), 14

If def_final <> 0 Then
   set_bit MW350(4), 8
   def_final = 0
   MW350(9) = 0
   Write_Buf_W350(9) = 0
End If

For i = 0 To 9
    Write_Buf_W350(i) = MW350(i)
Next



Call ecrit_tsx(350, Write_Buf_W350(), 10)
Unload Me

End Sub

Private Sub Cmd_relancer_Click()

set_bit MW350(6), 13

If def_final <> 0 Then
   set_bit MW350(4), 8
   def_final = 0
   MW350(9) = 0
   Write_Buf_W350(9) = 0
End If

For i = 0 To 9
    Write_Buf_W350(i) = MW350(i)
Next



Call ecrit_tsx(350, Write_Buf_W350(), 10)
Unload Me


End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = KEY_F3 Then
    Cmd_confirmer_Click
End If

If KeyCode = KEY_F10 Then
    Cmd_relancer_Click
End If


End Sub


Private Sub Form_Load()

If def_final <> 0 Then
    Lbl_defaut_final.Caption = mes_def_final
End If

End Sub


Private Sub Form_Paint()
'met la fenetre au premier plan, et on ne peut plus l'effacer !
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW

End Sub


Private Sub Lbl_defaut_ctrl_Click()

End Sub


