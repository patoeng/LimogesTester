VERSION 5.00
Begin VB.Form FRM_def_chargement 
   Caption         =   "Défaut chargement"
   ClientHeight    =   2280
   ClientLeft      =   576
   ClientTop       =   3216
   ClientWidth     =   8412
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   8412
   Begin VB.CommandButton Cmd_relancer 
      Caption         =   "F7 - Relancer"
      Height          =   495
      Left            =   3348
      TabIndex        =   3
      Top             =   1296
      Width           =   1665
   End
   Begin VB.CommandButton Cmd_valider 
      Caption         =   "F11 - Valider"
      Height          =   510
      Index           =   1
      Left            =   6516
      TabIndex        =   2
      Top             =   1296
      Width           =   1665
   End
   Begin VB.CommandButton Cmd_valider 
      Caption         =   "F1 - Valider"
      Height          =   510
      Index           =   0
      Left            =   288
      TabIndex        =   1
      Top             =   1296
      Width           =   1665
   End
   Begin VB.Label Lbl_message_def 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4068
      TabIndex        =   0
      Top             =   432
      Width           =   120
   End
End
Attribute VB_Name = "FRM_def_chargement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_relancer_Click()

relance1 = True
Unload Me

End Sub

Private Sub Cmd_valider_Click(Index As Integer)

If Index = 0 And Cmd_valider(0).Visible = True Then
    valid1 = True
    Unload Me
End If

If Index = 1 And Cmd_valider(1).Visible = True Then
    valid2 = True
    Unload Me
End If



End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = KEY_F1 Then
    Cmd_valider_Click (0)
    End If
    
    If KeyCode = KEY_F7 Then
    Cmd_relancer_Click
    End If
 
    
   If KeyCode = KEY_F11 Then
    Cmd_valider_Click (1)
    End If


End Sub

Private Sub Form_Load()

aff_def_char = True

If (Tab_Status_banc(67) = 70) Then 'Défaut produit 2
    Cmd_valider(1).Visible = True
    Cmd_valider(0).Visible = False

    If Defaut = 25 Then
    Lbl_message_def.Caption = "Défaut produit 2 en court circuit"
    End If
    If Defaut = 23 Then
    Lbl_message_def.Caption = "Défaut produit 2 : commutation permanente"
    End If
    If Defaut = 27 Then
    Lbl_message_def.Caption = "Défaut produit 2 : non commutation "
    End If
End If

If (Tab_Status_banc(67) = 60) Then 'Défaut produit 1
    Cmd_valider(0).Visible = True
    Cmd_valider(1).Visible = False
    
    If Defaut = 25 Then
    Lbl_message_def.Caption = "Défaut produit 1 en court circuit"
    End If
    If Defaut = 23 Then
    Lbl_message_def.Caption = "Défaut produit 1 : commutation permanente"
    End If
    If Defaut = 27 Then
    Lbl_message_def.Caption = "Défaut produit 1 : non commutation "
    End If
End If

aff_def = 0

End Sub


Private Sub Form_Paint()
'met la fenetre au premier plan, et on ne peut plus l'effacer !
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
End Sub


