VERSION 4.00
Begin VB.Form FRM_Affichage_manuel 
   Caption         =   "Affichage manuel"
   ClientHeight    =   2955
   ClientLeft      =   2085
   ClientTop       =   2490
   ClientWidth     =   8430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   3360
   KeyPreview      =   -1  'True
   Left            =   2025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   8430
   Top             =   2145
   Width           =   8550
   Begin VB.Frame FRA_Mes_p_h 
      Caption         =   "Mesure P && H"
      Enabled         =   0   'False
      Height          =   984
      Left            =   150
      TabIndex        =   3
      Top             =   1125
      Visible         =   0   'False
      Width           =   8184
      Begin VB.CommandButton CMD_P_H 
         Caption         =   "F2 - Mesurer P && H"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   696
         Left            =   180
         TabIndex        =   4
         Top             =   210
         Width           =   1812
      End
      Begin VB.Label LBL_Mes_h 
         Alignment       =   2  'Center
         Caption         =   "Hystérésis mesurée"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   2300
         TabIndex        =   8
         Top             =   666
         Width           =   2100
      End
      Begin VB.Label LBL_Mes_p 
         Alignment       =   2  'Center
         Caption         =   "Portée mesurée"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   2300
         TabIndex        =   7
         Top             =   270
         Width           =   2208
      End
      Begin VB.Label LBL_lim_h 
         Alignment       =   2  'Center
         Caption         =   "Limites d'hystérésis"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   4800
         TabIndex        =   6
         Top             =   666
         Width           =   3216
      End
      Begin VB.Label LBL_lim_p 
         Alignment       =   2  'Center
         Caption         =   "Limites de portée"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   4800
         TabIndex        =   5
         Top             =   270
         Width           =   3144
      End
   End
   Begin VB.CommandButton CMD_mauvais 
      Caption         =   "F5 - Mauvais"
      Height          =   540
      Left            =   5625
      TabIndex        =   2
      Top             =   2295
      Width           =   1380
   End
   Begin VB.CommandButton CMD_bon 
      Caption         =   "F1 - OK"
      Height          =   540
      Left            =   1215
      TabIndex        =   1
      Top             =   2295
      Width           =   1380
   End
   Begin VB.Label LBL_message_a_afficher 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "FRM_Affichage_manuel"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub CMD_bon_Click()
            
            LBL_message_a_afficher.Caption = "OK"
            fra_mes_p_h.Enabled = False
            fra_mes_p_h.Visible = False
            Write_Buf_W350(0) = W350(0)
            reset_bit W350(0), 12 'resultat bon
            reset_bit W350(0), 7 'acquis message
            reset_bit W350(0), 3 'mesure P & H
            reset_bit Write_Buf_W350(0), 12 'resultat bon
            reset_bit Write_Buf_W350(0), 7 'acquis message
            reset_bit Write_Buf_W350(0), 3 'mesure P & H
            Call ecrit_tsx(350, Write_Buf_W350(), 1)
End Sub

Private Sub CMD_mauvais_Click()
            
            LBL_message_a_afficher.Caption = "pas OK"
            fra_mes_p_h.Enabled = False
            fra_mes_p_h.Visible = False
            Write_Buf_W350(0) = W350(0)
            set_bit W350(0), 12 'resultat mauvais
            reset_bit W350(0), 7 'acquis message
            reset_bit W350(0), 3 'mesure P & H
            set_bit Write_Buf_W350(0), 12 'resultat mauvais
            reset_bit Write_Buf_W350(0), 7 'acquis message
            reset_bit Write_Buf_W350(0), 3 'mesure P & H
            Call ecrit_tsx(350, Write_Buf_W350(), 1)
End Sub



Private Sub Command1_Click()

End Sub

Private Sub CMD_ok_Click()
    CMD_bon_Click
End Sub

Private Sub CMD_P_H_Click()

LBL_lim_p.Caption = Format$(portee_mini_tor) & " < Portee (mm) < " & Format$(portee_maxi_tor)
If hyst_option_tor = 0 Then
    LBL_lim_h.Caption = Format$(hysteresis_mini_tor) & " < Hyst.(mm) < " & Format$(hysteresis_maxi_tor)
Else
    LBL_lim_h.Caption = Format$(hysteresis_mini_tor) & " < Hyst.(%) < " & Format$(hysteresis_maxi_tor)
End If

Write_Buf_W350(0) = W350(0)
set_bit W350(0), 3 'mesurer P & H
set_bit Write_Buf_W350(0), 3 'mesurer P & H
Call ecrit_tsx(350, Write_Buf_W350(), 1)
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_F1 Then
        Call CMD_bon_Click
    End If
    
    If KeyCode = KEY_F5 Or KeyCode = vbKeyDivide Then
        Call CMD_mauvais_Click
    End If
    
    If KeyCode = KEY_F2 And fra_mes_p_h.Enabled = True And fra_mes_p_h.Visible = True Then
        Call CMD_P_H_Click
    End If

End Sub

Private Sub Form_Load()
    Message_affiché = True
End Sub


Private Sub LBL_messsage_a_afficher_Click()

End Sub


Private Sub Form_Paint()

'met la fenetre au premier plan, et on ne peut plus l'effacer !
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW

End Sub


Private Sub Form_Unload(Cancel As Integer)

Message_a_afficher = ""
fra_mes_p_h.Enabled = False
fra_mes_p_h.Visible = False
mesure_p_h_manu = False
Message_affiché = False
End Sub


