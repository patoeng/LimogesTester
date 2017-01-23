VERSION 5.00
Begin VB.Form FRM_choix_température 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Temperature Choice"
   ClientHeight    =   1740
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bouton_Annulé 
      Caption         =   "ESC - Cancel"
      Height          =   492
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   1572
   End
   Begin VB.CommandButton Bouton_tmp_Haute 
      Caption         =   "F2 - Hot"
      Height          =   492
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1572
   End
   Begin VB.CommandButton Bouton_tmp_Basse 
      Caption         =   "F1 - COLD"
      Height          =   492
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1572
   End
End
Attribute VB_Name = "FRM_choix_température"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bouton_Annulé_Click()
    Commande_mode_de_marche_OK = False
    Commande_mode_de_marche_SET = True
    Unload Me
End Sub

Private Sub Bouton_tmp_Basse_Click()
Dim ComErreure As Long

Dim Mode_chargement As Long
Dim Mode_dielectric As Long
Dim Mode_final As Long
Dim Cycle As Integer
   
Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i

    Mode_chargement = BIT8 'toujours actif
    Mode_dielectric = 0 'toujours inatif
    Mode_final = API_TEMPERATURE_BASSE + BIT8
    
    Copy_Tab(10) = Mode_chargement
    Copy_Tab(11) = Mode_dielectric
    Copy_Tab(12) = Mode_final
    Dim Com_erreure As Long
    Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 3, 10)
    If Com_erreure <= 0 Then
        If MsgBox(ERREUR_ECRITURE_MODE_DE_MARCHE, vbCritical + vbRetryCancel) = vbCancel Then
            If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                End
            End If
        End If
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Tab_Status_banc(i) = Copy_Tab(i)
        Next i
        Exit Sub
    End If
    
    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Tab_Status_banc(i) = Copy_Tab(i)
    Next i
    
    If Générer_sequence_API_contrôle_pos_final(Code_produit_badge, CYCLE_TEMPERATURE_BASSE) < 0 Then
        'Défaut de com ou de base de donnée
        Bouton_Annulé_Click
        Exit Sub
    End If
    
    type_controle = TEMPERATURE_BASSE
    Commande_mode_de_marche_OK = True
    Commande_mode_de_marche_SET = True
    Unload Me

End Sub

Private Sub Bouton_tmp_Haute_Click()
Dim ComErreure As Long

Dim Mode_chargement As Long
Dim Mode_dielectric As Long
Dim Mode_final As Long
Dim Cycle As Integer

Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i

    Mode_chargement = BIT8 'toujours actif
    Mode_dielectric = 0 'toujours inactif
    Mode_final = API_TEMPERATURE_HAUTE + BIT8
    

    Copy_Tab(10) = Mode_chargement
    Copy_Tab(11) = Mode_dielectric
    Copy_Tab(12) = Mode_final
    Dim Com_erreure As Long
    Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 3, 10)
    If Com_erreure <= 0 Then
        If MsgBox(ERREUR_ECRITURE_MODE_DE_MARCHE, vbCritical + vbRetryCancel) = vbCancel Then
            If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                End
            End If
        End If
        
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Tab_Status_banc(i) = Copy_Tab(i)
        Next i
    
        Exit Sub
    End If
    
    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Tab_Status_banc(i) = Copy_Tab(i)
    Next i
    
    If Générer_sequence_API_contrôle_pos_final(Code_produit_badge, CYCLE_TEMPERATURE_HAUTE) < 0 Then
        'Défaut de com ou de base de donnée
        Bouton_Annulé_Click
        Exit Sub
    End If
    
    Commande_mode_de_marche_OK = True
    type_controle = TEMPERATURE_HAUTE
    Commande_mode_de_marche_SET = True
    Unload Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Bouton_tmp_Basse_Click
    End If
    
    If KeyCode = vbKeyF2 Then
        Bouton_tmp_Haute_Click
    End If
    
    If KeyCode = vbKeyEscape Then
        Bouton_Annulé_Click
    End If
End Sub

Private Sub Form_Load()

'met la fenetre au premier plan, et on ne peut plus l'effacer !
'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW

End Sub
