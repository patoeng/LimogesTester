VERSION 5.00
Begin VB.Form FRM_dial_tsx 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Dialogue avec le TSX"
   ClientHeight    =   1104
   ClientLeft      =   2616
   ClientTop       =   2436
   ClientWidth     =   3516
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1104
   ScaleWidth      =   3516
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer TIM_dial_TSX 
      Interval        =   50
      Left            =   1560
      Top             =   240
   End
End
Attribute VB_Name = "FRM_dial_tsx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tim_dial_TSX_Timer()

Dim Com_erreure As Long

Dim i As Integer
Dim Copy_Tab(20) As Long

For i = LBound(Copy_Tab) To UBound(Copy_Tab)
    Copy_Tab(i) = Tab_Status_banc(i)
Next i


'On Error GoTo 500

'Vérification qu'un produit a été sélectionné
If FRM_startup!Lbl_code_lu.Caption <> "" Then
    symbole_existe = True
Else
    symbole_existe = False
End If

'maj des variables etat de l'IHM
Maj_etat_banc

Select Case Contexte$

Case "AFFICHAGE"
    If aff_param = 1 Then
        aff_param = 0
        FRM_PARAM.Show
    End If

Case "STARTUP"
    'FRM_startup.Show
    If nouvelle_barquette = False Then
        If (Copy_Tab(6) And BIT0) <> 0 Then
            FRM_startup!Lbl_code_lu.Caption = ""
            FRM_startup.ImageEmpreinte.Picture = FRM_startup.ImageListEmpreinte.ListImages(1).Picture
        End If
        
        Call Maj_etat_banc
        
        maj_boite_dial_qualification
        
    End If
    
    If nouvelle_barquette = True Then
        Dim Tab_parametre() As Long
        
        Com_erreure = Lecture_paramétre_produit(Tab_parametre()) 'lecture de la table MW420 provenant directement du badge
        If Com_erreure <= 0 Then
            Exit Sub
        End If
        prepa_params Tab_parametre()
        FRM_startup!Lbl_code_lu.Caption = symb_com
        
        If empreinte_miniature Then
            FRM_startup.ImageEmpreinte.Picture = FRM_startup.ImageListEmpreinte.ListImages(3).Picture
        Else
            FRM_startup.ImageEmpreinte.Picture = FRM_startup.ImageListEmpreinte.ListImages(2).Picture
        End If
        
        nouvelle_barquette = False
        reset_bit Copy_Tab(0), 2
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 1)
        If Com_erreure <= 0 Then
        
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            
            Exit Sub
        End If
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
        Next i
    End If
    Call message_defaut_claquage
    Call message_defaut_general
    Call message_defaut_claquage

Case "CONTROLE"
    'maj des variables etat de l'IHM
    'Maj_etat_banc
        
    If lance_controle = 1 Then
        reset_bit Copy_Tab(0), 1
        set_bit Copy_Tab(0), 5 ' autorisation du cycle par IHM
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 1)
        If Com_erreure <= 0 Then
            
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            
            Exit Sub
        End If
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Tab_Status_banc(i) = Copy_Tab(i)
        Next i
        controle_affiche = True
        Unload FRM_controle
        FRM_controle.Show
        lance_controle = 0
        Call appel_controle
        Exit Sub
    End If
    
    If out_controle = 1 Then
        out_degrade = False
        affichage_dégradé = False
        controle_affiche = False
        out_controle = 0
        lance_controle = 0
        Contexte$ = "STARTUP"
        FRM_controle!LBL_symb_com.Caption = "Au revoir...."
        reset_bit Copy_Tab(0), 0
        reset_bit Copy_Tab(0), 1
        reset_bit Copy_Tab(0), 2
        reset_bit Copy_Tab(0), 3
        reset_bit Copy_Tab(0), 4
        reset_bit Copy_Tab(0), 5
        reset_bit Copy_Tab(0), 7
        reset_bit Copy_Tab(0), 8
        reset_bit Copy_Tab(0), 14
        Copy_Tab(3) = 0
        Copy_Tab(4) = 0
        Copy_Tab(5) = 0
        Copy_Tab(6) = 0
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 7)
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Tab_Status_banc(i) = Copy_Tab(i)
        Next i
        'If Com_Erreure <= 0 Then
        '    Exit Sub
        'End If
        Unload FRM_controle
        Unload FRM_FIN_CONTROLE
        Exit Sub
    End If
    
    'Gestion des défauts
    Call message_defaut_general
    'Call message_defaut_chargement 'inséré dans CONTROLE maj_controle
    Call message_defaut_claquage
    Call message_defaut_ctrl_final

    If controle_affiche = True Then
        Call maj_controle
    End If
    
Case "AUTOTEST"

    Call message_defaut_general
    Call message_defaut_claquage
    Call message_defaut_ctrl_final
    
    'Affichage etape autotest
    FRM_AUTOTEST!Lbl_pas.Caption = Str$(Tab_Status_banc(69))
    Call maj_boite_dial_operateur_autotest
    
    If out_autotest = 1 Then
        out_autotest = 0
        reset_bit Copy_Tab(0), 0
        reset_bit Copy_Tab(0), 1
        reset_bit Copy_Tab(0), 2
        reset_bit Copy_Tab(0), 3
        reset_bit Copy_Tab(0), 4
        reset_bit Copy_Tab(0), 5
        reset_bit Copy_Tab(0), 7
        reset_bit Copy_Tab(0), 8
        reset_bit Copy_Tab(0), 14
        Copy_Tab(3) = 0
        Copy_Tab(4) = 0
        Copy_Tab(5) = 0
        Copy_Tab(6) = 0
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 7)
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Tab_Status_banc(i) = Copy_Tab(i)
        Next i
        'If Com_Erreure <= 0 Then
        '    Exit Sub
        'End If
        Contexte$ = "STARTUP"
        Unload FRM_AUTOTEST
        Exit Sub
    End If

    If lance_autotest = 1 Then
        lance_autotest = 0
        autoris = 0
        Load FRM_AUTOTEST
        FRM_AUTOTEST.Show
        Exit Sub
    End If
    
    If autoris = 1 Then
        'maj des variables etat de l'IHM
            ' Maj_etat_banc
               
        'Si fin d'autotest on sort
        If (Tab_Status_banc(0) And BIT8) = 0 Then
                FRM_AUTOTEST!CMB_Validation_autotest.Visible = False
                FRM_AUTOTEST!CMB_Annulation_autotest.Visible = False
                FRM_AUTOTEST!LBL_AUTOTEST_EN_COURS.Visible = False
                FRM_AUTOTEST!LBL_Autotest_termine.Visible = True
                FRM_AUTOTEST!CMB_fin_autotest.Visible = True
        End If
        
         'Affiche les résultats en fin de cycle
        If (Tab_Status_banc(0) And BIT3) <> 0 Then
            Call ecrit_resultat_autotest
            reset_bit Copy_Tab(0), 3
            Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 6)
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            If Com_erreure <= 0 Then
                Exit Sub
            End If
        End If
    End If
    
End Select

    Exit Sub
500 MsgBox "Erreur " & Str$(Err) & " module Timer TIM_dial_tsx. Signalez immediatement cette erreur. Merci."
    Resume Next


End Sub


