Attribute VB_Name = "mdl_CONTROLE"
Option Explicit
DefInt A-Z

' Developpe par : A2SI - Fleac
' Voir SUIVI.VER pour les versions

Declare Function ShowCursor Lib "USER" (ByVal bShow As Integer) As Integer
Declare Sub setcursorpos Lib "USER" (ByVal x As Integer, ByVal y As Integer)
Declare Sub MessageBeep Lib "USER" (ByVal x As Integer)
Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer
Declare Function GetProfileString Lib "Kernel" Alias "GetProfilestring" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function GetProfileInt Lib "Kernel" Alias "GetProfileint" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
'Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Déclaration de la structure Message et
'Variable relative a la gestion des messages
Public Type Type_Message
    ID_Message As Long ' idantification du message
    Etat As Boolean 'Bon ou mauvais
    Valeur As Double 'Valeur messurée
    ID_fonction As Long ' idantification de la fonction source du message
End Type

Public Type Info_produit
    Code_comercial As String
    Code_lot As String
    Num_produit As Byte
    Mun_emprente As Byte
End Type

Public Enum Type_produit
    Prod_TOR_NC = 1
    Prod_TOR_NO = 2
    Prod_Ana_IS = 3
    Prod_Ana_IS_IC = 4
    Prod_TOR_Auto_adaptable_NC = 5
    Prod_TOR_Auto_adaptable_NO = 6
End Enum

' Variable pour traitement des message
Global Lob_un_defaut As Boolean
Global Log_un_defaut_majeur As Boolean

Global Etiquette_un_defaut As Boolean
Global Etiquette_un_defaut_majeur As Boolean

Global Imprimer_etiquette As Boolean

Const NON_TEST = 0
Const BON = 1
Const MAUVAIS = -1

' <<<<- Variables liees au contrôle
Global del_P2, del_P1
Global log_file     ' =1 si archivage
Global produit_en_test
Global symbole_existe, controle_affiche As Boolean
Global ac_temp As Boolean
Global ac_final As Boolean
Global ac As Boolean
Global alim(0 To 1) As Integer
Global ctrl_simplifie As Boolean
Global ctrl_3_4_fils As Boolean
Global type_commutation As Integer

Global Nombre_produit_à_tester As Integer

'parametres controle final
'-> SODATEC 9/5/1
Global si_ctrl_analogique As Boolean
Global si_ctrl_analogique_IS As Boolean
Global si_ctrl_analogique_IS_IC As Boolean
Global si_ctrl_tor As Boolean
Global si_ctrl_auto_adaptatif As Boolean
Global courant_charge As Double
Global courant_charge_T As Double
Global empreinte_miniature As Boolean
'<- SODATECGlobal
Global portee_mini As Double
Global portee_maxi As Double
Global portee_nominal As Double
Global portee_mini_auto As Double
Global hysteresis_mini As Double
Global hysteresis_maxi As Double
Global ic_act_min_u1(2) As Double
Global ic_act_max_u1(2) As Double
Global ic_act_min_u2(2) As Double
Global ic_act_max_u2(2) As Double
Global ic_repos_min_u1(2) As Double
Global ic_repos_max_u1(2) As Double
Global ic_repos_min_u2(2) As Double
Global ic_repos_max_u2(2) As Double
Global if_min_u1(2) As Double
Global if_max_u1(2) As Double
Global if_min_u2(2) As Double
Global if_max_u2(2) As Double
Global ud_min_u1(2) As Double
Global ud_max_u1(2) As Double
Global ud_min_u2(2) As Double
Global ud_max_u2(2) As Double
Global u_alim(0 To 1) As Integer
Global u_alim_F(0 To 1) As Integer
Global protégé_cc_final As Boolean
Global ctrl_simplifie_final
Global avec_ctrl_final

' affectation des parametres controle température
Global portee_mini_T As Double
Global portee_maxi_T As Double
Global hysteresis_mini_T As Double
Global hysteresis_maxi_T As Double
Global portee_nominal_T As Double
Global portee_mini_auto_T As Double

Global u_alim_T(0 To 1) As Integer
Global protégé_cc_T As Boolean
Global ctrl_simplifie_T As Boolean

'parametres controle diélectrique
Global Ufinal As Integer
Global I_fuite As Double
Global avec_claquage

'=======DEBUT MODIF BEEA 30/11/2007================
'parametres controle station diélectrique
Global ModeMarcheDielec As Integer 'O= marche normale, 1=demande produits test, 2=dégradé sur empreinte 1, 3= dégardé sur empreinte 2, 4=dégradé sur les 2 empreintes, 5=poste bloqué
Global Derniere_execution_Test_Claqu_OK As Date
Global Prochaine_Exec_Test_Claqu As Date
Global Derniere_Modif_Param_Test_Claqu As Date
Global Modifie_Par_Test_Claqu As String
Global Frequence_Test_Claqu As Integer
Global Heure_Test_Claqu As Date
Global ControleStationClaquageEnCours As Boolean
Global TestStationClaquage1OK As Boolean
Global TestStationClaquage2OK As Boolean
Global FinTestClaquageAffiche As Boolean
Global FinControleAffiche As Boolean
Global reponse As String
Global Enr_claqu_en_cours As Boolean
'=======FIN MODIF BEEA 30/11/2007==================

'Autotest
Global No_de_defaut, Valeur
Global ualim As Integer

' Drapeau bon/mauvais
Global produit1_bon As Boolean, produit2_bon As Boolean
' ...
Global type_controle    ' 0=dégradé 1=temperature 2=final
Global Code_produit_badge     ' Vien du badge %mw468 ou %mw524
                        ' identifie les cycles de contrôle à télécharger dans l'API

' Boîtes aux lettres pour la communication
Global com_err_tsx, previous_com_err_tsx, etape_avant_defaut, previous_def_final, previous_def_chargement, previous_def_claquage
Global def_final As Long, def_claquage As Long, def_chargement As Long
      '...
Global Commande_mode_de_marche_OK As Boolean
Global Commande_mode_de_marche_SET As Boolean
Global Contexte$
Global out_controle
Global lance_controle
Global runnig_controle
Global out_autotest
Global lance_autotest
Global Calibre_Multi$
Global aff_param
Global nouvelle_barquette As Boolean
Global cycle_chargement As Long 'W414
Global cycle_claquage As Long 'W415
Global cycle_final As Long 'W416
Global cycle_chargement_precedent As Long 'Mémorise l'état precedent de W414
Global cycle_claquage_precedent As Long 'Mémorise l'état precedent de W415
Global cycle_final_precedent As Long 'Mémorise l'état precedent de W416
Global demande_station As Boolean
Global npn As Boolean
Global appel_maintenance As Boolean
Global mes_def_claq As String
Global mes_def_final As String
Global protege_cc As Boolean
Global forçage_etiquette
Global sans_etiquette


Global Tab_Status_bancx7 As Boolean

'Variable affichage message
Global ARU_Affiché As Boolean
Global valid2, valid1, relance1
Global aff_def_char
Global Defaut

' <<<<- Fin variables liees au contrôle

' <<<<- Variables liees à la forme ATTENTION
Global go_config
' <<<<- Fin variables liees à la forme ATTENTION

Global symb_com As String, Ref_indus As String, Num_lot As String, nom_fichier As String

'Variable pour autotest
Global autoris As Integer
Global resultat_autotest As Integer

'Variable pour controle diélectrique
Global metallique As Boolean
Global fil_masse As Boolean
Global P1_bon_sefelec As Boolean
Global P2_bon_sefelec As Boolean

'Variable pour le mode dégradé
Global out_degrade As Boolean
Global pas_a_pas_chargement As Boolean
Global pas_a_pas_claquage As Boolean
Global pas_a_pas_final As Boolean
Global pas_a_pas_premier_ctl As Boolean
Global auto_chargement As Boolean
Global auto_claquage As Boolean
Global auto_final As Boolean
Global auto_premier_ctl As Boolean
Global affichage_dégradé As Boolean
Global aff_def

'Variable femin aced au fichier
Global file_l As String
Global file_save As String
Global file_log$
Global file_hist_def$
Global path_log$

'Constantes
Global Const BIT0 = 1
Global Const BIT1 = 2
Global Const BIT2 = 4
Global Const BIT3 = 8
Global Const BIT4 = 16
Global Const BIT5 = 32
Global Const BIT6 = 64
Global Const BIT7 = 128
Global Const BIT8 = 256
Global Const BIT9 = 512
Global Const BIT10 = 1024
Global Const BIT11 = 2048
Global Const BIT12 = 4096
Global Const BIT13 = 8192
Global Const BIT14 = 16384
Global Const BIT15 = 32768

Global Const POIDS0# = 1
Global Const POIDS8# = 256
Global Const POIDS16# = 65536
Global Const POIDS24# = 16777216
Global Const POIDS32# = 4294967296#

Global Const Y1dedans = 0
Global Const Y1dehors = 12000

Global Const Coeff_Ualim = 5
Global Const Coeff_Icharge = 5

Global Const Décalage_produit_2 = 32





Sub Ajout_nouvelle_barquette_log()

Dim Text As String

    If DataEnvironment1.rsTable_Log_Banc.State = adStateClosed Then
        DataEnvironment1.rsTable_Log_Banc.Open
    End If
    
    Select Case type_controle
        Case DEGRADE
            'Text = "Dégradé"
            Text = "Degrade"
        Case TEMPERATURE
            'Text = "Température"
            Text = "Temperature"
        Case TEMPERATURE_HAUTE
            'Text = "Température haute"
            Text = "Temperature HOT"
        Case TEMPERATURE_BASSE
            'Text = "Température basse"
            Text = "Temperature COLD"
        Case CONTROLE_FINAL
            Text = "Final"
        Case AUTO_TEST
            Text = "Auto test"
        Case PREMIER_CONTROLE
            'Text = "Premier contrôle"
            Text = "Test after laser"
    End Select
    
    If Not DataEnvironment1.rsTable_Log_Banc.BOF And Not DataEnvironment1.rsTable_Log_Banc.EOF Then
        If DataEnvironment1.rsTable_Log_Banc.Status = adRecConcurrencyViolation Then
            DataEnvironment1.rsTable_Log_Banc.CancelUpdate
        End If
    End If
    Dim wxc As Long
    'wxc = DataEnvironment1.rsTable_Log_Banc.Status
    'DataEnvironment1.rsTable_Log_Banc.Open
    'DataEnvironment1.rsTable_Log_Banc.MoveFirst
    DataEnvironment1.rsTable_Log_Banc.AddNew
    DataEnvironment1.rsTable_Log_Banc.Fields("Date_debut").Value = Now
    DataEnvironment1.rsTable_Log_Banc.Fields("Num_lot").Value = Num_lot
    DataEnvironment1.rsTable_Log_Banc.Fields("Num_com").Value = symb_com
    DataEnvironment1.rsTable_Log_Banc.Fields("Mode_fonctionnement").Value = Text
    
    'DataEnvironment1.rsTable_Log_Banc.Update
    
End Sub

Sub Fin_barquette_log(Optional Nombre_prob_bon As Long = 0)

    If DataEnvironment1.rsTable_Log_Banc.State = adStateClosed Then
        Exit Sub
    End If
    
    If DataEnvironment1.rsTable_Log_Banc.Status = adRecConcurrencyViolation Then
        DataEnvironment1.rsTable_Log_Banc.CancelUpdate
    End If
    
    'Mis à jour de toute la table
    'DataEnvironment1.rsTable_Log_Banc.Filter = ("")
    'DataEnvironment1.rsTable_Log_Banc.Requery
    'DataEnvironment1.rsTable_Log_Banc.MoveLast
    
    DataEnvironment1.rsTable_Log_Banc.Fields("Date_fin").Value = Now
    DataEnvironment1.rsTable_Log_Banc.Fields("Nombre_produit_bon").Value = CLng(Nombre_prob_bon)
    DataEnvironment1.rsTable_Log_Banc.Update
    
End Sub


'Décode la trame message
Function Decode_Trame_message(Trame_message() As Long, Message As Type_Message) As Boolean

Dim Val As Long
Dim val_f As Double
Dim Gain As Double

    Message.ID_Message = Trame_message(LBound(Trame_message) + 1)
    Message.ID_fonction = Trame_message(LBound(Trame_message))
    Message.Etat = BIT0 And Trame_message(LBound(Trame_message) + 2)
    
    Val = (Trame_message(LBound(Trame_message) + 3) And &HFFFF) + ((Trame_message(LBound(Trame_message) + 4) And &H7FFF) * 65536)
    If (Trame_message(LBound(Trame_message) + 4) And &H8000) <> 0 Then
    Val = 0 + Val
    End If
    
    If DataEnvironment1.rsTable_Message.State = adStateClosed Then
        DataEnvironment1.rsTable_Message.Open
    End If
    DataEnvironment1.rsTable_Message.Filter = "ID_Message = '" & Message.ID_Message & "'"
    
    If DataEnvironment1.rsTable_Message.RecordCount > 0 Then
        'Idantification message trouvé dans la basse de donnée message
        
        DataEnvironment1.rsTable_Message.MoveFirst
        
        If DataEnvironment1.rsTable_Message.Fields("Valeur_utilise").Value = True Then
            val_f = CDbl(Val)
                    
            Gain = DataEnvironment1.rsTable_Message.Fields("Gain_Valeur").Value
            If Gain = 0 Then Gain = 1 'Evite division par zero
            val_f = -DataEnvironment1.rsTable_Message.Fields("Decalage_Valeur").Value + (val_f / Gain)
        
        End If
        
        Message.Valeur = val_f
    
        Decode_Trame_message = True
    Else
        'Idantification message non trouvé dans la basse de donnée message
        Decode_Trame_message = False
        Message.Valeur = 0
    End If
End Function

'Ajoute un nouvelle enregistrement dans la base de donnée log
'Initialise les champs telle que la reférence commercial, la série, .....
Sub Ajoute_produit_log(Info As Info_produit)
    
    If DataEnvironment1.rsTable_Log.State = adStateClosed Then
        DataEnvironment1.rsTable_Log.Open
    End If
    
    'DataEnvironment1.rsTable_Log.CancelUpdate
    
    DataEnvironment1.rsTable_Log.AddNew
    
    DataEnvironment1.rsTable_Log.Fields("Num_prod") = Info.Num_produit
    DataEnvironment1.rsTable_Log.Fields("Num_emprunte") = Info.Mun_emprente
    DataEnvironment1.rsTable_Log.Fields("Num_lot") = Info.Code_lot
    DataEnvironment1.rsTable_Log.Fields("Num_com") = Info.Code_comercial
    DataEnvironment1.rsTable_Log.Fields("Date") = Now
    DataEnvironment1.rsTable_Log.Fields("Non_banc").Value = Non_banc
    
    If Len(Non_banc) > DataEnvironment1.rsTable_Log.Fields("Non_banc").DefinedSize Then
    'Retaille la chaine Non_banc  à la dimention du champ de la DB
    Non_banc = Left(Non_banc, DataEnvironment1.rsTable_Log.Fields("Non_banc").DefinedSize)
End If
DataEnvironment1.rsTable_Log.Fields("Non_banc").Value = Non_banc

    
    Lob_un_defaut = False
    Log_un_defaut_majeur = False
    Imprimer_etiquette = False
    
End Sub
'Rempli les champs de la base de donnée log
Sub Ajoute_donnee_log(Message As Type_Message, Afficheur As AfficheProj1.Affiche)


If DataEnvironment1.rsTable_Message.RecordCount > 0 Then
    If DataEnvironment1.rsTable_Message.Fields("Donnée_log").Value = True Then
         If DataEnvironment1.rsTable_Message.Fields("Champ_B_D_produit").Value <> "" Then
            If DataEnvironment1.rsTable_Message.Fields("Valeur_utilise").Value = True Then
                'on met la valeur dans le champs de la base de donnée log
                DataEnvironment1.rsTable_Log.Fields(DataEnvironment1.rsTable_Message.Fields("Champ_B_D_produit").Value).Value = Message.Valeur
            Else
                'on met les constantes Bon ou Mauvais dans le champs de la base de donnée log
                If Message.Etat = True Then
                    DataEnvironment1.rsTable_Log.Fields(DataEnvironment1.rsTable_Message.Fields("Champ_B_D_produit").Value).Value = BON
                Else
                    DataEnvironment1.rsTable_Log.Fields(DataEnvironment1.rsTable_Message.Fields("Champ_B_D_produit").Value).Value = MAUVAIS
                End If
            End If
        End If
        
        
        ' Affiche les messages sur l'écran
        Affiche_Message Message.Etat, Message.Valeur, Message.ID_Message, Afficheur
        
        If Message.Etat = False Then
            
            If DataEnvironment1.rsTable_Message.Fields("Défaut_est_mageur").Value = True Then
                'Defaut majeur
                If Not Log_un_defaut_majeur Then
                    DataEnvironment1.rsTable_Log.Fields("Num_def").Value = DataEnvironment1.rsTable_Message.Fields("Code_Defaut").Value
                    DataEnvironment1.rsTable_Log.Fields("Num_phase").Value = DataEnvironment1.rsTable_Message.Fields("Phase_Defaut").Value
                    Log_un_defaut_majeur = True
                    Lob_un_defaut = True
                    If DataEnvironment1.rsTable_Message.Fields("Imprimer_etiquette").Value = True Then
                        'Positionne l'indicateur pour imprimer une étiquette
                        Imprimer_etiquette = True
                        DataEnvironment1.rsTable_Log.Fields("Num_def_etiquette").Value = DataEnvironment1.rsTable_Message.Fields("Code_Etique").Value
                    End If
                    
                    'maj labelle code défaut etiquette
                     DataEnvironment1.rsTable_Code_defaut_etiquette.Filter = "Code = '" & DataEnvironment1.rsTable_Message.Fields("Code_Etique").Value & "'"
                     If DataEnvironment1.rsTable_Code_defaut_etiquette.State = adStateClosed Then
                         DataEnvironment1.rsTable_Code_defaut_etiquette.Open
                     End If
                     If DataEnvironment1.rsTable_Code_defaut_etiquette.RecordCount > 0 Then
                         DataEnvironment1.rsTable_Log.Fields("Labelle_def_etiquette").Value = DataEnvironment1.rsTable_Code_defaut_etiquette.Fields("Labelle").Value
                     Else
                         DataEnvironment1.rsTable_Log.Fields("Labelle_def_etiquette").Value = DataEnvironment1.rsTable_Message.Fields("Code_Etique").Value
                     End If
                
                 
                     'maj labelle phase
                     DataEnvironment1.rsTable_Code_phase.Filter = "Code = '" & DataEnvironment1.rsTable_Message.Fields("Phase_Defaut").Value & "'"
                     If DataEnvironment1.rsTable_Code_phase.State = adStateClosed Then
                         DataEnvironment1.rsTable_Code_phase.Open
                     End If
                     If DataEnvironment1.rsTable_Code_phase.RecordCount > 0 Then
                         DataEnvironment1.rsTable_Log.Fields("Labelle_phase").Value = DataEnvironment1.rsTable_Code_phase.Fields("Labelle").Value
                     Else
                         DataEnvironment1.rsTable_Log.Fields("Labelle_phase").Value = DataEnvironment1.rsTable_Message.Fields("Phase_Defaut").Value
                     End If
                    
                End If
            Else
                'Defaut non majeur
                If Not Lob_un_defaut And Not Log_un_defaut_majeur Then
                    DataEnvironment1.rsTable_Log.Fields("Num_def").Value = DataEnvironment1.rsTable_Message.Fields("Code_Defaut").Value
                    DataEnvironment1.rsTable_Log.Fields("Num_phase").Value = DataEnvironment1.rsTable_Message.Fields("Phase_Defaut").Value
                    Lob_un_defaut = True
                    
                    If DataEnvironment1.rsTable_Message.Fields("Imprimer_etiquette").Value = True Then
                        'Positionne l'indicateur pour imprimer une étiquette
                        Imprimer_etiquette = True
                        DataEnvironment1.rsTable_Log.Fields("Num_def_etiquette").Value = DataEnvironment1.rsTable_Message.Fields("Code_Etique").Value
                    End If
                    
                    
                     'maj labelle code défaut etiquette
                     DataEnvironment1.rsTable_Code_defaut_etiquette.Filter = "Code = '" & DataEnvironment1.rsTable_Message.Fields("Code_Etique").Value & "'"
                     If DataEnvironment1.rsTable_Code_defaut_etiquette.State = adStateClosed Then
                         DataEnvironment1.rsTable_Code_defaut_etiquette.Open
                     End If
                     If DataEnvironment1.rsTable_Code_defaut_etiquette.RecordCount > 0 Then
                         DataEnvironment1.rsTable_Log.Fields("Labelle_def_etiquette").Value = DataEnvironment1.rsTable_Code_defaut_etiquette.Fields("Labelle").Value
                     Else
                         DataEnvironment1.rsTable_Log.Fields("Labelle_def_etiquette").Value = DataEnvironment1.rsTable_Message.Fields("Code_Etique").Value
                     End If
                
                 
                     'maj labelle phase
                     DataEnvironment1.rsTable_Code_phase.Filter = "Code = '" & DataEnvironment1.rsTable_Message.Fields("Phase_Defaut").Value & "'"
                     If DataEnvironment1.rsTable_Code_phase.State = adStateClosed Then
                         DataEnvironment1.rsTable_Code_phase.Open
                     End If
                     If DataEnvironment1.rsTable_Code_phase.RecordCount > 0 Then
                         DataEnvironment1.rsTable_Log.Fields("Labelle_phase").Value = DataEnvironment1.rsTable_Code_phase.Fields("Labelle").Value
                     Else
                         DataEnvironment1.rsTable_Log.Fields("Labelle_phase").Value = DataEnvironment1.rsTable_Message.Fields("Phase_Defaut").Value
                     End If
                    
                End If
            End If 'Défaut majeur ou non majeur
            
            
            
        End If 'Etat message en defaut
    End If
End If
    
End Sub

'Valide l'enregistrement de la basse de donnée log
Function Valide_donnee_log(Optional Init_num_enregistrement As Boolean = False) As Long

    Static Num_enregistrement_min As Long
    
    If Init_num_enregistrement Then
        Num_enregistrement_min = 0
        Valide_donnee_log = 0
        Exit Function
    End If
    
    DataEnvironment1.rsTable_Log.Fields("Etat_produit").Value = Not (Lob_un_defaut Or Log_un_defaut_majeur)
    
    'Validation des modification de l'enregistrement en cours
    DataEnvironment1.rsTable_Log.Update
    
    'Mis à jour de toute la table
    'DataEnvironment1.rsTable_Log.Filter = ("")
    'DataEnvironment1.rsTable_Log.Requery
    'DataEnvironment1.rsTable_Log.MoveLast
    
    If Num_enregistrement_min = 0 Then
        
        'If DataEnvironment1.rsrequette_update.State <> adStateClosed Then
        '    DataEnvironment1.rsrequette_update.Close 'ferme le requette avant de l'ouvir
        'End If
        If DataEnvironment1.rsEnregistrement_max.State <> adStateClosed Then
            DataEnvironment1.rsEnregistrement_max.Close 'ferme le requette avant de l'ouvir
        End If
        
        'DataEnvironment1.requette_update (Num_enregistrement_min)
        'If DataEnvironment1.rsrequette_update.RecordCount > 0 Then
        If DataEnvironment1.rsTable_Log.RecordCount > 0 Then
            ' On cherche le N° enregistrement
        '    DataEnvironment1.rsrequette_update.MoveLast
        '    Valide_donnee_log = DataEnvironment1.rsrequette_update.Fields("N°_enregistrement").Value
            DataEnvironment1.rsEnregistrement_max.Open
            Valide_donnee_log = DataEnvironment1.rsEnregistrement_max.Fields("Nb_MAX_Log").Value
            Num_enregistrement_min = Valide_donnee_log
            If DataEnvironment1.rsEnregistrement_max.State <> adStateClosed Then
                DataEnvironment1.rsEnregistrement_max.Close 'ferme le requette avant de l'ouvir
            End If
        End If
        
        'Num_enregistrement_min = DataEnvironment1.rsTable_Log.Fields("N°_enregistrement").Value
        'Valide_donnee_log = Num_enregistrement_min
    Else
        Num_enregistrement_min = Num_enregistrement_min + 1
        Valide_donnee_log = Num_enregistrement_min
    End If
    
    'Valide_donnee_log = DataEnvironment1.rsTable_Log.Fields("N°_enregistrement").Value
        
End Function

Sub maj_controle()
Dim i As Integer

Dim Maj_HIM_vers_API_etat_banc As Boolean ' Indique si le faut écrir la table etat_banc a l'API
Dim Result_com As Long ' Résultat écriture table mots état API

Dim Copy_Tab(20) As Long

For i = LBound(Copy_Tab) To UBound(Copy_Tab)
    Copy_Tab(i) = Tab_Status_banc(i)
Next i

'On Error GoTo 400
        
    'Met à jour l'ecran en fonction des demandes de l'automate
    Call maj_ecran

    'Demande impression etiquette defaut led
        If del_P2 = True Then
            del_P2 = False
            
            set_bit Copy_Tab(6), 10 ' Produit 2 mauvais par PC
            Maj_HIM_vers_API_etat_banc = True
            
            'Gesttion du défaut et Impression etiquette défaut led
            def_chargement = MESSAGE_DEF_LED_PRODUIT_2
            Call message_defaut_chargement
            
        End If
            
       If del_P1 = True Then
            del_P1 = False
            
            reset_bit Copy_Tab(6), 9 ' Produit 1 mauvais par PC
            Maj_HIM_vers_API_etat_banc = True
            
            'Gesttion du défaut et Impression etiquette défaut led
            def_chargement = MESSAGE_DEF_LED_PRODUIT_1
            Call message_defaut_chargement
            
        End If
       
        'On sort du controle sur changement d'état AUTO ou PAS à PAS
        If (Tab_Status_banc(0) And BIT5) = 0 Then 'cycle interdit
                out_controle = 1
        End If
        
        'Affichage manque Bac produits à tester
        If (Tab_Status_banc(6) And BIT0) <> 0 Then 'Absence bac produit à tester
            'MsgBox "Remetre le bac de produit à tester : le conrole n'est pas fini"
            MsgBox "Letakkan kembali tray product yang akan ditest : Testing belum selesai"
        End If
        
        'Affiche les résultats en fin de cycle
        If (Tab_Status_banc(0) And BIT3) <> 0 Then 'Demande affichage résultat
            maj_etat_produit
            Analyse_message
            reset_bit Copy_Tab(0), 3
            Maj_HIM_vers_API_etat_banc = True
        End If
        
        'Efface les résultats en début du départ cycle chargement
         If (Tab_Status_banc(0) And BIT4) <> 0 Then 'Demmande effacer résultat
            Call RAZ_ecran
            reset_bit Copy_Tab(0), 4
            Maj_HIM_vers_API_etat_banc = True
        End If
        
         'Affiche l'état des produits en fin de cycle chargement
        If (Tab_Status_banc(0) And BIT7) <> 0 Then
            
            Call RAZ_ecran
            Call maj_etat_produit
            'Affichage fenetre défaut
            If ((Tab_Status_banc(7) <> 0)) Then 'And ((Tab_Status_banc(6) And BIT11) = 0 Or (Tab_Status_banc(6) And BIT12) = 0) Then 'Affichage défaut
                Call message_defaut_chargement
                Maj_HIM_vers_API_etat_banc = False
            Else
            reset_bit Copy_Tab(0), 7
            Maj_HIM_vers_API_etat_banc = True
            End If
            
            
        End If
        
        'Impression étiquette
        'Produit 1
         If ((Tab_Status_banc(6) And BIT1) <> 0) Then 'Demande imprétion étiquette produit 1
            'If (sans_etiquette = 0) Then
            '    Call lance_etiquette
            'End If
            reset_bit Copy_Tab(6), 1
            reset_bit Copy_Tab(6), 9 ' produit 1 bon
            Maj_HIM_vers_API_etat_banc = True
         End If
         
         'Produit 2
         If ((Tab_Status_banc(6) And BIT2) <> 0) Then 'Demande imprétion étiquette produit 2
            'If (sans_etiquette = 0) Then
            '    Call lance_etiquette
            'End If
            reset_bit Copy_Tab(6), 2
            reset_bit Copy_Tab(6), 10 ' produit 2 bon
            Maj_HIM_vers_API_etat_banc = True
         End If
         
                        
                
        If Maj_HIM_vers_API_etat_banc = True Then
            Result_com = Ecrir_mots_etat_banc(Copy_Tab(), 7)
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            Maj_HIM_vers_API_etat_banc = False
        End If
        
Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_controle. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_controle. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Sub appel_controle()
    ' Lance la fenêtre de contrôle
    Load FRM_controle
    FRM_startup.Hide
    FRM_controle.Show
    ' ajoute nouvelle entré dans la table "Log_banc" de la DB "Banc_XS_log"
    Ajout_nouvelle_barquette_log
End Sub

Sub Analyse_message()
Dim i As Integer
Dim deco_erreur_prod1 As Integer
Dim deco_erreur_prod2 As Integer
Dim Peut_afficher_U2 As Integer

Dim num_mes As Long
Dim Prod_Message() As Long

Dim Trame_message(5) As Long
Dim Message As Type_Message
Dim Prod_info As Info_produit

Dim N°enregistrement As Long ' Utiliser pour filtrer la de table de données jounal

''On Error GoTo 400

Prod_info.Code_comercial = symb_com
Prod_info.Code_lot = Num_lot

'Produit 1

'initialise variables globales
Lob_un_defaut = False
Log_un_defaut_majeur = False
Etiquette_un_defaut = False
Etiquette_un_defaut_majeur = False
Imprimer_etiquette = False

If (Tab_Status_banc(6) And BIT7) <> 0 Then 'Affichage des résultats si presence produit
    
    ' Lecture des message du produit 1
    num_mes = Lecture_messages_produit_1(Prod_Message)
    If num_mes <= 0 Then
        i = 1
    End If
    
    Prod_info.Mun_emprente = Tab_Status_banc(13)
    Prod_info.Num_produit = PRODUIT_1
    
    'Efface la zone d'affichage des résultats du produit 1
    FRM_controle.Affiche_prod_1.Clear
    
    'Ajout un nouveau produit à la base de donnée
    Ajoute_produit_log Prod_info
    
    If Prod_Message(0) > 0 Then
        For i = 0 To Prod_Message(0) - 1 'Iteration de touts les messages
            Trame_message(0) = Prod_Message((5 * i) + 1)
            Trame_message(1) = Prod_Message((5 * i) + 2)
            Trame_message(2) = Prod_Message((5 * i) + 3)
            Trame_message(3) = Prod_Message((5 * i) + 4)
            Trame_message(4) = Prod_Message((5 * i) + 5)
            
            If Decode_Trame_message(Trame_message(), Message) = True Then
                'L'idantification du message a été trouvé dans la table "Message"
                
                'Ajoute données a la base de donnée log
                'gére les défauts et positionne la demande d'impétion d'étiquette
                'affiche résultat à l'éran
                Ajoute_donnee_log Message, FRM_controle.Affiche_prod_1
            End If

        Next i
    Else
        Lob_un_defaut = True
        Log_un_defaut_majeur = True
    End If
    ' cloture de la basse de donnée log
    N°enregistrement = Valide_donnee_log
    
'====== DEBUT MODIF BEEA 21/12/2007 ============
    If Not ControleStationClaquageEnCours Then
        Imprimer_étiquette_post_final (N°enregistrement)
    End If
'====== DEBUT MODIF BEEA 21/12/2007 ============
    
    FRM_controle.Text1.Text = N°enregistrement
    
End If
    
'Produit 2

'initialise variables globales
Lob_un_defaut = False
Log_un_defaut_majeur = False
Etiquette_un_defaut = False
Etiquette_un_defaut_majeur = False
Imprimer_etiquette = False

If (Tab_Status_banc(6) And BIT8) <> 0 Then 'Affichage des résultats si presence produit
    ' Lecture des message du produit 2
    num_mes = Lecture_messages_produit_2(Prod_Message)
    If num_mes <= 0 Then
        i = 1
    End If
    
    Prod_info.Mun_emprente = Tab_Status_banc(14)
    Prod_info.Num_produit = PRODUIT_2
    
    'Efface la zone d'affichage des résultats du produit 2
    FRM_controle.Affiche_prod_2.Clear
    
    'Ajout un nouveau produit à la base de donnée
    Ajoute_produit_log Prod_info
    
    If Prod_Message(0) > 0 Then
        For i = 0 To Prod_Message(0) - 1 'Iteration de touts les messages
            Trame_message(0) = Prod_Message((5 * i) + 1)
            Trame_message(1) = Prod_Message((5 * i) + 2)
            Trame_message(2) = Prod_Message((5 * i) + 3)
            Trame_message(3) = Prod_Message((5 * i) + 4)
            Trame_message(4) = Prod_Message((5 * i) + 5)
            
            If Decode_Trame_message(Trame_message(), Message) = True Then
                'L'idantification du message a été trouvé dans la table "Message"
                
                'Ajoute données a la base de donnée log
                'gére les défauts et positionne la demande d'impétion d'étiquette
                'affiche résultat à l'éran
                Ajoute_donnee_log Message, FRM_controle.Affiche_prod_2
            End If
         
        Next i
    Else
        Lob_un_defaut = True
        Log_un_defaut_majeur = True
    End If
    ' cloture de la basse de donnée log
    N°enregistrement = Valide_donnee_log
    
'====== DEBUT MODIF BEEA 21/12/2007 ============
    If Not ControleStationClaquageEnCours Then
        Imprimer_étiquette_post_final (N°enregistrement)
    End If
'====== DEBUT MODIF BEEA 21/12/2007 ============
    
    FRM_controle.Text2.Text = N°enregistrement
    
End If

Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module Ecrit_resultat. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module Ecrit_resultat. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Public Function Affiche_Message(Etat As Boolean, Valeur As Double, ID_Message As Long, Afficheur As AfficheProj1.Affiche) As Variant

Dim Text As String
Dim Text_format_val As String

Dim i As Long

If DataEnvironment1.rsTable_Message.RecordCount > 0 Then
    If DataEnvironment1.rsTable_Message.Fields("Affiche_IHM").Value = True Then
        If (DataEnvironment1.rsTable_Message.Fields("Affiche_seulement_si_mauvais").Value = True And Not Etat) Or DataEnvironment1.rsTable_Message.Fields("Affiche_seulement_si_mauvais").Value = False Then
            Text = DataEnvironment1.rsTable_Message.Fields("Label").Value
            If DataEnvironment1.rsTable_Message.Fields("Valeur_utilise").Value = True Then
                Text = DataEnvironment1.rsTable_Message.Fields("Label_court").Value
                Text_format_val = "########0."
                For i = 0 To DataEnvironment1.rsTable_Message.Fields("Affichage_nb_decimale").Value - 1
                    Text_format_val = Text_format_val & "0"
                Next i
                
                Text = Text & " = " & Format(Valeur, Text_format_val) & " " & DataEnvironment1.rsTable_Message.Fields("Unite").Value
            End If
            If Etat = True Then
                If DataEnvironment1.rsTable_Message.Fields("Affiche_Gras_si_bon").Value = True Then
                    Afficheur.Gras = 1
                Else
                    Afficheur.Gras = -1
                End If
                
                Afficheur.Bare = -1
                
            Else
                If DataEnvironment1.rsTable_Message.Fields("Affiche_Gras_si_mauvais").Value = True Then
                    Afficheur.Gras = 1
                Else
                    Afficheur.Gras = -1
                End If
                
                If DataEnvironment1.rsTable_Message.Fields("Affiche_bare_si_mauvais").Value = True Then
                    Afficheur.Bare = 1
                Else
                    Afficheur.Bare = -1
                End If
                
            End If
            
            Afficheur.Add_text Text
            
        End If
    End If
Else
    Afficheur.Gras = 1
    Afficheur.Bare = -1
    Afficheur.Add_text "Idantifiant message inconnu N° " & ID_Message
    
End If

End Function

Sub Imprimer_étiquette_post_final(Num_enregistrement As Long)
    
    'Imprime etiquette si naissére
    Imprimer_etiquette = Imprimer_etiquette Or CBool(forçage_etiquette) 'Force l'imprétion de l'étiquette
    Imprimer_etiquette = Imprimer_etiquette And (Not CBool(sans_etiquette)) 'Interdi l'imprétion de l'étiquette
    If Imprimer_etiquette Then
       
        If DataEnvironment1.rsrequette_update.State <> adStateClosed Then
            DataEnvironment1.rsrequette_update.Close 'ferme le requette avant de l'ouvir
        End If
    
        DataEnvironment1.requette_update Num_enregistrement - 5
        If DataEnvironment1.rsrequette_update.RecordCount > 0 Then
        
            ' On cherche le N° enregistrement
            DataEnvironment1.rsrequette_update.MoveLast
            DataEnvironment1.rsrequette_update.Filter = "N°_enregistrement = " & DataEnvironment1.rsrequette_update.Fields("N°_enregistrement").Value
        Else
            Exit Sub
        End If
       
       
        Select Case Type_etiquette
            Case DB_PRODUIT_ANA_IS
                If Not Imprime_etiquette_ecran Then
                    DataReportANA_IS.PrintReport 'Imprime  l'étiquette
                Else
                    DataReportANA_IS.Show 'Affiche  l'étiquette à l'écran
                End If
            Case DB_PRODUIT_ANA_IS_IC
                If Not Imprime_etiquette_ecran Then
                    DataReportANA_IS_plus_IC.PrintReport 'Imprime  l'étiquette
                Else
                    DataReportANA_IS_plus_IC.Show 'Affiche  l'étiquette à l'écran
                End If
            Case DB_PRODUIT_TOR
                If Not Imprime_etiquette_ecran Then
                    DataReportTOR.PrintReport 'Imprime  l'étiquette
                Else
                    DataReportTOR.Show 'Affiche  l'étiquette à l'écran
                End If
            End Select
    End If
    DataEnvironment1.rsrequette_update.Filter = ("") ' Suprime le filtrage de la table de données journal
    
End Sub

Sub ecrit_resultat_autotest()
Dim ligne$

''On Error GoTo 400

'Premiere tension produit 1
    ligne$ = "Première tension:" & Format$(u_alim(0), "#0.00") & " V" & Chr$(13) & Chr$(10)
    ligne$ = ligne$ & "Ualim1 = " & Format$(CDbl(Tab_Status_banc(40)) / 100, "#0.00") & " V" & Chr$(13) & Chr$(10)
    ligne$ = ligne$ & "Icharge1 = " & Format$(CDbl(Tab_Status_banc(41)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
    ligne$ = ligne$ & "Icharge2 = " & Format$(CDbl(Tab_Status_banc(42)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
    
    If u_alim(1) <> 0 Then
        'Deuxieme tension produit 1
        ligne$ = ligne$ & "Deuxième tension :" & Format$(u_alim(1), "#0.00") & " V" & Chr$(13) & Chr$(10)
        ligne$ = ligne$ & "Ualim2 = " & Format$(CDbl(Tab_Status_banc(43)) / 100, "#0.00") & " V" & Chr$(13) & Chr$(10)
        ligne$ = ligne$ & "Icharge1 = " & Format$(CDbl(Tab_Status_banc(44)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
        ligne$ = ligne$ & "Icharge2 = " & Format$(CDbl(Tab_Status_banc(45)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
    End If
FRM_AUTOTEST!IMG_valeurs(0).Print ligne$

'Premiere tension produit 2
    ligne$ = "Première tension:" & Format$(u_alim(0), "#0.00") & " V" & Chr$(13) & Chr$(10)
    ligne$ = ligne$ & "Ualim1 = " & Format$(CDbl(Tab_Status_banc(46)) / 100, "#0.00") & " V" & Chr$(13) & Chr$(10)
    ligne$ = ligne$ & "Icharge1 = " & Format$(CDbl(Tab_Status_banc(47)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
    ligne$ = ligne$ & "Icharge2 = " & Format$(CDbl(Tab_Status_banc(48)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
    
    If u_alim(1) <> 0 Then
        'Deuxieme tension produit 1
        ligne$ = ligne$ & "Deuxième tension:" & Format$(u_alim(1), "#0.00") & " V" & Chr$(13) & Chr$(10)
        ligne$ = ligne$ & "Ualim2 = " & Format$(CDbl(Tab_Status_banc(49)) / 100, "#0.00") & " V" & Chr$(13) & Chr$(10)
        ligne$ = ligne$ & "Icharge1 = " & Format$(CDbl(Tab_Status_banc(50)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
        ligne$ = ligne$ & "Icharge2 = " & Format$(CDbl(Tab_Status_banc(51)) / 100, "#0.00") & " mA" & Chr$(13) & Chr$(10)
    End If
FRM_AUTOTEST!IMG_valeurs(1).Print ligne$
Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module Ecrit_resultat_autotest. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module Ecrit_resultat_autotest. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Sub Maj_etat_banc()
Dim index_trame, index_table
Dim relance, i, j, k

''On Error GoTo 200
    
    'W370,W371,W372
    cycle_chargement = Tab_Status_banc(64)
    cycle_claquage = Tab_Status_banc(65)
    cycle_final = Tab_Status_banc(66)
    
    ' Gestion des défauts MW354:X0 à X5, MW357,MW358,W359
    com_err_tsx = Tab_Status_banc(4) And (BIT0 + BIT1 + BIT2 + BIT3 + BIT4 + BIT5)
    def_chargement = Tab_Status_banc(7)
    def_claquage = Tab_Status_banc(8)
    def_final = Tab_Status_banc(9)
    
    ' Nouvelle barquette de produits
    If (Tab_Status_banc(0) And BIT2) <> 0 Then
        nouvelle_barquette = True
    End If
                
    Exit Sub
                
200 MsgBox "Erreur " & Str$(err) & " module Maj_etat_banc. Signalez immediatement cette erreur. Merci."
    Resume Next

End Sub

Sub maj_boite_dial_qualification()

Static old_code As Long

    If Tab_Status_banc(24) > 0 And (old_code <> Tab_Status_banc(24)) Then
        old_code = Tab_Status_banc(24)
        FRM_Message_mes_electrique_qualification.Show
        If old_code < 30000 Then
            FRM_Message_mes_electrique_qualification.Set_message CInt(old_code)
        End If
    End If
    
    If Tab_Status_banc(24) = 0 And (old_code <> 0) Then
        old_code = 0
        Unload FRM_Message_mes_electrique_qualification
    End If

End Sub


Sub maj_boite_dial_operateur()
Dim NL$

''On Error GoTo 400

    NL$ = Chr$(13) & Chr$(10)
    FRM_controle!LBL_dialogue_operateur.BackColor = &HFFFFC0
    If Tab_Status_banc(5) = 0 Then
        FRM_controle!LBL_dialogue_operateur.Caption = ""
        FRM_controle!LBL_dialogue_operateur.Visible = False
    Else
        FRM_controle!LBL_dialogue_operateur.Visible = True
        
    End If
    
    If Tab_Status_banc(5) = 10 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Enlever le produit 1 et le mettre dans la goulotte"
        FRM_controle!LBL_dialogue_operateur.Caption = "Ambil product POS 1 dan letakkan di reject bin."
    End If
    
    If Tab_Status_banc(5) = 11 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Remettre le produit 1 car il est bon"
        FRM_controle!LBL_dialogue_operateur.Caption = "Taruh kembali product pada POS 1 karena masih GOOD."
    End If
    
    If Tab_Status_banc(5) = 12 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Produit 1 en court-circuit (confirmation ou annulation"
        FRM_controle!LBL_dialogue_operateur.Caption = "Product POS 1 short-circuit (OK atau CANCEL)?"
    End If
    
    If Tab_Status_banc(5) = 13 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Mettre le produit 1 dans la goulotte"
        FRM_controle!LBL_dialogue_operateur.Caption = "Taruh product POS 1 di reject bin"
    End If
    
    If Tab_Status_banc(5) = 14 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Produit 1 mauvais : non commutation"
        FRM_controle!LBL_dialogue_operateur.Caption = "Product POS 1 FAIL : no switching"
    End If
    
    If Tab_Status_banc(5) = 15 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Produit 1 mauvais : commutation permanente"
        FRM_controle!LBL_dialogue_operateur.Caption = "Product POS 1 FAIL : NC"
    End If
    
    '========DEBUT MODIF BEEA 26/12/2007 ==========
    If Tab_Status_banc(5) = 16 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Enlever les produits de test claquage"
        FRM_controle!LBL_dialogue_operateur.Caption = "Ambil product dummy test HIPOT"
    End If
    '=======FIN MODIF BEEA 26/12/2007 ==========
    
    If Tab_Status_banc(5) = 20 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Enlever le produit 2 et le mettre dans la goulotte"
        FRM_controle!LBL_dialogue_operateur.Caption = "Ambil product POS 2 dan letakkan di reject bin"
    End If
    
    If Tab_Status_banc(5) = 21 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Remettre le produit 2 car il est bon"
        FRM_controle!LBL_dialogue_operateur.Caption = "Taruh kembali product POS 2 karena GOOD"
    End If
    
    If Tab_Status_banc(5) = 22 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Produit 2 en court-circuit (confirmation ou annulation"
        FRM_controle!LBL_dialogue_operateur.Caption = "Product POS 2 short-circuit (OK or CANCEL)?"
    End If
    
    If Tab_Status_banc(5) = 23 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Mettre le produit 2 dans la goulotte"
        FRM_controle!LBL_dialogue_operateur.Caption = "Taruh product POS 2 di reject bin"
    End If
    
    If Tab_Status_banc(5) = 24 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Produit 2 mauvais : non commutation"
        FRM_controle!LBL_dialogue_operateur.Caption = "Product POS 2 FAIL : no switching"
    End If
    
    If Tab_Status_banc(5) = 25 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Produit 2 mauvais : commutation permanente"
        FRM_controle!LBL_dialogue_operateur.Caption = "Product POS 2 FAIL : NC"
    End If
    
    If Tab_Status_banc(5) = 30 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Enlever les produits 1 et 2 et les mettre dans la goulotte"
        FRM_controle!LBL_dialogue_operateur.Caption = "Ambil product POS 1 dan POS 2 lalu letakkan di reject bin"
    End If
    
    If Tab_Status_banc(5) = 31 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Enlever le ou les produit bons "
        FRM_controle!LBL_dialogue_operateur.Caption = "Ambil product yang GOOD "
    End If
    
    If Tab_Status_banc(5) = 40 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Attente chargement produit" & NL$ & "Appuyer sur le bouton vert "
        FRM_controle!LBL_dialogue_operateur.Caption = "Silahkan taruh product" & NL$ & "lalu Tekan tombol HIJAU."
    End If
    
    If Tab_Status_banc(5) = 50 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Appuyer sur le bouton bleu : Départ rotation"
        FRM_controle!LBL_dialogue_operateur.Caption = "Tekan Tombol BIRU : PUTAR"
    End If
    
    If Tab_Status_banc(5) = 51 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Appuyer sur le bouton vert : Départ cycle"
        FRM_controle!LBL_dialogue_operateur.Caption = "Tekan tombol HIJAU : Mulai TEST"
    End If
    
    If Tab_Status_banc(5) = 60 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Mettre un support pour produit normal"
        FRM_controle!LBL_dialogue_operateur.Caption = "Taruh support untuk product normal"
        FRM_controle!LBL_dialogue_operateur.BackColor = &HFF&
    End If
    
    If Tab_Status_banc(5) = 62 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Mettre un support pour produit miniature"
        FRM_controle!LBL_dialogue_operateur.Caption = "Taruh support untuk product miniature"
        FRM_controle!LBL_dialogue_operateur.BackColor = &HFF&
    
    End If
    
    If Tab_Status_banc(5) = 250 Then
        'FRM_controle!LBL_dialogue_operateur.Caption = "Appuyer sur le bouton marche coups à coups : Cycle poste final en pause"
        FRM_controle!LBL_dialogue_operateur.Caption = "Tekan tombol START step by step : Final Test cycle pause"
        FRM_controle!LBL_dialogue_operateur.BackColor = &HFF&
    
    End If
    
    Exit Sub
    
'400     MsgBox "Erreur " & Str$(err) & " module maj_boite_dial_operateur. Signalez immediatement cette erreur. Merci."
400     MsgBox "Erreur " & Str$(err) & " module maj_boite_dial_operateur. Beritahu ke maintenance soal ERROR ini. Terima Kasih."
    Resume Next
    
End Sub

Sub maj_boite_dial_operateur_autotest()

''On Error GoTo 400

    FRM_AUTOTEST!Lbl_message_operateur.Caption = ""
    
    If (Tab_Status_banc(0) And BIT8) <> 0 Then
    
        If (Tab_Status_banc(69) = 10) Then
            ualim = 1
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Test communication SEFELEC"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Test communication SEFELEC"
        End If
        
        If ((Tab_Status_banc(69) > 10) And (Tab_Status_banc(69) <= 30)) Then
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Configuration coffret et alimentation HP"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Configure Power Supply HP"
        End If
        
        If (((Tab_Status_banc(69) > 30) And (Tab_Status_banc(69) <= 36)) And (ualim = 1)) Then
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Mesure Tension d'alimentation 1"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Ukur Tegangan 1"
        End If
        
        If ((Tab_Status_banc(69) > 36) And (Tab_Status_banc(69) <= 41)) Then
            ualim = 2
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Mesure courant de charge première sortie"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Ukur Load current di output 1"
        End If
        
        If ((Tab_Status_banc(69) > 41) And (Tab_Status_banc(69) <= 46)) Then
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Mesure courant de charge deuxiéme sortie"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Ukur Load current di output 2"
        End If
        
        If (((Tab_Status_banc(69) > 30) And (Tab_Status_banc(69) <= 36)) And (ualim = 2)) Then
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Mesure Tension d'alimentation 2"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Ukur Tegangan 2"
        End If
        
        If ((Tab_Status_banc(69) > 46) And (Tab_Status_banc(69) <= 110)) Then
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Arret alimentation"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Stop power supply"
        End If
        
        If ((Tab_Status_banc(69) > 200) And (Tab_Status_banc(69) <= 230)) Then
            ualim = 1
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Positionne poste en conditions initiales puis rotation plateau"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Position station di kondisi awal lalu putar plate"
        End If
        
        If ((Tab_Status_banc(69) > 230) And (Tab_Status_banc(69) <= 240)) Then
            ualim = 1
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Rotation plateau"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Putar Plate"
        End If
        
        If ((Tab_Status_banc(69) = 245)) Then
            'FRM_AUTOTEST!Lbl_message_operateur.Caption = "Appuyer sur le BP bleu pour faire tourner le plateau ou sur le BP vert pour afficher les résultats"
            FRM_AUTOTEST!Lbl_message_operateur.Caption = "Tekan tombol BIRU untuk memutar plate or tombol HIJAU untuk menampilkan hasil test"
        End If
    End If
    
    Exit Sub
    
'400     MsgBox "Erreur " & Str$(err) & " module maj_boite_dial_operateur_autotest. Signalez immediatement cette erreur. Merci."
400     MsgBox "Erreur " & Str$(err) & " module maj_boite_dial_operateur_autotest. Beritahu ke maintenance soal ERROR ini. Terima Kasih."
    Resume Next
    
End Sub

Sub maj_ecran()
Dim NL$

    ''On Error GoTo 400
    
    NL$ = Chr$(13) & Chr$(10)
    
    ' W351 W352 mise à jour nbre prod bons et mauvais
        'FRM_controle!Lbl_bon.Caption = "Nombre de produits bons : " & Str$(Tab_Status_banc(1))
        'FRM_controle!Lbl_mauvais.Caption = "Nombre de produits mauvais : " & Str$(Tab_Status_banc(2))
        FRM_controle!Lbl_bon.Caption = "QTY GOOD Product : " & Str$(Tab_Status_banc(1))
        FRM_controle!Lbl_mauvais.Caption = "QTY FAIL Product : " & Str$(Tab_Status_banc(2))
        
        
        ''' Disini Update Traceability quantity
        On Error GoTo exError
        If Not FRM_controle.XsClient Is Nothing Then
           FRM_controle.XsClient.SetOutputQuantity Tab_Status_banc(1)
           FRM_controle.XsClient.SetRejectQuantity Tab_Status_banc(2)
        GoTo notError
        End If
        
exError: FRM_controle.Label1.Caption = Error
notError:
        
    ' W418 W417 W419 mise à jour n° d'etape TSX
        FRM_controle.StatusBar1.Panels(1).Text = Tab_Status_banc(68)
        
        If DataEnvironment1.rsTable_liste_fonction_final.State = adStateClosed Then
            DataEnvironment1.rsTable_liste_fonction_final.Open
        End If
        If DataEnvironment1.rsTable_liste_etape_final.State = adStateClosed Then
            DataEnvironment1.rsTable_liste_etape_final.Open
        End If
        
        'cherche si le numéro étape se trouve dans la table "liste_etape_final"
        'si oui affiche le labelle asocié
        'si non le code etape doit être un code fonction
        DataEnvironment1.rsTable_liste_etape_final.Filter = "Code = '" & Tab_Status_banc(69) & "'"
        If DataEnvironment1.rsTable_liste_etape_final.RecordCount > 0 Then
            DataEnvironment1.rsTable_liste_etape_final.MoveFirst
            'affiche le labelle
            FRM_controle.StatusBar1.Panels(3).Text = DataEnvironment1.rsTable_liste_etape_final.Fields("Labelle").Value
        Else
            'code fonction = code etape - ETAPE_GESTION_CYCLE_POST_FINAL
            'cherche si le code fonction se trouve dans la base de donnée
            'si affiche le labelle asocié à ce code
            'sinon affiche de code
            DataEnvironment1.rsTable_liste_fonction_final.Filter = "Code = '" & Tab_Status_banc(69) - ETAPE_GESTION_CYCLE_POST_FINAL & "'"
            If DataEnvironment1.rsTable_liste_fonction_final.RecordCount > 0 Then
                DataEnvironment1.rsTable_liste_fonction_final.MoveFirst
                'affiche le labelle de la fonction
                FRM_controle.StatusBar1.Panels(3).Text = DataEnvironment1.rsTable_liste_fonction_final.Fields("Labelle").Value
            Else
                'affiche le numméro de la fonction
                FRM_controle.StatusBar1.Panels(2).Text = Tab_Status_banc(69)
            End If 'fonction
        End If 'etape
        
        FRM_controle.StatusBar1.Panels(2).Text = Tab_Status_banc(67)
        
    'Mise à jour boite de dialogue operateur
        Call maj_boite_dial_operateur
    
    'Maj boite dialogue qualification
        maj_boite_dial_qualification
    
    'Affichage BP déclaration produits mauvais
        If ((Tab_Status_banc(67) >= 30) And (Tab_Status_banc(67) < 35)) Then
                If (Tab_Status_banc(6) And BIT7) <> 0 Then
                    FRM_controle!Cmd_prod1_mauvais.Visible = True
                Else
                    FRM_controle!Cmd_prod1_mauvais.Visible = False
                End If
                
                If (Tab_Status_banc(6) And BIT8) <> 0 Then
                    FRM_controle!Cmd_prod2_mauvais.Visible = True
                Else
                    FRM_controle!Cmd_prod2_mauvais.Visible = False
                End If
        End If
        
        Const COLOR_CYCLE_PAS_A_PAS = &H80FF&
        Const COLOR_CYCLE_AUTO = &HFF00&
        Const COLOR_CYCLE_EN_COURS = &HFF0000
        Const COLOR_CYCLE_TERMINE = &HFF00&
        Const COLOR_CYCLE_EN_DEFAUT = &HFF&
        Const COLOR_CYCLE_NON_VALIDE = &H808080
    
            
    'Mise à jour état cycle chargement
        If cycle_chargement_precedent <> cycle_chargement Then
            cycle_chargement_precedent = cycle_chargement
            If (cycle_chargement And BIT4) <> 0 Then 'PAS à PAS
                    FRM_controle!LBL_etat_cycle_chargement.BackColor = COLOR_CYCLE_AUTO
                    'FRM_controle!LBL_etat_cycle_chargement.Caption = " Chargement" & NL$ & "en PAS à PAS"
                    FRM_controle!LBL_etat_cycle_chargement.Caption = " Loading" & NL$ & "in STEP by STEP"
            Else
                FRM_controle!LBL_etat_cycle_chargement.BackColor = COLOR_CYCLE_NON_VALIDE
                'FRM_controle!LBL_etat_cycle_chargement.Caption = " Chargement" & NL$ & "non valide"
                FRM_controle!LBL_etat_cycle_chargement.Caption = " Loading" & NL$ & "not OK"
                
                If (cycle_chargement And BIT3) <> 0 Then 'mode auto
                    FRM_controle!LBL_etat_cycle_chargement.BackColor = COLOR_CYCLE_NON_VALIDE
                    'FRM_controle!LBL_etat_cycle_chargement.Caption = " Chargement" & NL$ & "en AUTO"
                    FRM_controle!LBL_etat_cycle_chargement.Caption = " Loading" & NL$ & "in AUTO"
               
                    If (cycle_chargement And BIT10) <> 0 Then 'cycle en cours
                        FRM_controle!LBL_etat_cycle_chargement.BackColor = COLOR_CYCLE_EN_COURS
                        'FRM_controle!LBL_etat_cycle_chargement.Caption = " Chargement" & NL$ & "en cours"
                        FRM_controle!LBL_etat_cycle_chargement.Caption = " Loading" & NL$ & "in progress"
                    End If
                    If (cycle_chargement And BIT12) <> 0 Then 'cycle terminé
                        FRM_controle!LBL_etat_cycle_chargement.BackColor = COLOR_CYCLE_TERMINE
                        'FRM_controle!LBL_etat_cycle_chargement.Caption = " Chargement" & NL$ & "terminé"
                        FRM_controle!LBL_etat_cycle_chargement.Caption = " Loading" & NL$ & "Finish"
                    End If
                    If (cycle_chargement And BIT13) <> 0 Then 'post en défaut
                        FRM_controle!LBL_etat_cycle_chargement.BackColor = COLOR_CYCLE_EN_DEFAUT
                        'FRM_controle!LBL_etat_cycle_chargement.Caption = " Chargement" & NL$ & "en défaut"
                        FRM_controle!LBL_etat_cycle_chargement.Caption = " Loading" & NL$ & "in ERROR"
                    End If
                End If
            End If
        End If
        
    'Mise à jour état cycle claquage
        If cycle_claquage_precedent <> cycle_claquage Then
            cycle_claquage_precedent = cycle_claquage
            If (cycle_claquage And BIT4) <> 0 Then 'PAS à PAS
                    FRM_controle!LBL_etat_cycle_claquage.BackColor = COLOR_CYCLE_PAS_A_PAS
                    'FRM_controle!LBL_etat_cycle_claquage.Caption = "  Claquage" & NL$ & "en PAS à PAS"
                    FRM_controle!LBL_etat_cycle_claquage.Caption = "  HIPOT" & NL$ & "in STEP by STEP"
            Else
                FRM_controle!LBL_etat_cycle_claquage.BackColor = COLOR_CYCLE_NON_VALIDE
                'FRM_controle!LBL_etat_cycle_claquage.Caption = "  Claquage" & NL$ & "non valide"
                FRM_controle!LBL_etat_cycle_claquage.Caption = "  HIPOT" & NL$ & "not OK"
                
                If (cycle_claquage And BIT3) <> 0 Then
                    FRM_controle!LBL_etat_cycle_claquage.BackColor = COLOR_CYCLE_AUTO
                    'FRM_controle!LBL_etat_cycle_claquage.Caption = "  Claquage" & NL$ & "en AUTO"
                    FRM_controle!LBL_etat_cycle_claquage.Caption = "  HIPOT" & NL$ & "in AUTO"
                    
                    If (cycle_claquage And BIT10) <> 0 Then
                        FRM_controle!LBL_etat_cycle_claquage.BackColor = COLOR_CYCLE_EN_COURS
                        'FRM_controle!LBL_etat_cycle_claquage.Caption = "  Claquage" & NL$ & "en cours"
                        FRM_controle!LBL_etat_cycle_claquage.Caption = "  HIPOT" & NL$ & "in progress"
                    End If
                    If (cycle_claquage And BIT12) <> 0 Then
                        FRM_controle!LBL_etat_cycle_claquage.BackColor = COLOR_CYCLE_TERMINE
                        'FRM_controle!LBL_etat_cycle_claquage.Caption = "  Claquage" & NL$ & "terminé"
                        FRM_controle!LBL_etat_cycle_claquage.Caption = "  HIPOT" & NL$ & "finish"
                        'FRM_controle!LBL_etat_cycle_claquage.ForeColor = vbBlack
                    End If
                    If (cycle_claquage And BIT13) <> 0 Then
                        FRM_controle!LBL_etat_cycle_claquage.BackColor = COLOR_CYCLE_EN_DEFAUT
                        'FRM_controle!LBL_etat_cycle_claquage.Caption = "  Claquage" & NL$ & "en défaut"
                        FRM_controle!LBL_etat_cycle_claquage.Caption = "  HIPOT" & NL$ & "in error"
                    End If
                End If
            End If
        End If
    
    'Mise à jour état cycle controle final
        If cycle_final_precedent <> cycle_final Then
           cycle_final_precedent = cycle_final
            If (cycle_final And BIT4) <> 0 Then 'PAS à PAS
                FRM_controle!LBL_etat_cycle_ctrl_final.BackColor = COLOR_CYCLE_PAS_A_PAS
                'FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "Controle final" & NL$ & "en PAS à PAS"
                FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "FINAL TEST" & NL$ & "in STEP by STEP"
            Else
                FRM_controle!LBL_etat_cycle_ctrl_final.BackColor = COLOR_CYCLE_NON_VALIDE
                'FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "Controle final" & NL$ & "non valide"
                FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "FINAL TEST" & NL$ & "not OK"
                
                If (cycle_final And BIT3) <> 0 Then
                    FRM_controle!LBL_etat_cycle_ctrl_final.BackColor = COLOR_CYCLE_AUTO
                    'FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "Controle final" & NL$ & "en AUTO"
                    FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "FINAL TEST" & NL$ & "in AUTO"
                
                    If (cycle_final And BIT10) <> 0 Then
                        FRM_controle!LBL_etat_cycle_ctrl_final.BackColor = COLOR_CYCLE_EN_COURS
                        'FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "Controle final" & NL$ & "en cours"
                        FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "FINAL TEST" & NL$ & "in progress"
                    End If
                    If (cycle_final And BIT12) <> 0 Then
                        FRM_controle!LBL_etat_cycle_ctrl_final.BackColor = COLOR_CYCLE_TERMINE
                        'FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "Controle final" & NL$ & "terminé"
                        FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "FINAL TEST" & NL$ & "finish"
                    End If
                    If (cycle_final And BIT13) <> 0 Then
                        FRM_controle!LBL_etat_cycle_ctrl_final.BackColor = COLOR_CYCLE_EN_DEFAUT
                        'FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "Controle final" & NL$ & "en défaut"
                        FRM_controle!LBL_etat_cycle_ctrl_final.Caption = "FINAL TEST" & NL$ & "in error"
                    End If
                End If
            End If
        End If
    Exit Sub
    
'400     MsgBox "Erreur " & Str$(err) & " module maj_ecran. Signalez immediatement cette erreur. Merci."
400     MsgBox "Erreur " & Str$(err) & " module maj_ecran. Beritahu ke maintenance soal ERROR ini. Terima Kasih."
    Resume Next
       
End Sub

Sub RAZ_ecran()

    FRM_controle!SHP_voyant_vert(0).Visible = False
    FRM_controle!LBL_pbon(0).Visible = False
        
    FRM_controle!SHP_voyant_vert(1).Visible = False
    FRM_controle!LBL_pbon(1).Visible = False
        
    FRM_controle.Affiche_prod_1.Clear
    FRM_controle.Affiche_prod_2.Clear
        
    FRM_controle!Cmd_prod1_mauvais.Visible = False
    FRM_controle!Cmd_prod2_mauvais.Visible = False
        
End Sub

Sub maj_etat_produit()

''On Error GoTo 400
    
    Const COLOR_PRODUIT_BON = &HFF00&
    Const COLOR_PRODUIT_PAS_SUR = &H80FF&
    Const COLOR_PRODUIT_MAUVAIS = &HFF&
    Const COLOR_PRODUIT_ABANDONNE = &H808080
    
    Const ETAPE_ATTENTE_DECHARGEMENT = 10
    Const ETAPE_FIN_CHARGEMENT = 41
    Const ETAPE_DEBUT_CHARGEMENT = 20
    
'=======DEBUT MODIF BEEA 30/11/2007================
    ' Affichage resultats en fin de cycle
    If ((Tab_Status_banc(0) And (BIT3 + BIT7)) <> 0) And (com_err_tsx = 0) Then
        ' Affichage resultats selon présence produit(s)
        If (Tab_Status_banc(6) And BIT7) <> 0 Then 'produit 1 présent
            FRM_controle!SHP_voyant_vert(0).Visible = True
            FRM_controle!LBL_pbon(0).Visible = True
        Else
            FRM_controle!SHP_voyant_vert(0).Visible = False
            FRM_controle!LBL_pbon(0).Visible = False
        End If

        If (Tab_Status_banc(6) And BIT8) <> 0 Then 'produit 2 présent
            FRM_controle!SHP_voyant_vert(1).Visible = True
            FRM_controle!LBL_pbon(1).Visible = True
        Else
            FRM_controle!SHP_voyant_vert(1).Visible = False
            FRM_controle!LBL_pbon(1).Visible = False
        End If
    
        'Fin de cycle chargement
        If ((Tab_Status_banc(0) And BIT7) <> 0) Then
            'Etat produit 1
            If (Tab_Status_banc(6) And BIT11) <> 0 Then 'produit 1 bon par TSX
                FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_BON
                FRM_controle!LBL_pbon(0).Caption = "GOOD"
            Else
                FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_MAUVAIS
                FRM_controle!LBL_pbon(0).Caption = "FAIL"
            End If
            'Etat produit 2
            If ((Tab_Status_banc(0) And BIT7) <> 0) Then
                If (Tab_Status_banc(6) And BIT12) <> 0 Then 'produit 2 bon par TSX
                    FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_BON
                    FRM_controle!LBL_pbon(1).Caption = "GOOD"
                Else
                    FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_MAUVAIS
                    FRM_controle!LBL_pbon(1).Caption = "FAIL"
                End If
            End If
        End If
        
        'Fin de cycle de contrôle
        If ((Tab_Status_banc(0) And BIT3) <> 0) Then
            'Etat produit 1
            If (Tab_Status_banc(6) And BIT11) <> 0 Then 'produit 1 bon par TSX
                If ControleStationClaquageEnCours = False Then
                    Select Case ModeMarcheDielec
                        Case 0, 3
                            FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_BON
                            FRM_controle!LBL_pbon(0).Caption = "GOOD"
                        Case 1
                            FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_PAS_SUR
                            FRM_controle!LBL_pbon(0).Caption = "? HIPOT ?"
                        Case 2, 4, 5
                            FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_PAS_SUR
                            FRM_controle!LBL_pbon(0).Caption = "HIPOT HS"
                    End Select
                Else
                    FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_MAUVAIS
                    FRM_controle!LBL_pbon(0).Caption = "NOT TEST"
                    TestStationClaquage1OK = False
                End If
            Else
                If ControleStationClaquageEnCours = False Then
                    FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_MAUVAIS
                    FRM_controle!LBL_pbon(0).Caption = "FAIL"
                Else
                    FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_BON
                    FRM_controle!LBL_pbon(0).Caption = "GOOD"
                    TestStationClaquage1OK = True
                End If
            End If
            'Etat produit 2
            If (Tab_Status_banc(6) And BIT12) <> 0 Then 'produit 2 bon par TSX
                If ControleStationClaquageEnCours = False Then
                    Select Case ModeMarcheDielec
                        Case 0, 2
                            FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_BON
                            FRM_controle!LBL_pbon(1).Caption = "GOOD"
                        Case 1
                            FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_PAS_SUR
                            FRM_controle!LBL_pbon(1).Caption = "? HIPOT ?"
                        Case 3, 4, 5
                            FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_PAS_SUR
                            FRM_controle!LBL_pbon(1).Caption = "HIPOT HS"
                    End Select
                Else
                    FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_MAUVAIS
                    FRM_controle!LBL_pbon(1).Caption = "NOT TEST"
                    TestStationClaquage2OK = False
                End If
            Else
                If ControleStationClaquageEnCours = False Then
                    FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_MAUVAIS
                    FRM_controle!LBL_pbon(1).Caption = "FAIL"
                Else
                    FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_BON
                    FRM_controle!LBL_pbon(1).Caption = "GOOD"
                    TestStationClaquage2OK = True
                End If
            End If
        End If
        
        'Attente déchargement sur cycle interrompu
        If (Tab_Status_banc(67) <> ETAPE_ATTENTE_DECHARGEMENT) And ((Tab_Status_banc(67) < ETAPE_DEBUT_CHARGEMENT) Or (Tab_Status_banc(67) > ETAPE_FIN_CHARGEMENT)) Then
            'Produit 1
            If (Tab_Status_banc(20) And (BIT12 + BIT11)) <> 0 Then 'produit abandonné
                FRM_controle!SHP_voyant_vert(0).FillColor = COLOR_PRODUIT_ABANDONNE
                FRM_controle!LBL_pbon(0).Caption = "NOT TESTED"
            End If
            'Produit 2
            If (Tab_Status_banc(22) And (BIT12 + BIT11)) <> 0 Then
                FRM_controle!SHP_voyant_vert(1).FillColor = COLOR_PRODUIT_ABANDONNE
                FRM_controle!LBL_pbon(1).Caption = "NOT TESTED"
            End If
        End If
    
        'Fin de cycle de contrôle de la station de claquage
        If ControleStationClaquageEnCours And (Tab_Status_banc(0) And BIT3) <> 0 Then
            With FRM_controle
                'Les 2 stations sont OK = le poste est bon, retour à la normale
                If TestStationClaquage1OK And TestStationClaquage2OK Then
                       ModeMarcheDielec = 0
                End If
                '1 des station est HS = le poste est mauvais, retour à l'état de demande
                If (Not TestStationClaquage1OK) Or (Not TestStationClaquage2OK) Then
                       ModeMarcheDielec = 5
                End If
            End With
        End If
'=======FIN MODIF BEEA 30/11/2007==================
    End If
    
    'garder le focus sur la feuille contrôle, sauf en fin de contrôle
    If Not FinControleAffiche Or Not FinTestClaquageAffiche Then
        FRM_controle.SetFocus
    End If
    
Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_etat_produit. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_etat_produit. Laporkan ke maintenance segera. Thanks."
Resume Next
    
End Sub

'=======DEBUT MODIF BEEA 30/11/2007================
Sub EcritModeTestStationClaquage(indModClaqu As Integer)
On Error GoTo erreur
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_Frequentiel.State <> adStateClosed Then
        DataEnvironment1.rsTable_Frequentiel.Close
    End If
    DataEnvironment1.rsTable_Frequentiel.Open
    
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.State <> adStateClosed Then
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Close
    End If
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Open
    
    Select Case indModClaqu
        Case 0
        'Mise à jour du frequentiel pour le test de la station de claquage
        DataEnvironment1.rsTable_Frequentiel.AddNew
        DataEnvironment1.rsTable_Frequentiel.Fields("frequence").Value = Frequence_Test_Claqu
        DataEnvironment1.rsTable_Frequentiel.Fields("heure_demande").Value = Heure_Test_Claqu
        DataEnvironment1.rsTable_Frequentiel.Fields("derniere_modif").Value = Derniere_Modif_Param_Test_Claqu
        DataEnvironment1.rsTable_Frequentiel.Fields("modifie_par").Value = Modifie_Par_Test_Claqu
        DataEnvironment1.rsTable_Frequentiel.Fields("derniere_execution").Value = Now
        DataEnvironment1.rsTable_Frequentiel.Fields("prochaine_execution").Value = Int(Now) + Frequence_Test_Claqu + Heure_Test_Claqu
        DataEnvironment1.rsTable_Frequentiel.Update
        'Mise à jour du mode de marche pour le test de la station de claquage
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.MoveLast
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Fields("MODE_ACTUEL_CLAQUAGE").Value = indModClaqu
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Update
        
        Case Else
        'Mise à jour du mode de marche pour le test de la station de claquage
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.MoveLast
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Fields("MODE_ACTUEL_CLAQUAGE").Value = indModClaqu
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Update
        
    End Select

Exit Sub
erreur:
'reponse = MsgBox("erreur procédure écriture mode test station de claquage", vbCritical, "ERREUR")
reponse = MsgBox("Error procedure writing test mode in station Hipot", vbCritical, "ERREUR")

End Sub
'=======FIN MODIF BEEA 30/11/2007==================

'=======DEBUT MODIF BEEA 30/11/2007================
Sub LitModeTestStationClaquage()
On Error GoTo erreur
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_Frequentiel.State <> adStateClosed Then
        DataEnvironment1.rsTable_Frequentiel.Close
    End If
    DataEnvironment1.rsTable_Frequentiel.Open
    
    DataEnvironment1.rsTable_Frequentiel.MoveLast
    Prochaine_Exec_Test_Claqu = DataEnvironment1.rsTable_Frequentiel.Fields("prochaine_execution").Value
    
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.State <> adStateClosed Then
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Close
    End If
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Open
    
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.MoveLast
    ModeMarcheDielec = DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Fields("MODE_ACTUEL_CLAQUAGE").Value
    Derniere_execution_Test_Claqu_OK = DataEnvironment1.rsTable_Frequentiel.Fields("derniere_execution").Value
    Prochaine_Exec_Test_Claqu = DataEnvironment1.rsTable_Frequentiel.Fields("prochaine_execution").Value
    Derniere_Modif_Param_Test_Claqu = DataEnvironment1.rsTable_Frequentiel.Fields("derniere_modif").Value
    Modifie_Par_Test_Claqu = DataEnvironment1.rsTable_Frequentiel.Fields("modifie_par").Value
    Frequence_Test_Claqu = DataEnvironment1.rsTable_Frequentiel.Fields("frequence").Value
    Heure_Test_Claqu = DataEnvironment1.rsTable_Frequentiel.Fields("heure_demande").Value
        
    With FRM_startup
        If (Now >= Prochaine_Exec_Test_Claqu) Or (ModeMarcheDielec = 1) Or (ModeMarcheDielec = 5) Then
            .Frame3D1.Enabled = False
            .Frame3D1.Visible = False
            .fraTestStationClaquage.Enabled = True
            .fraTestStationClaquage.Visible = True
            .cmdTestClaquage.Visible = False
            .cmdTestClaquage.Enabled = False
            If (Now >= Prochaine_Exec_Test_Claqu) Then
                ModeMarcheDielec = 1
                DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Filter = ("")
                DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.MoveLast
                DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Fields("MODE_ACTUEL_CLAQUAGE").Value = ModeMarcheDielec
                DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Update
            End If
        Else
            .Frame3D1.Enabled = True
            .Frame3D1.Visible = True
            .fraTestStationClaquage.Enabled = False
            .fraTestStationClaquage.Visible = False
            .cmdTestClaquage.Visible = False
            .cmdTestClaquage.Enabled = False
        End If
    End With
    
Exit Sub
erreur:
'reponse = MsgBox("erreur procédure lecture mode test station de claquage", vbCritical, "ERREUR")
reponse = MsgBox("Error procedure reading test mode in station Hipot", vbCritical, "ERREUR")
    
End Sub
'=======FIN MODIF BEEA 30/11/2007==================

Sub maj_text_limites()
Dim NL$
Dim inter$
Dim forme$
Dim text_courant As String

''On Error GoTo 400

' AC ou DC ?
If ac Then forme$ = "~" Else forme$ = "="

'Remplissage de la boite texte où figure les limites
'de la Base de Donnees

'-> SODATEC 9/5/1
'mode non analogique
If si_ctrl_tor Then
'<- SODATEC
        NL$ = Chr$(13) & Chr$(10)
     
     If type_controle = CONTROLE_FINAL Or type_controle = DEGRADE Then
            inter$ = ""
            'inter$ = inter$ & "Produit alimenté en " & Format$(u_alim(0)) & " V" & forme$ & " :" & NL$
            inter$ = inter$ & "Product supply " & Format$(u_alim(0)) & " V" & forme$ & " :" & NL$
            
            'inter$ = inter$ & " " & portee_mini & " < Portee (mm) < " & portee_maxi & NL$
            'inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
            inter$ = inter$ & " " & portee_mini & " < SR (mm) < " & portee_maxi & NL$
            inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
            
            ' is produit autoadaptatif affiche distances suivantes
            If si_ctrl_auto_adaptatif Then
                'inter$ = inter$ & " " & "Portée minimal = " & portee_mini_auto & " mm" & NL$
                'inter$ = inter$ & " " & "Portée nominal = " & portee_nominal & " mm" & NL$
                inter$ = inter$ & " " & "SR minimal = " & portee_mini_auto & " mm" & NL$
                inter$ = inter$ & " " & "SR nominal = " & portee_nominal & " mm" & NL$
            End If
                        
            'inter$ = inter$ & " " & ud_min_u1(POS_TAB_FINAL) & " V < Ud < " & ud_max_u2(POS_TAB_FINAL) & " V" & NL$
            inter$ = inter$ & " " & ud_min_u1(POS_TAB_FINAL) & " V < Vd < " & ud_max_u2(POS_TAB_FINAL) & " V" & NL$
            If if_min_u1(POS_TAB_FINAL) = 0 And if_max_u1(POS_TAB_FINAL) = 0 Then
                'inter$ = inter$ & " " & ic_act_min_u1(POS_TAB_FINAL) & "mA < Ic saturée < " & ic_act_max_u1(POS_TAB_FINAL) & " mA" & NL$
                'inter$ = inter$ & " " & ic_repos_min_u1(POS_TAB_FINAL) & "mA < Ic bloquée < " & ic_repos_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_act_min_u1(POS_TAB_FINAL) & "mA < Ic detect < " & ic_act_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_repos_min_u1(POS_TAB_FINAL) & "mA < Ic no detect < " & ic_repos_max_u1(POS_TAB_FINAL) & " mA" & NL$
            Else
                'inter$ = inter$ & " " & if_min_u1(POS_TAB_FINAL) & "mA < I consommé < " & if_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & if_min_u1(POS_TAB_FINAL) & "mA < I con < " & if_max_u1(POS_TAB_FINAL) & " mA" & NL$
            End If
            
            If (u_alim(0) > u_alim(1)) And protege_cc = True Then
                'inter$ = inter$ & "Test court-circuit" & NL$
                inter$ = inter$ & "Test short-circuit" & NL$
            End If
            
            
            inter$ = inter$ & NL$
        
        If (u_alim(1) <> 0) Then
        
            'inter$ = inter$ & "Produit alimenté en " & Format$(u_alim(1)) & " V" & forme$ & " :" & NL$
            inter$ = inter$ & "Product supply " & Format$(u_alim(1)) & " V" & forme$ & " :" & NL$
            
            If ctrl_simplifie_final = 0 Then
                'inter$ = inter$ & " " & portee_mini & " < Portee (mm) < " & portee_maxi & NL$
                'inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
                inter$ = inter$ & " " & portee_mini & " < SR (mm) < " & portee_maxi & NL$
                inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
            Else
                'inter$ = inter$ & "Controle simplifie " & NL$
                inter$ = inter$ & "Simple Test " & NL$
            End If
            
            'inter$ = inter$ & " " & ud_min_u2(POS_TAB_FINAL) & " V < Ud < " & ud_max_u2(POS_TAB_FINAL) & " V" & NL$
            inter$ = inter$ & " " & ud_min_u2(POS_TAB_FINAL) & " V < Vd < " & ud_max_u2(POS_TAB_FINAL) & " V" & NL$
            If if_min_u2(POS_TAB_FINAL) = 0 And if_max_u2(POS_TAB_FINAL) = 0 Then
                'inter$ = inter$ & " " & ic_act_min_u2(POS_TAB_FINAL) & "mA < Ic saturée < " & ic_act_max_u2(POS_TAB_FINAL) & " mA" & NL$
                'inter$ = inter$ & " " & ic_repos_min_u2(POS_TAB_FINAL) & "mA < Ic bloquée < " & ic_repos_max_u2(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_act_min_u2(POS_TAB_FINAL) & "mA < Ic detect < " & ic_act_max_u2(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_repos_min_u2(POS_TAB_FINAL) & "mA < Ic no detect < " & ic_repos_max_u2(POS_TAB_FINAL) & " mA" & NL$
            Else
                'inter$ = inter$ & " " & if_min_u1(POS_TAB_FINAL) & "mA < I consommé < " & if_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & if_min_u1(POS_TAB_FINAL) & "mA < I con < " & if_max_u1(POS_TAB_FINAL) & " mA" & NL$
            End If
            
            If (u_alim(0) < u_alim(1)) And protege_cc = True Then
                'inter$ = inter$ & "Test court-circuit" & NL$
                inter$ = inter$ & "Test short-circuit" & NL$
            End If
            
        End If
     End If

     If type_controle = TEMPERATURE_HAUTE Or type_controle = TEMPERATURE_BASSE Then
     
            inter$ = ""
            'inter$ = inter$ & "Produit alimenté en " & u_alim_T(0) & " V" & forme$ & " :" & NL$
            'inter$ = inter$ & " " & portee_mini_T & " < Portee (mm) < " & portee_maxi_T & NL$
            'inter$ = inter$ & " " & hysteresis_mini_T & " < Hyst.(%) < " & hysteresis_maxi_T & NL$
            inter$ = inter$ & "Product supply " & u_alim_T(0) & " V" & forme$ & " :" & NL$
            inter$ = inter$ & " " & portee_mini_T & " < SR (mm) < " & portee_maxi_T & NL$
            inter$ = inter$ & " " & hysteresis_mini_T & " < Hyst.(%) < " & hysteresis_maxi_T & NL$
            
            ' is produit autoadaptatif affiche distances suivantes
            If si_ctrl_auto_adaptatif Then
                'inter$ = inter$ & " " & "Portée minimal = " & portee_mini_auto & " mm" & NL$
                'inter$ = inter$ & " " & "Portée nominal = " & portee_nominal & " mm" & NL$
                inter$ = inter$ & " " & "SR minimal = " & portee_mini_auto & " mm" & NL$
                inter$ = inter$ & " " & "SR nominal = " & portee_nominal & " mm" & NL$
            End If

            inter$ = inter$ & NL$
        
        If u_alim_T(1) <> 0 And Not si_ctrl_auto_adaptatif Then
                'inter$ = inter$ & "Produit alimenté en " & u_alim_T(1) & " V" & forme$ & " :" & NL$
                inter$ = inter$ & "Product supply " & u_alim_T(1) & " V" & forme$ & " :" & NL$
            If Not si_ctrl_auto_adaptatif Then
                If ctrl_simplifie_T = 0 Then
                    'inter$ = inter$ & " " & portee_mini_T & " < Portee (mm) < " & portee_maxi_T & NL$
                    'inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
                    inter$ = inter$ & " " & portee_mini_T & " < SR (mm) < " & portee_maxi_T & NL$
                    inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
                Else
                    'inter$ = inter$ & "Controle simplifie "
                    inter$ = inter$ & "Simple test "
                End If
            End If
    End If
    End If

    If type_controle = PREMIER_CONTROLE Then ' premier contrôl
        inter$ = ""
        'inter$ = inter$ & "Produit alimenté en " & u_alim_T(0) & " V" & forme$ & " :" & NL$
        'inter$ = inter$ & " " & "Portée nominal = " & portee_nominal & " mm" & NL$
        inter$ = inter$ & "Product supply " & u_alim_T(0) & " V" & forme$ & " :" & NL$
        inter$ = inter$ & " " & "SR nominal = " & portee_nominal & " mm" & NL$
        inter$ = inter$ & NL$
    End If

'-> SODATEC 9/5/1
End If
'Remplissage de la boite texte où figure les limites
'de la Base de Donnees
'mode analogique
If si_ctrl_analogique Then

   inter$ = Message_affichage_limite_prod_ana(type_controle)

End If 'analogique
'<- SODATEC
        'Si contole diélectrique à modifier
    If metallique And (type_controle <> 1 And type_controle <> 3) Then
        inter$ = inter$ & NL$
        'inter$ = inter$ & "Contrôle diélectrique" & NL$
        'inter$ = inter$ & "Claquage à :" & Ufinal & "V~" & " avec If < " & I_fuite & "mA" & NL$
        inter$ = inter$ & "HIPOT TEST" & NL$
        inter$ = inter$ & "HIPOT at :" & Ufinal & "V~" & " With I < " & I_fuite & "mA" & NL$
       If fil_masse Then
            'inter$ = inter$ & "Continuité à : 5A avec R < 0,5 Ohms" & NL$
            inter$ = inter$ & "Continuity at : 5A with R < 0,5 Ohms" & NL$
       End If
    End If

        
        FRM_controle!Affiche1.Clear
        FRM_controle!Affiche1.Gras = True
        FRM_controle!Affiche1.Bare = True
        FRM_controle!Affiche1.Add_text inter$


Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_text_limites. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_text_limites. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Sub maj_text_param_temp()
Dim NL$
Dim inter$
Dim forme$
Dim text_courant As String

'On Error GoTo 400

' AC ou DC ?
If ac_temp Then forme$ = "~" Else forme$ = "="

'Remplissage de la boite texte où figure les limites
'de la Base de Donnees

'-> SODATEC 9/5/1
'mode non analogique
If si_ctrl_tor Then
'<- SODATEC
        NL$ = Chr$(13) & Chr$(10)
        'inter$ = inter$ & "Produit alimenté en " & u_alim_T(0) & " V" & forme$ & " :" & NL$
        inter$ = inter$ & "Product supply " & u_alim_T(0) & " V" & forme$ & " :" & NL$
        
            'inter$ = inter$ & " " & portee_mini_T & " < Portee (mm) < " & portee_maxi_T & NL$
            'inter$ = inter$ & " " & hysteresis_mini_T & " < Hyst.(%) < " & hysteresis_maxi_T & NL$
            inter$ = inter$ & " " & portee_mini_T & " < SR (mm) < " & portee_maxi_T & NL$
            inter$ = inter$ & " " & hysteresis_mini_T & " < Hyst.(%) < " & hysteresis_maxi_T & NL$
            ' si produit autoadaptatif affiche distances suivantes
            If si_ctrl_auto_adaptatif Then
                'inter$ = inter$ & " " & "Portée minimal = " & portee_mini_auto_T & " mm" & NL$
                inter$ = inter$ & " " & "SR nominal = " & portee_nominal_T & " mm" & NL$
            End If
        inter$ = inter$ & NL$
        
        If u_alim_T(1) <> 0 Then
            'inter$ = inter$ & "Produit alimenté en " & u_alim_T(1) & " V" & forme$ & " :" & NL$
            inter$ = inter$ & "Product supply " & u_alim_T(1) & " V" & forme$ & " :" & NL$
            If Not si_ctrl_auto_adaptatif Then
                If ctrl_simplifie_T = 0 Then
                    'inter$ = inter$ & " " & portee_mini_T & " < Portee (mm) < " & portee_maxi_T & NL$
                    'inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
                    inter$ = inter$ & " " & portee_mini_T & " < SR (mm) < " & portee_maxi_T & NL$
                    inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
                Else
                    'inter$ = inter$ & "Controle simplifie "
                    inter$ = inter$ & "Simple test "
                End If
            End If
        End If
        
'-> SODATEC 9/5/1
End If
'Remplissage de la boite texte où figure les limites
'de la Base de Donnees
'mode analogique
If si_ctrl_analogique Then
    
    inter$ = inter$ & Message_affichage_limite_prod_ana(TEMPERATURE)
 
End If 'analogique
'<- SODATEC
FRM_PARAM!TxT_aff_temp.Text = inter$

Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_temp. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_temp. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub


Sub maj_text_param_final()
Dim NL$
Dim inter$
Dim forme$
Dim text_courant As String
'On Error GoTo 400

' AC ou DC ?
If ac_final Then forme$ = "~" Else forme$ = "="

'Remplissage de la boite texte où figure les limites
'de la Base de Donné

'-> SODATEC 9/5/1
'mode non analogique
If si_ctrl_tor Then
'<- SODATEC
        
        type_controle = 0
        
        NL$ = Chr$(13) & Chr$(10)
        'inter$ = inter$ & "Produit alimenté en " & Format$(u_alim_F(0)) & " V" & forme$ & " :" & NL$
        inter$ = inter$ & "Product supply " & Format$(u_alim_F(0)) & " V" & forme$ & " :" & NL$
            
            'inter$ = inter$ & " " & portee_mini & " < Portee (mm) < " & portee_maxi & NL$
            'inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
            inter$ = inter$ & " " & portee_mini & " < SR (mm) < " & portee_maxi & NL$
            inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
            
            ' si produit autoadaptatif affiche distances suivantes
            If si_ctrl_auto_adaptatif Then
                'inter$ = inter$ & " " & "Portée minimal = " & portee_mini_auto & " mm" & NL$
                'inter$ = inter$ & " " & "Portée nominal = " & portee_nominal & " mm" & NL$
                inter$ = inter$ & " " & "SR minimal = " & portee_mini_auto & " mm" & NL$
                inter$ = inter$ & " " & "SR nominal = " & portee_nominal & " mm" & NL$
            End If
            
            'inter$ = inter$ & " " & ud_min_u1(POS_TAB_FINAL) & " V < Ud < " & ud_max_u1(POS_TAB_FINAL) & " V" & NL$
            inter$ = inter$ & " " & ud_min_u1(POS_TAB_FINAL) & " V < Vd < " & ud_max_u1(POS_TAB_FINAL) & " V" & NL$
            If ctrl_3_4_fils Then
                'inter$ = inter$ & " " & ic_act_min_u1(POS_TAB_FINAL) & "mA < Ic saturée < " & ic_act_max_u1(POS_TAB_FINAL) & " mA" & NL$
                'inter$ = inter$ & " " & ic_repos_min_u1(POS_TAB_FINAL) & "mA < Ic bloquée < " & ic_repos_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_act_min_u1(POS_TAB_FINAL) & "mA < Ic detect < " & ic_act_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_repos_min_u1(POS_TAB_FINAL) & "mA < Ic no detect < " & ic_repos_max_u1(POS_TAB_FINAL) & " mA" & NL$
            Else
                'inter$ = inter$ & " " & if_min_u1(POS_TAB_FINAL) & "mA < I consommé < " & if_max_u1(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & if_min_u1(POS_TAB_FINAL) & "mA < I con < " & if_max_u1(POS_TAB_FINAL) & " mA" & NL$
            End If
            
            If (u_alim_F(0) > u_alim_F(1)) And protégé_cc_final = True Then
                'inter$ = inter$ & "Test court-circuit"
                inter$ = inter$ & "Test short-circuit"
            End If
            
            inter$ = inter$ & NL$
            
        If u_alim_F(1) <> 0 Then
        
            'inter$ = inter$ & "Produit alimenté en " & Format$(u_alim_F(1)) & " V" & forme$ & " :" & NL$
            inter$ = inter$ & "Product supply " & Format$(u_alim_F(1)) & " V" & forme$ & " :" & NL$
            
            If ctrl_simplifie_final = 0 Then
                'inter$ = inter$ & " " & portee_mini & " < Portee (mm) < " & portee_maxi & NL$
                'inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
                inter$ = inter$ & " " & portee_mini & " < SR (mm) < " & portee_maxi & NL$
                inter$ = inter$ & " " & hysteresis_mini & " < Hyst.(%) < " & hysteresis_maxi & NL$
            Else
                If Not si_ctrl_auto_adaptatif Then
                    'inter$ = inter$ & "Controle simplifie " & NL$
                    inter$ = inter$ & "Simple test " & NL$
                End If
            End If
            
            'inter$ = inter$ & " " & ud_min_u2(POS_TAB_FINAL) & " V < Ud < " & ud_max_u2(POS_TAB_FINAL) & " V" & NL$
            inter$ = inter$ & " " & ud_min_u2(POS_TAB_FINAL) & " V < Vd < " & ud_max_u2(POS_TAB_FINAL) & " V" & NL$
            If ctrl_3_4_fils Then
                'inter$ = inter$ & " " & ic_act_min_u2(POS_TAB_FINAL) & "mA < Ic saturée < " & ic_act_max_u2(POS_TAB_FINAL) & " mA" & NL$
                'inter$ = inter$ & " " & ic_repos_min_u2(POS_TAB_FINAL) & "mA < Ic bloquée < " & ic_repos_max_u2(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_act_min_u2(POS_TAB_FINAL) & "mA < Ic detect < " & ic_act_max_u2(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & ic_repos_min_u2(POS_TAB_FINAL) & "mA < Ic no detect < " & ic_repos_max_u2(POS_TAB_FINAL) & " mA" & NL$
            Else
                'inter$ = inter$ & " " & if_min_u2(POS_TAB_FINAL) & "mA < I consommé < " & if_max_u2(POS_TAB_FINAL) & " mA" & NL$
                inter$ = inter$ & " " & if_min_u2(POS_TAB_FINAL) & "mA < I con < " & if_max_u2(POS_TAB_FINAL) & " mA" & NL$
            End If
            
            If (u_alim_F(0) < u_alim_F(1)) And protégé_cc_final = True Then
                'inter$ = inter$ & "Test court-circuit"
                inter$ = inter$ & "Test short-circuit"
            End If
        
        End If

'-> SODATEC 9/5/1
End If ' sortie tor
'Remplissage de la boite texte où figure les limites
'de la Base de Donnees
'mode analogique
If si_ctrl_analogique Then
 
    inter$ = inter$ & Message_affichage_limite_prod_ana(CONTROLE_FINAL)
     
End If 'analogique
'<- SODATEC
FRM_PARAM!TxT_aff_final.Text = inter$


Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_final. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_final. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Public Function Message_affichage_limite_prod_ana(Type_ctrl As Integer) As String

Dim NL$
Dim inter$
Dim forme$
Dim text_courant As String

Dim U1 As Double
Dim U2 As Double
Dim Portée_sub As Double
Dim Charge_sub As Double
Dim Contrôle_CC As Boolean
Dim Contrôle_simplifié As Boolean

Dim Pos_tab As Integer

NL$ = Chr$(13) & Chr$(10)
     text_courant = "??"
     If si_ctrl_analogique_IS Then
        text_courant = "Is"
     End If
     If si_ctrl_analogique_IS_IC Then
        text_courant = "Is+Ic"
     End If
     
     If Type_ctrl = TEMPERATURE Or Type_ctrl = TEMPERATURE_HAUTE Or Type_ctrl = TEMPERATURE_BASSE Then
        U1 = u_alim_T(0)
        U2 = u_alim_T(1)
        Pos_tab = POS_TAB_TEMPERATURE
        Portée_sub = portee_maxi_T
        Charge_sub = courant_charge_T
        Contrôle_CC = protégé_cc_T
        Contrôle_simplifié = ctrl_simplifie_T
     Else
        U1 = u_alim_F(0)
        U2 = u_alim_F(1)
        Pos_tab = POS_TAB_FINAL
        Portée_sub = portee_maxi
        Charge_sub = courant_charge
        Contrôle_CC = protégé_cc_final
        Contrôle_simplifié = ctrl_simplifie_final
     End If
     
     NL$ = Chr$(13) & Chr$(10)
            
    'inter$ = inter$ & "Produit alimenté en " & Format$(U1) & " V" & forme$ & " :" & NL$
    inter$ = inter$ & "Product supply " & Format$(U1) & " V" & forme$ & " :" & NL$
    
    'inter$ = inter$ & " Portee (mm) = " & Portée_sub & NL$
    'inter$ = inter$ & "Charge (homs) = " & Charge_sub & NL$
    'inter$ = inter$ & " " & ud_min_u1(Pos_tab) & " mA < Ic < " & ud_max_u1(Pos_tab) & " mA" & NL$
    'inter$ = inter$ & " " & ic_act_min_u1(Pos_tab) & "mA < " & text_courant & " à Sn =0 < " & ic_act_max_u1(Pos_tab) & " mA" & NL$
    'inter$ = inter$ & " " & ic_repos_min_u1(Pos_tab) & "mA < " & text_courant & " à Sn/2 < " & ic_repos_max_u1(Pos_tab) & " mA" & NL$
    'inter$ = inter$ & " " & if_min_u1(Pos_tab) & "mA < " & text_courant & " à Sn < " & if_max_u1(Pos_tab) & " mA" & NL$
    inter$ = inter$ & " SR (mm) = " & Portée_sub & NL$
    inter$ = inter$ & " Load (oms) = " & Charge_sub & NL$
    inter$ = inter$ & " " & ud_min_u1(Pos_tab) & " mA < Ic < " & ud_max_u1(Pos_tab) & " mA" & NL$
    inter$ = inter$ & " " & ic_act_min_u1(Pos_tab) & "mA < " & text_courant & " at Sn=0 < " & ic_act_max_u1(Pos_tab) & " mA" & NL$
    inter$ = inter$ & " " & ic_repos_min_u1(Pos_tab) & "mA < " & text_courant & " at Sn/2 < " & ic_repos_max_u1(Pos_tab) & " mA" & NL$
    inter$ = inter$ & " " & if_min_u1(Pos_tab) & "mA < " & text_courant & " at Sn < " & if_max_u1(Pos_tab) & " mA" & NL$
    
    
    If (U1 > U2) And Contrôle_CC = True Then
        'inter$ = inter$ & "Test court-circuit"
        inter$ = inter$ & "Test short-circuit"
    End If
    inter$ = inter$ & NL$
    
If U2 <> 0 Then
    'inter$ = inter$ & "Produit alimenté en " & Format$(U2) & " V" & forme$ & " :" & NL$
    inter$ = inter$ & "Product supply " & Format$(U2) & " V" & forme$ & " :" & NL$
    
    If ctrl_simplifie = 0 Then
        'inter$ = inter$ & " " & ud_min_u2(Pos_tab) & " mA < Ic < " & ud_max_u2(Pos_tab) & " mA" & NL$
        'inter$ = inter$ & " " & ic_act_min_u2(Pos_tab) & "mA < " & text_courant & " à Sn=0 < " & ic_act_max_u2(Pos_tab) & " mA" & NL$
        'inter$ = inter$ & " " & ic_repos_min_u2(Pos_tab) & "mA < " & text_courant & " à Sn/2 < " & ic_repos_max_u2(Pos_tab) & " mA" & NL$
        inter$ = inter$ & " " & ud_min_u2(Pos_tab) & " mA < Ic < " & ud_max_u2(Pos_tab) & " mA" & NL$
        inter$ = inter$ & " " & ic_act_min_u2(Pos_tab) & "mA < " & text_courant & " at Sn=0 < " & ic_act_max_u2(Pos_tab) & " mA" & NL$
        inter$ = inter$ & " " & ic_repos_min_u2(Pos_tab) & "mA < " & text_courant & " at Sn/2 < " & ic_repos_max_u2(Pos_tab) & " mA" & NL$
        'inter$ = inter$ & " Portee (mm) = " & portee_maxi & NL$
    Else
        'inter$ = inter$ & "Controle simplifie " & NL$
        inter$ = inter$ & "Simple test " & NL$
    End If
    
    'inter$ = inter$ & " " & if_min_u2(Pos_tab) & "mA < " & text_courant & " à Sn < " & if_max_u2(Pos_tab) & " mA" & NL$
    inter$ = inter$ & " " & if_min_u2(Pos_tab) & "mA < " & text_courant & " at Sn < " & if_max_u2(Pos_tab) & " mA" & NL$
    
    If (U1 < U2) And Contrôle_CC = True Then
        'inter$ = inter$ & "Test court-circuit"
        inter$ = inter$ & "Test short-circuit"
    End If
End If

Message_affichage_limite_prod_ana = inter$
End Function

Sub maj_text_param_géné()
Dim NL$
Dim inter$
Dim forme$

'On Error GoTo 400

'Remplissage de la boite texte où figure les limites
'de la Base de Donnees
        NL$ = Chr$(13) & Chr$(10)
        'inter$ = "SYMBOLE COMMERCIAL : " & symb_com & NL$
        'inter$ = inter$ & "Référence industrielle : " & Ref_indus & NL$
        'inter$ = inter$ & "Numéro de lot : " & Num_lot & NL$
        inter$ = "COMMERCIAL SYMBOLE : " & symb_com & NL$
        inter$ = inter$ & "INDUSTRIAL SYMBOLE : " & Ref_indus & NL$
        inter$ = inter$ & "Lot Number : " & Num_lot & NL$
        
        FRM_PARAM!TxT_aff_géné.Text = inter$

Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_géné. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_géné. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Sub maj_text_param_claq()
Dim NL$
Dim inter$
Dim forme$

'On Error GoTo 400

'Remplissage de la boite texte où figure les limites
'de la Base de Donnees
        NL$ = Chr$(13) & Chr$(10)
        
    If metallique Then
        'inter$ = inter$ & "Claquage à :" & Ufinal & "V~" & " avec If < " & I_fuite & "mA" & NL$
        inter$ = inter$ & "HIPOT at :" & Ufinal & "V~" & " with I < " & I_fuite & "mA" & NL$
       If fil_masse Then
            'inter$ = inter$ & "Continuité à : 5A avec R < 0,5 Ohms" & NL$
            inter$ = inter$ & "Continuity at : 5A with R < 0,5 Ohms" & NL$
       End If
    End If
    
        FRM_PARAM!TxT_aff_claq.Text = inter$

Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_claq. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module maj_text_param_claq. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Sub prepa_params(Donnée() As Long)
    Dim octed_met As Integer
    Dim Decalage_base As Long
    Dim Tmp As Integer
    
On Error GoTo 400
    
    'lecture de la réference commercial et industrielle ainsi que le numéro de lot
    'est assurée par l'evenement "DataArrival" de la socket de communication
    'voir "Winsock1_DataArrival" dans la FRM "Utwmait"
    
    empreinte_miniature = (Donnée(2) = 13360) Or (Donnée(2) = 13688) Or (Donnée(2) = 13616)
    
    ' Si contrôle température on va chercher les paramétres entre WM468 si non WW524
    Decalage_base = 104 ' IIf(type_controle = TEMPERATURE_HAUTE Or type_controle = TEMPERATURE_BASSE, 48, 104)
    
    'Nombre de produit à tester
    Nombre_produit_à_tester = Donnée(70)
    
    ' detection produit auto adaptatif
    If (Donnée(Decalage_base) And (&HF0)) = &HD0 Then
        si_ctrl_auto_adaptatif = True
    Else
        si_ctrl_auto_adaptatif = False
    End If
    
    Code_produit_badge = (Donnée(Decalage_base) And &HFF) 'garde premier et le deuxième  quarté
                                                    'identifie les cycle de contrôle à télécharger dans l'API
    ' détection produit analogique
    If (Donnée(Decalage_base) And (&HF0)) = &HC0 Then
        si_ctrl_analogique = True
        Select Case (Donnée(Decalage_base) And &HF) 'garde premier quarté
            Case 0
                ' cas C0 Is et Ic
                si_ctrl_analogique_IS = True
                si_ctrl_analogique_IS_IC = False
            Case 1
                'cas C1 Ic et Is+Ic
                si_ctrl_analogique_IS = False
                si_ctrl_analogique_IS_IC = True
        End Select
    Else
        si_ctrl_analogique = False
        si_ctrl_analogique_IS = False
        si_ctrl_analogique_IS_IC = False
    End If
   si_ctrl_tor = Not (si_ctrl_analogique)
    
    'Converti chaine pour nouvelle codification produit
    Select Case (Donnée(Decalage_base) And &HFF)
        Case &HD1
            Tmp = &H16
        Case &HD0
            Tmp = &H1A
        Case &HD2
            Tmp = &H2A
        Case &HD3
            Tmp = &H26
        Case Else
            Tmp = (Donnée(Decalage_base) And &HFF)
    End Select
    
    Select Case Tmp
        Case &H9, &H1A, &H2A, &HD0, &HD2 ' produit NO
            type_commutation = PRODUIT_NO
        Case &H5, &H16, &H26, &HD1, &HD3 ' produit NC
            type_commutation = PRODUIT_NC
        Case &H1F, &H2F ' produit NO + NC
            type_commutation = PRODUIT_NO_NC
        Case &H40, &H43 ' produit prgrammable
            type_commutation = PRODUIT_PROGRAMMABLE
    End Select
    
    
    'affectation des paramétres type branchement (2,3,4,.... fils)
    Select Case Tmp
        Case &H1A, &H16, &H2A, &H26, &H1F, &H2F, &H40, &H43
            ctrl_3_4_fils = True
       Case Else
            ctrl_3_4_fils = False
    End Select
    
    npn = (Donnée(Decalage_base) And BIT5) / BIT5
    
    'empreinte_miniature = (Donnée(Decalage_base + 1) And BIT12) / BIT12
    
    ' Si pas ontrôle température on va chercher les paramétres en WM452 à WM466
    
    Tmp = 0
    For Tmp = POS_TAB_FINAL To POS_TAB_TEMPERATURE
        Decalage_base = IIf(Tmp <> 0, 32, 88)
        ic_act_min_u1(Tmp) = Donnée(Decalage_base) / 10
        ic_act_max_u1(Tmp) = Donnée(Decalage_base + 1) / 10
        ic_act_min_u2(Tmp) = Donnée(Decalage_base + 2) / 10
        ic_act_max_u2(Tmp) = Donnée(Decalage_base + 3) / 10
        
        ic_repos_min_u1(Tmp) = Donnée(Decalage_base + 4) / 10
        ic_repos_max_u1(Tmp) = Donnée(Decalage_base + 5) / 10
        ic_repos_min_u2(Tmp) = Donnée(Decalage_base + 6) / 10
        ic_repos_max_u2(Tmp) = Donnée(Decalage_base + 7) / 10
        
        if_min_u1(Tmp) = Donnée(Decalage_base + 8) / 10
        if_max_u1(Tmp) = Donnée(Decalage_base + 9) / 10
        if_min_u2(Tmp) = Donnée(Decalage_base + 10) / 10
        if_max_u2(Tmp) = Donnée(Decalage_base + 11) / 10
        
        ud_min_u1(Tmp) = Donnée(Decalage_base + 12) / 10
        ud_max_u1(Tmp) = Donnée(Decalage_base + 13) / 10
        ud_min_u2(Tmp) = Donnée(Decalage_base + 14) / 10
        ud_max_u2(Tmp) = Donnée(Decalage_base + 15) / 10
    Next Tmp
        
    ' affectation des parametres controle final ou premier contrôle
    ac_final = (Donnée(104) And BIT9) / BIT9
    If si_ctrl_auto_adaptatif Then
        portee_mini = CDbl(Donnée(81)) * 0.009
        portee_maxi = CDbl(Donnée(81)) * 0.011
        portee_nominal = CDbl(Donnée(81)) * 0.01
        portee_mini_auto = CDbl(Donnée(82)) * 0.01
    Else
        portee_mini = Donnée(82) / 100
        portee_maxi = Donnée(81) / 100
    End If
    hysteresis_mini = Donnée(84)
    hysteresis_maxi = Donnée(83)
    
    u_alim_F(0) = Donnée(85) / 10
    u_alim_F(1) = Donnée(86) / 10
    
    courant_charge = Donnée(87)

    protégé_cc_final = (Donnée(104) And BIT10) / BIT10
    ctrl_simplifie_final = (Donnée(104) And BIT12) / BIT12

    ' affectation des parametres controle température
    If si_ctrl_auto_adaptatif Then
        portee_mini_T = CDbl(Donnée(21)) * 0.008
        portee_maxi_T = CDbl(Donnée(21)) * 0.012
        portee_nominal_T = CDbl(Donnée(21)) * 0.01
        portee_mini_auto_T = CDbl(Donnée(22)) * 0.01
    Else
        portee_mini_T = Donnée(22) / 100
        portee_maxi_T = Donnée(21) / 100
    End If
    
    hysteresis_mini_T = Donnée(26)
    hysteresis_maxi_T = Donnée(25)
    ac_temp = (Donnée(48) And BIT9) / BIT9
    courant_charge_T = Donnée(31)
    u_alim_T(0) = Donnée(29) / 10
    u_alim_T(1) = Donnée(30) / 10
    
    protégé_cc_T = (Donnée(48) And BIT10) / BIT10
    ctrl_simplifie_T = (Donnée(48) And BIT12) / BIT12
           
    ' affectation des parametres controle diélectrique
    octed_met = (Donnée(1) And (BIT0 + BIT1 + BIT2 + BIT3 + BIT4 + BIT5 + BIT6 + BIT7))
        
    If octed_met = 49 Or octed_met = 50 Or octed_met = 53 Or octed_met = 54 Then
        metallique = True 'Produits XS1,XS2,XS5,XS6'
    Else
        metallique = False ' Produits XS3,XS4'
    End If
        
    fil_masse = (Donnée(71) And BIT0) / BIT0
    Ufinal = Donnée(60)
    I_fuite = Donnée(66) / 10

    'affectation des valeurs en fonction du type de controle
    If type_controle = TEMPERATURE Then ' contrôle température
        u_alim(0) = u_alim_T(0)
        u_alim(1) = u_alim_T(1)
        ac = ac_temp
        protege_cc = protégé_cc_T
        ctrl_simplifie = ctrl_simplifie_T
        
    Else 'controle final ou premier controle
        u_alim(0) = u_alim_F(0)
        u_alim(1) = u_alim_F(1)
        ac = ac_final
        protege_cc = protégé_cc_final
        ctrl_simplifie = ctrl_simplifie_final
        
    End If
    
Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module prepa_param. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module prepa_param. Laporkan ke maintenance segera. Thanks."
Resume Next
    
End Sub


Sub reset_bit(var As Long, n As Integer)
'Procedure de mise à 0 d'un bit
Dim mask

    mask = 2 ^ n
    mask = mask Xor -1
    var = var And mask
    
End Sub


Sub retour_depart()

Unload FRM_AUTOTEST
FRM_startup.Show

End Sub

Sub set_bit(var As Long, n As Integer)
'procedure de mise à 1 d'un bit
Dim mask

    mask = 2 ^ n
    var = var Or mask

End Sub


Public Sub message_defaut_general()

'com_err_tsx : numero du défaut a afficher en général

Dim chaine$, ret$, type_def

Dim i As Integer
Dim Copy_Tab(20) As Long

Dim Com_erreure As Integer

'On Error GoTo 400

chaine$ = ""
ret$ = Chr(13) & Chr(10) 'retour a la ligne, a inserer pour message trop long
    
    
Select Case com_err_tsx

    Case 0
        chaine$ = ""
        If (previous_com_err_tsx = 8) Or (previous_com_err_tsx = 9) Then
            Unload FRM_ARU
            Unload FRM_Reprise
        End If
        GoTo Fin_trait_def
    
        ' Reprise après ARU traité en timer_def
    Case 8
        If previous_com_err_tsx = 9 Then
            Unload FRM_ARU
        End If
        Load FRM_Reprise
        FRM_Reprise.Show
        If Contexte$ = "CONTROLE" Then
            out_controle = 1
        End If
        If Contexte$ = "AUTOTEST" Then
            out_autotest = 1
        End If
        GoTo Trait_defaut
    ' Défaut ARU traité en timer_def
    Case 9
        Load FRM_ARU
        FRM_ARU.Show
        If Contexte$ = "CONTROLE" Then
            out_controle = 1
        End If
        If Contexte$ = "AUTOTEST" Then
            out_autotest = 1
        End If
        GoTo Trait_defaut
            
    Case Else ' on va trouve le text du message à afficher dans la basse de donnée "Deffaut_general"
            DataEnvironment1.rsTable_defaut_general.Filter = "Code = '" & com_err_tsx & "'"
            If DataEnvironment1.rsTable_defaut_general.State = adStateClosed Then
                DataEnvironment1.rsTable_defaut_general.Open
            End If
            If DataEnvironment1.rsTable_defaut_general.RecordCount > 0 Then
                type_def = 1 ' pas d'affichage message par défaut
                If DataEnvironment1.rsTable_defaut_general.Fields("Affiche").Value Then
                    
                    type_def = 0 ' affichage message autorisé
                    chaine$ = DataEnvironment1.rsTable_defaut_general.Fields("Labelle").Value
            
                End If
            Else
                chaine$ = "DEFAUT AUTOMATE NUMERO " & Str$(com_err_tsx) & ret$ & "NON IDENTIFIE"
                type_def = 0
            End If

End Select

'TRAITEMENT DEFAUT
Trait_defaut:

'Defaut ARU ou reprise : on s'en va...
        If (com_err_tsx = 8) Or (com_err_tsx = 9) Then
            previous_com_err_tsx = com_err_tsx
            GoTo Fin_trait_def
        End If

'En contrôle, on affiche une MSGBOX si le défaut est 1 nouveau défaut poste, sinon on passe...
If (Contexte$ = "CONTROLE") And (com_err_tsx <> 0) Then
    If type_def = 0 Then
        'MsgBox chaine$, 48, "DEFAUT POSTE : validez pour acquiter"
        MsgBox chaine$, 48, "POST FAULT : Confirm to quit"
    Else
        If (previous_com_err_tsx <> com_err_tsx) Then
            'chaine$ = "DEFAUT PRODUIT : " + ret$ + chaine$
            chaine$ = "FAULT PRODUCT : " + ret$ + chaine$
            GoTo Acquit_defaut
        End If
    End If
End If

'Si on n'est pas en contrôle, on affiche une MSGBOX si il y a un defaut
If (Contexte$ <> "CONTROLE") And (com_err_tsx <> 0) Then
    If Contexte$ = "AUTOTEST" Then
        'MsgBox chaine$, 64, "ABANDON AUTOTEST : validez pour revenir au menu précédent"
        MsgBox chaine$, 64, "CANCEL AUTOTESS : Confirm to return to last menu"
    Else
        'MsgBox chaine$, 48, "DEFAUT : validez pour acquiter"
        MsgBox chaine$, 48, "FAULT : Confirm to quit"
    End If
End If

'on fait l'acquis du défaut s'il existe
Acquit_defaut:
If com_err_tsx <> 0 Then
    'on refait une lecture des infos du tsx avant de faire l'acquis defaut, car des choses se sont produites entre temps...
    Lecture_etat_banc
    
    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
    
    com_err_tsx = 0
    Copy_Tab(4) = Copy_Tab(4) And &HFFC0
    'Copy_Tab(7) = Copy_Tab(4)
    set_bit Copy_Tab(0), 14

    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 10)
        If Com_erreure <= 0 Then ' Communication en erreur
            If MsgBox(ERREUR_ECRITURE_DEFAUT, vbCritical + vbRetryCancel) = vbCancel Then
                If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                    End
                End If
            End If
        Else ' Communication sans erreur
        
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            
            'memo defaut precedent
            previous_com_err_tsx = com_err_tsx
            Exit Sub
        End If
    Loop
End If

'FIN DE TRAITEMENT DEFAUT
Fin_trait_def:



Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module message_defaut_general. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module message_defaut_general. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub

Sub message_defaut_chargement()

'def_chargement : numero du défaut a afficher

'On Error GoTo 400
    
Dim Num_produit As Integer 'Numéro du produit en défaut
Dim Titre As String 'Titre de la feunêtre affichage message
Dim Enregistrement As Long 'Numéro enregistre de la base de donnée log à imprimer

Dim i As Integer
Dim Copy_Tab(20) As Long

'TRAITEMENT DEFAUT

' on quit si il n'y a pas de défaut
If def_chargement = 0 Then
    Exit Sub
End If

'Mis en forme du titre de la fenetre d'affichage
Select Case Contexte$
    Case "CONTROLE"
        Titre = "DEFAUT POSTE CHARGEMENT: validez pour acquiter"
    Case "AUTOTEST"
        Titre = "ABANDON AUTOTEST : validez pour revenir au menu précédent"
    Case Else
        Titre = "DEFAUT : validez pour acquiter"
End Select
    
DataEnvironment1.rsTable_defaut_chargement.Filter = "Code = '" & def_chargement & "'"

If DataEnvironment1.rsTable_defaut_chargement.State = adStateClosed Then
    DataEnvironment1.rsTable_defaut_chargement.Open
End If

    If DataEnvironment1.rsTable_defaut_chargement.RecordCount > 0 Then
        If DataEnvironment1.rsTable_defaut_chargement.Fields("Affiche").Value Then
                
            Num_produit = DataEnvironment1.rsTable_defaut_chargement.Fields("Num_produit").Value
                
            Dim Affiche As New Affiche_defaut.Affiche ' Fenetre d'affichage de défaut
            Dim Reponce As Affiche_defaut.Bouton ' Reponce de la fenetre d'affichage de défaut

            Affiche.Text_Color = vbRed
            Reponce = Affiche.Affiche_message_chargement(DataEnvironment1.rsTable_defaut_chargement.Fields("Labelle"), DataEnvironment1.rsTable_defaut_chargement.Fields("Reprise_autorise"), Num_produit, Titre)
            
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Copy_Tab(i) = Tab_Status_banc(i)
            Next i
            
            If Reponce = Reprise Then
                If def_chargement <> 0 Then
                   set_bit Copy_Tab(4), 6 ' Quitement défaut chargement
                   set_bit Copy_Tab(6), 6 ' Annulation défaut reprise chargement
                   set_bit Copy_Tab(0), 7 ' Fin de demande d'affichage fin de chargement
                   def_chargement = 0
                   Copy_Tab(7) = 0
                End If

            Else ' Abondant test
                If def_chargement <> 0 Then
                   set_bit Copy_Tab(4), 6 ' Quitement défaut chargement
                   set_bit Copy_Tab(0), 7 ' Fin de demande d'affichage fin de chargement
                   def_chargement = 0
                   Copy_Tab(7) = 0
                   
                   Select Case Contexte$
                        Case "CONTROLE" 'Imprime étiquette défaut seulement en dans le contexte "CONTROLE"
                            Dim Prod_info As Info_produit
                            Dim Imprimer_etiquette As Boolean
                            
                            Prod_info.Code_comercial = symb_com
                            Prod_info.Code_lot = Num_lot
        
                            Select Case Num_produit
                                Case PRODUIT_1
                                    Prod_info.Mun_emprente = Tab_Status_banc(13)
                                    set_bit Copy_Tab(6), 4 ' Confirmation produit 1 maiuvais
                                    set_bit Copy_Tab(6), 9 ' Produit 1 maiuvais par PC
                                    set_bit Copy_Tab(0), 13 ' Demande impression etiquette par PC
                                Case PRODUIT_2
                                    Prod_info.Mun_emprente = Tab_Status_banc(14)
                                    set_bit Copy_Tab(6), 5 'Confirmation produit 2 mauvais
                                    set_bit Copy_Tab(6), 10 ' Produit 2 maiuvais par PC
                                    set_bit Copy_Tab(0), 13 ' Demande impression etiquette par PC
                                    
                                Case Else
                                     Prod_info.Mun_emprente = 0
                            End Select
                            
                            Prod_info.Num_produit = Num_produit

                            'Ajout un nouveau produit à la base de donnée
                            Ajoute_produit_log Prod_info
                            
                            DataEnvironment1.rsTable_Log.Fields("Num_phase").Value = DataEnvironment1.rsTable_defaut_chargement.Fields("Phase_Defaut").Value
                            DataEnvironment1.rsTable_Log.Fields("Num_def").Value = DataEnvironment1.rsTable_defaut_chargement.Fields("Code_Defaut").Value
                            DataEnvironment1.rsTable_Log.Fields("Num_def_etiquette").Value = DataEnvironment1.rsTable_defaut_chargement.Fields("Code_Etique").Value
                            
                            'maj labelle phase
                            DataEnvironment1.rsTable_Code_phase.Filter = "Code = '" & DataEnvironment1.rsTable_defaut_chargement.Fields("Phase_Defaut").Value & "'"
                            If DataEnvironment1.rsTable_Code_phase.State = adStateClosed Then
                                DataEnvironment1.rsTable_Code_phase.Open
                            End If
                            If DataEnvironment1.rsTable_Code_phase.RecordCount > 0 Then
                                DataEnvironment1.rsTable_Log.Fields("Labelle_phase").Value = DataEnvironment1.rsTable_Code_phase.Fields("Labelle").Value
                            Else
                                DataEnvironment1.rsTable_Log.Fields("Labelle_phase").Value = DataEnvironment1.rsTable_defaut_chargement.Fields("Phase_Defaut").Value
                            End If
                            
                            'maj labelle code défaut etiquette
                            DataEnvironment1.rsTable_Code_defaut_etiquette.Filter = "Code = '" & DataEnvironment1.rsTable_defaut_chargement.Fields("Code_Etique").Value & "'"
                            If DataEnvironment1.rsTable_Code_defaut_etiquette.State = adStateClosed Then
                                DataEnvironment1.rsTable_Code_defaut_etiquette.Open
                            End If
                            If DataEnvironment1.rsTable_Code_defaut_etiquette.RecordCount > 0 Then
                                DataEnvironment1.rsTable_Log.Fields("Labelle_def_etiquette").Value = DataEnvironment1.rsTable_Code_defaut_etiquette.Fields("Labelle").Value
                            Else
                                DataEnvironment1.rsTable_Log.Fields("Labelle_def_etiquette").Value = DataEnvironment1.rsTable_defaut_chargement.Fields("Code_Etique").Value
                            End If
                            
                            'cloture de la basse de donnée log
                            Enregistrement = Valide_donnee_log

                            Imprimer_etiquette = DataEnvironment1.rsTable_defaut_chargement.Fields("Imprimer_etiquette").Value
                            'Imprime etiquette si naissére
                            Imprimer_etiquette = Imprimer_etiquette And (Not CBool(sans_etiquette)) 'Interdi l'imprétion de l'étiquette
                            If Imprimer_etiquette Then
                                
                                If DataEnvironment1.rsrequette_update.State <> adStateClosed Then
                                    DataEnvironment1.rsrequette_update.Close 'ferme le requette avant de l'ouvir
                                End If
                                DataEnvironment1.requette_update Enregistrement
                                If DataEnvironment1.rsrequette_update.RecordCount > 0 Then
                                
                                    ' On cherche le N° enregistrement
                                    DataEnvironment1.rsrequette_update.MoveLast
                                    DataEnvironment1.rsrequette_update.Filter = "N°_enregistrement = " & DataEnvironment1.rsrequette_update.Fields("N°_enregistrement").Value
                                
                                    If Not Imprime_etiquette_ecran Then
                                        DataReportChargement.PrintReport 'Imprime  l'étiquette
                                    Else
                                        DataReportChargement.Show 'Affiche  l'étiquette à l'écran
                                    End If
                                    DataEnvironment1.rsrequette_update.Filter = ("")

                                End If
                            End If
                   End Select ' fin select contexte
                End If
            End If
        End If
    Else
        Affiche.Affiche_Message "Défaut N° " & def_chargement, "Défaut post chargement", False
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Copy_Tab(i) = Tab_Status_banc(i)
        Next i
        'Aquitement défaut
        If def_chargement <> 0 Then
           set_bit Copy_Tab(4), 6
           def_chargement = 0
           Copy_Tab(7) = 0
        End If
    End If
    
    reset_bit Copy_Tab(0), 7 ' Fin de demande d'affichage fin de chargement

    Dim Com_erreure As Integer
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 10)
        If Com_erreure <= 0 Then ' Communication en erreur
            If MsgBox(ERREUR_ECRITURE_DEFAUT, vbCritical + vbRetryCancel) = vbCancel Then
                If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                    End
                End If
            End If
        Else ' Communication sans erreur
            'memo defaut precedent
            previous_def_chargement = def_chargement
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            Exit Sub
        End If
    Loop
    
'FIN DE TRAITEMENT DEFAUT

Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module message_defaut_chargement. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module message_defaut_chargement. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub
Sub message_defaut_claquage()

'def_claquage : numero du défaut a afficher

'On Error GoTo 400

Dim i As Integer
Dim Copy_Tab(20) As Long

' on quit si il n'y a pas de défaut
If def_claquage = 0 Then
    Exit Sub
End If

'TRAITEMENT DEFAUT

DataEnvironment1.rsTable_defaut_dielectric.Filter = "Code = '" & def_claquage & "'"
    
    If DataEnvironment1.rsTable_defaut_dielectric.State = adStateClosed Then
        DataEnvironment1.rsTable_defaut_dielectric.Open
    End If

    If DataEnvironment1.rsTable_defaut_dielectric.RecordCount > 0 Then
        If DataEnvironment1.rsTable_defaut_dielectric.Fields("Affiche").Value = True Then
            
            Dim Affiche As New Affiche_defaut.Affiche ' Fenetre d'affichage de défaut
            Dim Reponce As Affiche_defaut.Bouton ' Reponce de la fenetre d'affichage de défaut
            
            Affiche.Text_Color = vbRed
            Reponce = Affiche.Affiche_Message(DataEnvironment1.rsTable_defaut_dielectric.Fields("Labelle"), "Défaut post contrôle claquage", DataEnvironment1.rsTable_defaut_dielectric.Fields("Reprise_autorise"))
            
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Copy_Tab(i) = Tab_Status_banc(i)
            Next i
            
            If Reponce = Reprise Then
                set_bit Copy_Tab(6), 13
                If def_claquage <> 0 Then
                   set_bit Copy_Tab(4), 7
                   def_claquage = 0
                   Copy_Tab(8) = 0
                End If

            Else 'Abondan teste
                set_bit Copy_Tab(6), 14
                If def_claquage <> 0 Then
                   set_bit Copy_Tab(4), 7
                   def_claquage = 0
                   Copy_Tab(8) = 0
                End If
            End If
        End If
    Else
        Affiche.Affiche_Message "Défaut N° " & def_claquage, "Défaut post claquage", False
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Copy_Tab(i) = Tab_Status_banc(i)
        Next i
        'Aquitement défaut
        set_bit Copy_Tab(6), 14
        If def_claquage <> 0 Then
           set_bit Copy_Tab(4), 8
           def_claquage = 0
           Copy_Tab(8) = 0
        End If
    End If
    
    Dim Com_erreure As Integer
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 10)
        If Com_erreure < 0 Then ' Communication en erreur
            If MsgBox(ERREUR_ECRITURE_DEFAUT, vbCritical + vbRetryCancel) = vbCancel Then
                If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                    End
                End If
            End If
        Else ' Communication sans erreur
            'memo defaut precedent
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            previous_def_claquage = def_claquage
            Exit Sub
        End If
    Loop
    
'FIN DE TRAITEMENT DEFAUT

Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module message_defaut_claquage. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module message_defaut_claquage. Laporkan ke maintenance segera. Thanks."
Resume Next


End Sub

Sub message_defaut_ctrl_final()

Dim Com_erreure As Integer

Dim i As Integer
Dim Copy_Tab(20) As Long

' on quit si il n'y a pas de défaut
If def_final = 0 Then
    Exit Sub
End If

'On Error GoTo 400

    DataEnvironment1.rstable_defaut_final.Filter = "Code = '" & def_final & "'"

    If DataEnvironment1.rstable_defaut_final.State = adStateClosed Then
        DataEnvironment1.rstable_defaut_final.Open
    End If

    If DataEnvironment1.rstable_defaut_final.RecordCount > 0 Then
        If DataEnvironment1.rstable_defaut_final.Fields("Affiche").Value = True Then
            
            Dim Affiche As New Affiche_defaut.Affiche ' Fenetre d'affichage de défaut
            Dim Reponce As Affiche_defaut.Bouton ' Reponce de la fenetre d'affichage de défaut
            
            Affiche.Text_Color = vbRed
            Reponce = Affiche.Affiche_Message(DataEnvironment1.rstable_defaut_final.Fields("Labelle"), "Défaut post contrôle final", DataEnvironment1.rstable_defaut_final.Fields("Reprise_autorise"))
            
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Copy_Tab(i) = Tab_Status_banc(i)
            Next i
            
            If Reponce = Reprise Then
                set_bit Copy_Tab(6), 13
                If def_final <> 0 Then
                   set_bit Copy_Tab(4), 8
                   def_final = 0
                   Copy_Tab(9) = 0
                End If

            Else 'Vidage poste
                set_bit Copy_Tab(6), 14
                If def_final <> 0 Then
                   set_bit Copy_Tab(4), 8
                   def_final = 0
                   Copy_Tab(9) = 0
                End If
            End If
        End If
    Else
        Affiche.Affiche_Message "Défaut N° " & def_final, "Défaut post contrôle final", False
        
        For i = LBound(Copy_Tab) To UBound(Copy_Tab)
            Copy_Tab(i) = Tab_Status_banc(i)
        Next i
        
        'Vidange poste
        set_bit Copy_Tab(6), 14
        If def_final <> 0 Then
           set_bit Copy_Tab(4), 8
           def_final = 0
           Copy_Tab(9) = 0
        End If
    End If
    
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 10)
        If Com_erreure <= 0 Then ' Communication en erreur
            If MsgBox(ERREUR_ECRITURE_DEFAUT, vbCritical + vbRetryCancel) = vbCancel Then
                If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                    End
                End If
            End If
        Else ' Communication sans erreur
            'memo defaut precedent
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            previous_def_final = def_final
            Exit Sub
        End If
    Loop
    
Exit Sub

'400 MsgBox "Erreur " & Str$(err) & " module message_defaut_ctrl_final. Signalez immediatement cette erreur. Merci."
400 MsgBox "Erreur " & Str$(err) & " module message_defaut_ctrl_final. Laporkan ke maintenance segera. Thanks."
Resume Next

End Sub




