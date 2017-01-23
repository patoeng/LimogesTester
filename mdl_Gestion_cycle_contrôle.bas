Attribute VB_Name = "mdl_Gestion_cycle_contrôle"
Global Const MAX_CAPACITE_PILLE = 252 ' Nombre maximun de fonction élémétaire par cycle

Global Const ADRESSE_MEM_API_CYCLE_NOMINAL = 5000 'Adresse mémoire de base API pir les diférents cycles
Global Const ADRESSE_MEM_API_CYCLE_VIDANGE_POST = 5260
Global Const ADRESSE_MEM_API_CYCLE_POST_INI = 5780
Global Const ADRESSE_MEM_API_CYCLE_REPRISE = 5520
Global Const ADRESSE_MEM_API_CYCLE_POST_NON_ACTIF = 6040
Global Const ADRESSE_MEM_API_CYCLE_EN_DEFAUT_MAJEUR = 6300

Global Const LONGUEUR_TRAM_MAX = 100 ' Nombre de mot maximum par trame

Global Const CODE_PRODUIT_SERVICE = 1000 ' La modification de ces valeur implique une mofication des base de données
Global Const CYCLE_POST_INIT = 6         ' Relative au stockage des cycles
Global Const CYCLE_VIDANGE_POST = 5      ' Afain d'asuré la une coérence
Global Const CYCLE_REPRISE = 7
Global Const CYCLE_POST_NON_ACTIF = 8
Global Const CYCLE_EN_DEFAUT_MAJEUR = 9
Global Const CYCLE_TEMPERATURE_BASSE = 1
Global Const CYCLE_TEMPERATURE_HAUTE = 2
Global Const CYCLE_CONTROLE_FINAL = 4
Global Const CYCLE_PREMIER_CONTROLE = 3

Global Const DB_PRODUIT_ANA_IS = 2 'défini le type etiquette pour produit ana is
Global Const DB_PRODUIT_ANA_IS_IC = 3 'défini le type etiquette pour produit ana is+ic
Global Const DB_PRODUIT_TOR = 1 'défini le type etiquette pour produit TOR ou TOR auto adaptable

Global Type_etiquette As Integer 'Type d'étiquette à utiliser

'Global Const ERREUR_ECRITURE_CYCLE_CONTROLE = "Erreur d'écriture des cycles de contrôle dans l'API"
Global Const Msg_Error_01 = "Error write test cyles in PLC"

'Ecrir les cycles de contrôle dans l'API
Function Ecrire_cycle_mes_API(Donnée() As Long, Longeur As Integer, Adresse_base As Integer) As Integer

Dim Com_erreure As Long
Dim i As Integer
Dim n As Integer
Dim Longueur As Integer

'Ecriture des trames
i = 0

While LONGUEUR_TRAM_MAX * i < UBound(Donnée)
    Longueur = UBound(Donnée) - LONGUEUR_TRAM_MAX * i
    If Longueur > LONGUEUR_TRAM_MAX Then Longueur = LONGUEUR_TRAM_MAX
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots(Adresse_base, Donnée, Longueur, LONGUEUR_TRAM_MAX * i)
        If Com_erreure <= 0 Then ' Communication en erreur
            'If MsgBox(ERREUR_ECRITURE_CYCLE_CONTROLE, vbCritical + vbRetryCancel) = vbCancel Then
            If MsgBox(Msg_Error_01, vbCritical + vbRetryCancel) = vbCancel Then
               Ecrire_cycle_mes_API = -1
               Exit Function
            End If
        End If
        i = i + 1
    Loop
Wend

End Function


'Génére les sequences de contrôles apropriées au produits
'et les envoie vers l'Api
Function Générer_sequence_API_contrôle_pos_final(Code_badge As Integer, Code_cycle_nominal As Integer) As Integer

Dim Code_produit As Integer
Dim Code_produit_générique As Integer
Dim Code_cycle As Integer

Dim Sequence(259) As Long
Dim i As Integer
Dim n As Integer
Dim Sequence_taille As Integer

Dim Coumpteur As Integer

Dim Com_erreur As Integer

' Affiche une fenetre indiquant l'etat du chargment des cycles dans l'API
FRM_Telechargement_cycle.Show
Coumpteur = 0
FRM_Telechargement_cycle.Set_avancement Coumpteur, 6
Coumpteur = Coumpteur + 1

If DataEnvironment1.rsRequette_produit.State <> adStateClosed Then
    DataEnvironment1.rsRequette_produit.Close 'ferme le requette avant de l'ouvir
End If

DataEnvironment1.Requette_produit Code_badge 'ouvre la requette
If DataEnvironment1.rsRequette_produit.RecordCount > 0 Then
    
    ' On ferche le code produit à utiliser
    DataEnvironment1.rsRequette_produit.MoveFirst
    Code_produit = DataEnvironment1.rsRequette_produit.Fields("Code").Value
    Type_etiquette = DataEnvironment1.rsRequette_produit.Fields("Type_etiquette").Value
    
    If DataEnvironment1.rsRequette_produit_générique.State <> adStateClosed Then
        DataEnvironment1.rsRequette_produit_générique.Close 'ferme le requette avant de l'ouvir
    End If
    
    DataEnvironment1.Requette_produit_générique Code_produit 'ouvre la requette
    If DataEnvironment1.rsRequette_produit_générique.RecordCount > 0 Then
        
        ' On ferche le code produit générique à utiliser
        DataEnvironment1.rsRequette_produit_générique.MoveFirst
        Code_produit_générique = DataEnvironment1.rsRequette_produit_générique.Fields("Produit_generique").Value
        
        'génére sequence contrôle nominal
        If DataEnvironment1.rsRequette_code_cycle_nominal.State <> adStateClosed Then
            DataEnvironment1.rsRequette_code_cycle_nominal.Close 'ferme le requette avant de l'ouvir
        End If
        
        DataEnvironment1.Requette_code_cycle_nominal Code_cycle_nominal, ctrl_simplifie, Code_produit, Code_produit_générique 'ouvre la requette
        If DataEnvironment1.rsRequette_code_cycle_nominal.RecordCount > 0 Then
            
            ' On cherche le code cycle
            DataEnvironment1.rsRequette_produit_générique.MoveFirst
            Code_cycle = DataEnvironment1.rsRequette_code_cycle_nominal.Fields("N°").Value
            
            'génére sequence contrôle nominal
            If DataEnvironment1.rsSequence.State <> adStateClosed Then
                DataEnvironment1.rsSequence.Close 'ferme le requette avant de l'ouvir
            End If
            
            DataEnvironment1.Sequence Code_cycle  'ouvre la requette
            If DataEnvironment1.rsSequence.RecordCount > 0 Then
                
                ' On cherche la sequence
                i = 0
                DataEnvironment1.rsSequence.MoveFirst
                'Remplis la sequence dans le tableau Sequence
                While Not DataEnvironment1.rsSequence.EOF And i < MAX_CAPACITE_PILLE
                    Sequence(i) = DataEnvironment1.rsSequence.Fields("Fonction").Value
                    i = i + 1
                    DataEnvironment1.rsSequence.MoveNext
                Wend
                For n = i To UBound(Sequence)
                    Sequence(n) = 0
                Next n
                
                FRM_Telechargement_cycle.Set_avancement (Coumpteur)
                Coumpteur = Coumpteur + 1
                
                'écrire le cycle dans la mémoire API
                Com_erreur = Ecrire_cycle_mes_API(Sequence, i - 1, ADRESSE_MEM_API_CYCLE_NOMINAL)
                If Com_erreur < 0 Then
                    Générer_sequence_API_contrôle_pos_final = -1
                    Unload FRM_Telechargement_cycle
                    Exit Function
                End If
            Else
                ' Pas de séquence programmée
                'MsgBox "Pas de séquence programmée pour de type de contrôle", vbCritical + vbOKOnly, App.Comments
                MsgBox "No Programmed sequence for selected test type", vbCritical + vbOKOnly, App.Comments
                Générer_sequence_API_contrôle_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
            
        Else
            ' Pas pas de code cycle pour le type de contrôle demandé
            'MsgBox "Pas de code cylce touvé pour ce type de contrôle", vbCritical + vbOKOnly, App.Comments
            MsgBox "No cycle code found for selected test type", vbCritical + vbOKOnly, App.Comments
            Générer_sequence_API_contrôle_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = Géré_cycle_service(CYCLE_POST_INIT, Code_produit_générique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            'écrire le cycle dans le mémoire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_POST_INI)
            If Com_erreur < 0 Then
                Générer_sequence_API_contrôle_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            Générer_sequence_API_contrôle_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = Géré_cycle_service(CYCLE_VIDANGE_POST, Code_produit_générique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            'écrire le cycle dans le mémoire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_VIDANGE_POST)
            If Com_erreur < 0 Then
                Générer_sequence_API_contrôle_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            Générer_sequence_API_contrôle_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = Géré_cycle_service(CYCLE_REPRISE, Code_produit_générique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            'écrire le cycle dans le mémoire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_REPRISE)
            If Com_erreur < 0 Then
                Générer_sequence_API_contrôle_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            Générer_sequence_API_contrôle_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = Géré_cycle_service(CYCLE_POST_NON_ACTIF, Code_produit_générique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            'écrire le cycle dans le mémoire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_POST_NON_ACTIF)
            If Com_erreur < 0 Then
                Générer_sequence_API_contrôle_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            Générer_sequence_API_contrôle_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = Géré_cycle_service(CYCLE_EN_DEFAUT_MAJEUR, Code_produit_générique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            'écrire le cycle dans le mémoire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_EN_DEFAUT_MAJEUR)
            If Com_erreur < 0 Then
                Générer_sequence_API_contrôle_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            Générer_sequence_API_contrôle_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        
        
        ' Touts les cycles sont tranfairé dans l'API
        Générer_sequence_API_contrôle_pos_final = 1
        Unload FRM_Telechargement_cycle
        
    Else
        ' Pas de produit générique reférencé pour le code cycle
        'MsgBox "Pas de produit générique asocié au type de produit", vbCritical + vbOKOnly, App.Comments
        MsgBox "No Generic product related to selected product type", vbCritical + vbOKOnly, App.Comments
        Générer_sequence_API_contrôle_pos_final = -1
        Unload FRM_Telechargement_cycle
        Exit Function
    End If

Else
    ' Pas de code produit associé au code badge (%Mw468 ou %Mw524)
    'MsgBox "Configuration produit non géré par ce poste de contrôle", vbCritical + vbOKOnly, App.Comments
    MsgBox "Product configuration is not managed in this station", vbCritical + vbOKOnly, App.Comments
    Générer_sequence_API_contrôle_pos_final = -1
    Unload FRM_Telechargement_cycle
    Exit Function
End If

End Function

Function Géré_cycle_service(Cycle As Integer, Produit_générique As Integer, Tableau() As Long) As Integer

Dim i As Integer
Dim n As Integer

Dim Text As String
Dim Code_cycle

'Asocié un text au type de cycle de service
Select Case Cycle
    Case CYCLE_POST_INIT
        Text = "Int poste"
    Case CYCLE_VIDANGE_POST
        Text = "Vidange poste"
    Case CYCLE_REPRISE
        Text = "Reprise"
    Case CYCLE_POST_NON_ACTIF
        Text = "Post non actif"
End Select

'génére sequence de service
If DataEnvironment1.rsRequette_code_cycle_service.State <> adStateClosed Then
    DataEnvironment1.rsRequette_code_cycle_service.Close 'ferme le requette avant de l'ouvir
End If

DataEnvironment1.Requette_code_cycle_service Cycle, Produit_générique, CODE_PRODUIT_SERVICE  'ouvre la requette
If DataEnvironment1.rsRequette_code_cycle_service.RecordCount > 0 Then
    
    ' On cherche le code cyle
    DataEnvironment1.rsRequette_code_cycle_service.MoveFirst
    Code_cycle = DataEnvironment1.rsRequette_code_cycle_service.Fields("N°").Value
    
    'génére sequence contrôle nominal
    If DataEnvironment1.rsSequence.State <> adStateClosed Then
        DataEnvironment1.rsSequence.Close 'ferme le requette avant de l'ouvir
    End If
    
    DataEnvironment1.Sequence Code_cycle  'ouvre la requette
    If DataEnvironment1.rsSequence.RecordCount > 0 Then
        
        ' On cherche la sequence
        i = 0
        DataEnvironment1.rsSequence.MoveFirst
        'Remplis la sequence dans le tableau Sequence
        While Not DataEnvironment1.rsSequence.EOF And i < MAX_CAPACITE_PILLE
            Tableau(i) = DataEnvironment1.rsSequence.Fields("Fonction").Value
            i = i + 1
            DataEnvironment1.rsSequence.MoveNext
        Wend
        For n = i To UBound(Tableau)
            Tableau(n) = 0
        Next n
        Géré_cycle_service = i - 1
    Else
        ' Pas de séquence programmée
        'MsgBox "Pas de séquence programmée pour ce type de famille de produit" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
        MsgBox "No Programmed sequence for this type of product family" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
        Géré_cycle_service = -1
        Exit Function
    End If
    
Else
    ' Pas pas de code cycle pour le type de famille de produit
    'MsgBox "Pas de code cylce de service touvé pour ce type de famille de produit" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
    MsgBox "No code cycle service found for this type of product family" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
    Géré_cycle_service = -1
    Exit Function
End If

End Function

