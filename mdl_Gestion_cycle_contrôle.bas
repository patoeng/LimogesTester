Attribute VB_Name = "mdl_Gestion_cycle_contr�le"
Global Const MAX_CAPACITE_PILLE = 252 ' Nombre maximun de fonction �l�m�taire par cycle

Global Const ADRESSE_MEM_API_CYCLE_NOMINAL = 5000 'Adresse m�moire de base API pir les dif�rents cycles
Global Const ADRESSE_MEM_API_CYCLE_VIDANGE_POST = 5260
Global Const ADRESSE_MEM_API_CYCLE_POST_INI = 5780
Global Const ADRESSE_MEM_API_CYCLE_REPRISE = 5520
Global Const ADRESSE_MEM_API_CYCLE_POST_NON_ACTIF = 6040
Global Const ADRESSE_MEM_API_CYCLE_EN_DEFAUT_MAJEUR = 6300

Global Const LONGUEUR_TRAM_MAX = 100 ' Nombre de mot maximum par trame

Global Const CODE_PRODUIT_SERVICE = 1000 ' La modification de ces valeur implique une mofication des base de donn�es
Global Const CYCLE_POST_INIT = 6         ' Relative au stockage des cycles
Global Const CYCLE_VIDANGE_POST = 5      ' Afain d'asur� la une co�rence
Global Const CYCLE_REPRISE = 7
Global Const CYCLE_POST_NON_ACTIF = 8
Global Const CYCLE_EN_DEFAUT_MAJEUR = 9
Global Const CYCLE_TEMPERATURE_BASSE = 1
Global Const CYCLE_TEMPERATURE_HAUTE = 2
Global Const CYCLE_CONTROLE_FINAL = 4
Global Const CYCLE_PREMIER_CONTROLE = 3

Global Const DB_PRODUIT_ANA_IS = 2 'd�fini le type etiquette pour produit ana is
Global Const DB_PRODUIT_ANA_IS_IC = 3 'd�fini le type etiquette pour produit ana is+ic
Global Const DB_PRODUIT_TOR = 1 'd�fini le type etiquette pour produit TOR ou TOR auto adaptable

Global Type_etiquette As Integer 'Type d'�tiquette � utiliser

'Global Const ERREUR_ECRITURE_CYCLE_CONTROLE = "Erreur d'�criture des cycles de contr�le dans l'API"
Global Const Msg_Error_01 = "Error write test cyles in PLC"

'Ecrir les cycles de contr�le dans l'API
Function Ecrire_cycle_mes_API(Donn�e() As Long, Longeur As Integer, Adresse_base As Integer) As Integer

Dim Com_erreure As Long
Dim i As Integer
Dim n As Integer
Dim Longueur As Integer

'Ecriture des trames
i = 0

While LONGUEUR_TRAM_MAX * i < UBound(Donn�e)
    Longueur = UBound(Donn�e) - LONGUEUR_TRAM_MAX * i
    If Longueur > LONGUEUR_TRAM_MAX Then Longueur = LONGUEUR_TRAM_MAX
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas r�usi
        Com_erreure = Ecrir_mots(Adresse_base, Donn�e, Longueur, LONGUEUR_TRAM_MAX * i)
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


'G�n�re les sequences de contr�les apropri�es au produits
'et les envoie vers l'Api
Function G�n�rer_sequence_API_contr�le_pos_final(Code_badge As Integer, Code_cycle_nominal As Integer) As Integer

Dim Code_produit As Integer
Dim Code_produit_g�n�rique As Integer
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
    
    ' On ferche le code produit � utiliser
    DataEnvironment1.rsRequette_produit.MoveFirst
    Code_produit = DataEnvironment1.rsRequette_produit.Fields("Code").Value
    Type_etiquette = DataEnvironment1.rsRequette_produit.Fields("Type_etiquette").Value
    
    If DataEnvironment1.rsRequette_produit_g�n�rique.State <> adStateClosed Then
        DataEnvironment1.rsRequette_produit_g�n�rique.Close 'ferme le requette avant de l'ouvir
    End If
    
    DataEnvironment1.Requette_produit_g�n�rique Code_produit 'ouvre la requette
    If DataEnvironment1.rsRequette_produit_g�n�rique.RecordCount > 0 Then
        
        ' On ferche le code produit g�n�rique � utiliser
        DataEnvironment1.rsRequette_produit_g�n�rique.MoveFirst
        Code_produit_g�n�rique = DataEnvironment1.rsRequette_produit_g�n�rique.Fields("Produit_generique").Value
        
        'g�n�re sequence contr�le nominal
        If DataEnvironment1.rsRequette_code_cycle_nominal.State <> adStateClosed Then
            DataEnvironment1.rsRequette_code_cycle_nominal.Close 'ferme le requette avant de l'ouvir
        End If
        
        DataEnvironment1.Requette_code_cycle_nominal Code_cycle_nominal, ctrl_simplifie, Code_produit, Code_produit_g�n�rique 'ouvre la requette
        If DataEnvironment1.rsRequette_code_cycle_nominal.RecordCount > 0 Then
            
            ' On cherche le code cycle
            DataEnvironment1.rsRequette_produit_g�n�rique.MoveFirst
            Code_cycle = DataEnvironment1.rsRequette_code_cycle_nominal.Fields("N�").Value
            
            'g�n�re sequence contr�le nominal
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
                
                '�crire le cycle dans la m�moire API
                Com_erreur = Ecrire_cycle_mes_API(Sequence, i - 1, ADRESSE_MEM_API_CYCLE_NOMINAL)
                If Com_erreur < 0 Then
                    G�n�rer_sequence_API_contr�le_pos_final = -1
                    Unload FRM_Telechargement_cycle
                    Exit Function
                End If
            Else
                ' Pas de s�quence programm�e
                'MsgBox "Pas de s�quence programm�e pour de type de contr�le", vbCritical + vbOKOnly, App.Comments
                MsgBox "No Programmed sequence for selected test type", vbCritical + vbOKOnly, App.Comments
                G�n�rer_sequence_API_contr�le_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
            
        Else
            ' Pas pas de code cycle pour le type de contr�le demand�
            'MsgBox "Pas de code cylce touv� pour ce type de contr�le", vbCritical + vbOKOnly, App.Comments
            MsgBox "No cycle code found for selected test type", vbCritical + vbOKOnly, App.Comments
            G�n�rer_sequence_API_contr�le_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = G�r�_cycle_service(CYCLE_POST_INIT, Code_produit_g�n�rique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            '�crire le cycle dans le m�moire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_POST_INI)
            If Com_erreur < 0 Then
                G�n�rer_sequence_API_contr�le_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            G�n�rer_sequence_API_contr�le_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = G�r�_cycle_service(CYCLE_VIDANGE_POST, Code_produit_g�n�rique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            '�crire le cycle dans le m�moire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_VIDANGE_POST)
            If Com_erreur < 0 Then
                G�n�rer_sequence_API_contr�le_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            G�n�rer_sequence_API_contr�le_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = G�r�_cycle_service(CYCLE_REPRISE, Code_produit_g�n�rique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            '�crire le cycle dans le m�moire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_REPRISE)
            If Com_erreur < 0 Then
                G�n�rer_sequence_API_contr�le_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            G�n�rer_sequence_API_contr�le_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = G�r�_cycle_service(CYCLE_POST_NON_ACTIF, Code_produit_g�n�rique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            '�crire le cycle dans le m�moire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_POST_NON_ACTIF)
            If Com_erreur < 0 Then
                G�n�rer_sequence_API_contr�le_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            G�n�rer_sequence_API_contr�le_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        Sequence_taille = G�r�_cycle_service(CYCLE_EN_DEFAUT_MAJEUR, Code_produit_g�n�rique, Sequence)
        FRM_Telechargement_cycle.Set_avancement (Coumpteur)
        Coumpteur = Coumpteur + 1
        If Sequence_taille > 0 Then
            '�crire le cycle dans le m�moire API
            Com_erreur = Ecrire_cycle_mes_API(Sequence, Sequence_taille, ADRESSE_MEM_API_CYCLE_EN_DEFAUT_MAJEUR)
            If Com_erreur < 0 Then
                G�n�rer_sequence_API_contr�le_pos_final = -1
                Unload FRM_Telechargement_cycle
                Exit Function
            End If
        Else
            G�n�rer_sequence_API_contr�le_pos_final = -1
            Unload FRM_Telechargement_cycle
            Exit Function
        End If
        
        
        
        ' Touts les cycles sont tranfair� dans l'API
        G�n�rer_sequence_API_contr�le_pos_final = 1
        Unload FRM_Telechargement_cycle
        
    Else
        ' Pas de produit g�n�rique ref�renc� pour le code cycle
        'MsgBox "Pas de produit g�n�rique asoci� au type de produit", vbCritical + vbOKOnly, App.Comments
        MsgBox "No Generic product related to selected product type", vbCritical + vbOKOnly, App.Comments
        G�n�rer_sequence_API_contr�le_pos_final = -1
        Unload FRM_Telechargement_cycle
        Exit Function
    End If

Else
    ' Pas de code produit associ� au code badge (%Mw468 ou %Mw524)
    'MsgBox "Configuration produit non g�r� par ce poste de contr�le", vbCritical + vbOKOnly, App.Comments
    MsgBox "Product configuration is not managed in this station", vbCritical + vbOKOnly, App.Comments
    G�n�rer_sequence_API_contr�le_pos_final = -1
    Unload FRM_Telechargement_cycle
    Exit Function
End If

End Function

Function G�r�_cycle_service(Cycle As Integer, Produit_g�n�rique As Integer, Tableau() As Long) As Integer

Dim i As Integer
Dim n As Integer

Dim Text As String
Dim Code_cycle

'Asoci� un text au type de cycle de service
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

'g�n�re sequence de service
If DataEnvironment1.rsRequette_code_cycle_service.State <> adStateClosed Then
    DataEnvironment1.rsRequette_code_cycle_service.Close 'ferme le requette avant de l'ouvir
End If

DataEnvironment1.Requette_code_cycle_service Cycle, Produit_g�n�rique, CODE_PRODUIT_SERVICE  'ouvre la requette
If DataEnvironment1.rsRequette_code_cycle_service.RecordCount > 0 Then
    
    ' On cherche le code cyle
    DataEnvironment1.rsRequette_code_cycle_service.MoveFirst
    Code_cycle = DataEnvironment1.rsRequette_code_cycle_service.Fields("N�").Value
    
    'g�n�re sequence contr�le nominal
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
        G�r�_cycle_service = i - 1
    Else
        ' Pas de s�quence programm�e
        'MsgBox "Pas de s�quence programm�e pour ce type de famille de produit" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
        MsgBox "No Programmed sequence for this type of product family" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
        G�r�_cycle_service = -1
        Exit Function
    End If
    
Else
    ' Pas pas de code cycle pour le type de famille de produit
    'MsgBox "Pas de code cylce de service touv� pour ce type de famille de produit" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
    MsgBox "No code cycle service found for this type of product family" & vbCrLf & Text, vbCritical + vbOKOnly, App.Comments
    G�r�_cycle_service = -1
    Exit Function
End If

End Function

