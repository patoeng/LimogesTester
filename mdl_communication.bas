Attribute VB_Name = "mdl_Communication"
'Global Const ERREUR_ECRITURE_DEFAUT = "Erreure écriture mots de gestion de défaut l'API"
'Global Const ERREUR_ECRITURE_MODE_DE_MARCHE = "Erreure écriture du mode de marche de l'API"
'Global Const ERREUR_ECRITURE_API = "Erreure écriture vers l'API"
'Global Const ERREUR_ECRITURE_APPELE_MAINTENANCE = "Erreure écriture du mode maintenance"
'Global Const ERREUR_ECRITURE_AUTO_TEST = "Errreur écriture API en mode Auto test"
Global Const ERREUR_ECRITURE_DEFAUT = "Error writing memory word on PLC"
Global Const ERREUR_ECRITURE_MODE_DE_MARCHE = "Error writing operation mode on PLC"
Global Const ERREUR_ECRITURE_API = "Error writing to PLC"
Global Const ERREUR_ECRITURE_APPELE_MAINTENANCE = "Error Writing maintenance Mode"
Global Const ERREUR_ECRITURE_AUTO_TEST = "Error Writing Autotest mode to PLC"
Global Const MAX_TEMPS_OCCUPE = 20 ' 5 seconds
Global Const MAX_TEMPS_REPONCE = 20 ' 1 seconds

Global Const COMMUNICATION_OCCUPE = -1  ' port de communication toujours ocupé aprés un temps de MAX_TEMPS
                                        '' Response time
Global Const COMMUNICATION_TIME_OUT = -2 ' dépassement de temps de réponce de MAX_TEMS
                                        '' Time out time
Global Const FONCTION_ABONDONE = -3 ' Fonction annulé par une autre fonction
                                    '' Abort Function
Global Const INDICE_HORS_PLAGE = -4 ' Index initial n'apatien pas au tableau
                                    '' Initial Index for tbale

Global Const ADRESSE_BASE_ETAT_BANC = 350   ' Adresse de début de la zone mémoire etat banc de l'API
                                            '' Start Memory Address of bench PLC

Public Socket_valid As Boolean 'la connection à la socket est valide
Public Socket_occupé As Boolean 'La ocket est occupé
Public Fonction_com_en_cours As Boolean 'Une fonction de communication est en cours d'execution

Public Unitel_Adresse As Structure_unite_adresse 'Adresse unitel de l'IHM et de l'API

Public cycle_mise_a_jour_données_API As Integer 'Etape de la boucle de mis à jours des information de API

Public Enum Trame_commende
   Fonction_Lecture_etat_banc = 0
   Fonction_Lecture_position_régle = 1
   Fonction_Lecture_paramétre_produit = 2
   Fonction_Lecture_messages_produit_1 = 3
   Fonction_Lecture_messages_produit_2 = 4
   Fonction_Ecriture_mots = 5
End Enum

Public Type Utwmaitmende
    Nombre_total_trame As Long
    Nombre_trame As Long
    Commende As Trame_commende
    Nombre_mots As Integer
    Trame(128) As Long
End Type

Global Commende_com As Utwmaitmende

' Mot echange PC<->TSX
Public Tab_Status_banc(70) As Long

Public Function Debut_communication() As Long
Dim Msg_bouton As Integer

début:
    Utwmait.Winsock1.Connect
    DoEvents
    Socket_valid = True 'La socket est valide par défaut
    Do While Utwmait.Winsock1.State = sckConnecting Or Utwmait.Winsock1.State = sckResolvingHost Or Utwmait.Winsock1.State = sckHostResolved
        DoEvents
    Loop
    
    If Utwmait.Winsock1.State <> sckConnected Then
            'Msg_bouton = MsgBox("Pas de connection possible avec l'API" & vbCrLf & _
            '    "Adresse de l'API : " & Utwmait.Winsock1.RemoteHost, vbAbortRetryIgnore + vbSystemModal + vbCritical, App.Title)
            Msg_bouton = MsgBox("No Connection possible with PLC" & vbCrLf & _
                "Address of PLC : " & Utwmait.Winsock1.RemoteHost, vbAbortRetryIgnore + vbSystemModal + vbCritical, App.Title)
            Select Case Msg_bouton
                Case vbIgnore
                    Socket_valid = False 'Fonctions de communication ne seront pas utilisable
                    Utwmait.Timer_mis_a_jour.Interval = 0 'Le timer de mis à jour des données API n'est plus activé
                    Utwmait.Option_fonction_en_cours = False
                    Utwmait.Option_socket_occupée = False
                    Utwmait.Option_socket_valide = False
                    
                    'MsgBox "Les fonctions de communication avec API ne seront pas ativées", vbInformation + vbApplicationModal, App.Title
                    MsgBox "Communication function with PLC is not activate", vbInformation + vbApplicationModal, App.Title
                    Utwmait.Winsock1.Close
                    Debut_communication = 0
                    
                    Exit Function
                Case vbAbort
                     End ' Fin du programme
                Case vbRetry
                    Utwmait.Winsock1.Close
                    GoTo début
            End Select
    End If
        
    Unitel_Adresse = Get_Adresse_API(File_para_ini)
        
    cycle_mise_a_jour_données_API = 0
    Socket_occupé = False
    Fonction_com_en_cours = False
    
    Debut_communication = Utwmait.Winsock1.LocalPort 'Renvoie le port local de la socket
End Function

Public Sub Fin_communication()

    If Utwmait.Winsock1.State <> sckClosed Then
        Utwmait.Winsock1.Close
    End If
    Socket_valid = False

End Sub

Public Function Lecture_messages_produit_2(Messages() As Long) As Long
Dim Adresse_table As Integer
Dim Longueur As Integer

Dim i As Integer
Dim j As Integer

ReDim Messages(504)
    
    Début_temps = Timer
    Do While Socket_occupé = True 'Or Fonction_com_en_cours = True
        If Timer > Début_temps + MAX_TEMPS_OCCUPE Then
           Fonction_com_en_cours = False
           Lecture_messages_produit_2 = COMMUNICATION_OCCUPE
           Exit Function
        End If
        DoEvents
        'Utwmait.Events
    Loop
    
    Fonction_com_en_cours = True
    
    Commende_com.Nombre_total_trame = 4
    Commende_com.Nombre_trame = 1
    Commende_com.Commende = Trame_commende.Fonction_Lecture_messages_produit_2
    
    Do While Commende_com.Nombre_total_trame >= Commende_com.Nombre_trame
        Longueur = 126
        Adresse_table = 4000 + (Longueur * (Commende_com.Nombre_trame - 1))
        Commende_com.Nombre_mots = Longeur
        Utwmait.Envoie_UTW Chr$(&H36) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(Adresse_table Mod 256) & Chr$(Int(Adresse_table / 256)) & Chr$(Longueur Mod 256) & Chr$(Int(Longueur / 256))
        Début_temps = Timer
        Do While Socket_occupé = True
            If Timer > Début_temps + MAX_TEMPS_REPONCE Then
               Socket_occupé = False
               Fonction_com_en_cours = False
               Lecture_messages_produit_2 = COMMUNICATION_TIME_OUT
               Exit Function
            End If
            DoEvents
        Loop
        
        If Commende_com.Commende <> Trame_commende.Fonction_Lecture_messages_produit_2 Then
            Lecture_messages_produit_2 = FONCTION_ABONDONE
            Exit Function
        End If
        
        If Commende_com.Nombre_mots >= Longueur Then
            j = 1
            For i = (Commende_com.Nombre_mots * (Commende_com.Nombre_trame - 1)) To (Commende_com.Nombre_mots * (Commende_com.Nombre_trame - 1)) + Commende_com.Nombre_mots
                Messages(i) = Commende_com.Trame(j)
            j = j + 1
            Next i
        End If
        Commende_com.Nombre_trame = Commende_com.Nombre_trame + 1
    Loop
    Lecture_messages_produit_2 = i
    Fonction_com_en_cours = False
End Function

Public Function Lecture_messages_produit_1(Messages() As Long) As Long

Dim Adresse_table As Integer
Dim Longueur As Integer

Dim i As Integer
Dim j As Integer

ReDim Messages(504)

    Début_temps = Timer
    Do While Socket_occupé = True 'Or Fonction_com_en_cours = True
        If Timer > Début_temps + MAX_TEMPS_OCCUPE Then
           Fonction_com_en_cours = False
           Lecture_messages_produit_1 = COMMUNICATION_OCCUPE
           Exit Function
        End If
        DoEvents
        'Utwmait.Events
    Loop
    
    Fonction_com_en_cours = True
    
    Commende_com.Nombre_total_trame = 4
    Commende_com.Nombre_trame = 1
    Commende_com.Commende = Trame_commende.Fonction_Lecture_messages_produit_1
    
    Do While Commende_com.Nombre_total_trame >= Commende_com.Nombre_trame
        Longueur = 126
        Adresse_table = 3000 + (Longueur * (Commende_com.Nombre_trame - 1))
        Commende_com.Nombre_mots = Longeur
        Utwmait.Envoie_UTW (Chr$(&H36) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(Adresse_table Mod 256) & Chr$(Int(Adresse_table / 256)) & Chr$(Longueur Mod 256) & Chr$(Int(Longueur / 256)))
        Début_temps = Timer
        Do While Socket_occupé = True
            If Timer > Début_temps + MAX_TEMPS_REPONCE Then
               Socket_occupé = False
               Fonction_com_en_cours = False
               Lecture_messages_produit_1 = COMMUNICATION_TIME_OUT
               Exit Function
            End If
            DoEvents
        Loop
        
        If Commende_com.Commende <> Trame_commende.Fonction_Lecture_messages_produit_1 Then
            Lecture_messages_produit_1 = FONCTION_ABONDONE
            Exit Function
        End If
        
        If Commende_com.Nombre_mots >= Longueur Then
            j = 1
            For i = (Commende_com.Nombre_mots * (Commende_com.Nombre_trame - 1)) To (Commende_com.Nombre_mots * (Commende_com.Nombre_trame - 1)) + Commende_com.Nombre_mots
                Messages(i) = Commende_com.Trame(j)
            j = j + 1
            Next i
        End If
        Commende_com.Nombre_trame = Commende_com.Nombre_trame + 1
    Loop
    Lecture_messages_produit_1 = i
    
    Fonction_com_en_cours = False
    
End Function

Public Function Lecture_etat_banc() As Long

Dim Adresse_table As Integer
Dim Longueur As Integer

Dim i As Integer
Dim j As Integer

    Début_temps = Timer
    Do While Socket_occupé = True 'Or Fonction_com_en_cours = True
        If Timer > Début_temps + MAX_TEMPS_OCCUPE Then
           Fonction_com_en_cours = False
           Lecture_etat_banc = COMMUNICATION_OCCUPE
           Exit Function
        End If
        DoEvents
    Loop
    
    Fonction_com_en_cours = True
    
    Commende_com.Nombre_total_trame = 1
    Commende_com.Nombre_trame = 1
    Commende_com.Commende = Trame_commende.Fonction_Lecture_etat_banc
    
    Do While Commende_com.Nombre_total_trame >= Commende_com.Nombre_trame
        Longueur = 70
        Adresse_table = 350 + Longueur * (Commende_com.Nombre_trame - 1)
        Commende_com.Nombre_mots = Longeur
        Utwmait.Envoie_UTW (Chr$(&H36) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(Adresse_table Mod 256) & Chr$(Int(Adresse_table / 256)) & Chr$(Longueur Mod 256) & Chr$(Int(Longueur / 256)))
        Début_temps = Timer
        Do While Socket_occupé = True
            If Timer > Début_temps + MAX_TEMPS_REPONCE Then
               Socket_occupé = False
               Fonction_com_en_cours = False
               Lecture_etat_banc = COMMUNICATION_TIME_OUT
               Exit Function
            End If
            DoEvents
        Loop
        
        If Commende_com.Commende <> Trame_commende.Fonction_Lecture_etat_banc Then
            Lecture_etat_banc = FONCTION_ABONDONE
            Exit Function
        End If
        
        Commende_com.Nombre_trame = Commende_com.Nombre_trame + 1
    Loop
    Lecture_etat_banc = Commende_com.Nombre_mots
    Fonction_com_en_cours = False
End Function

Public Function Lecture_position_rêgle(Position_rêgle_1 As Double, Position_rêgle_2 As Double) As Long

Dim Adresse_table As Integer
Dim Longueur As Integer

Dim Résolution As Double

    Début_temps = Timer
    Do While Socket_occupé = True 'Or Fonction_com_en_cours = True
        If Timer > Début_temps + MAX_TEMPS_OCCUPE Then
           Fonction_com_en_cours = False
           Lecture_position_rêgle = COMMUNICATION_OCCUPE
           Exit Function
        End If
        DoEvents
    Loop
    
    Fonction_com_en_cours = True
    
    Commende_com.Nombre_total_trame = 1
    Commende_com.Nombre_trame = 1
    Commende_com.Commende = Trame_commende.Fonction_Lecture_position_régle
    
    Do While Commende_com.Nombre_total_trame >= Commende_com.Nombre_trame
        Longueur = 10
        Adresse_table = 360 + (Longueur * (Commende_com.Nombre_trame - 1))
        Commende_com.Nombre_mots = Longeur
        Utwmait.Envoie_UTW (Chr$(&H36) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(Adresse_table Mod 256) & Chr$(Int(Adresse_table / 256)) & Chr$(Longueur Mod 256) & Chr$(Int(Longueur / 256)))
        Début_temps = Timer
        Do While Socket_occupé = True
            If Timer > Début_temps + MAX_TEMPS_REPONCE Then
               Socket_occupé = False
               Fonction_com_en_cours = False
               Lecture_position_rêgle = COMMUNICATION_TIME_OUT
               Exit Function
            End If
            DoEvents
        Loop
        
        If Commende_com.Commende <> Trame_commende.Fonction_Lecture_position_régle Then
            Lecture_position_rêgle = FONCTION_ABONDONE
            Exit Function
        End If
        
        If Commende_com.Nombre_mots >= Longueur Then
            j = 1
            Résolution = CDbl(Commende_com.Trame(6))
            If Résolution > 0 Then
                Position_rêgle_1 = (CDbl(Commende_com.Trame(7)) + CDbl(Commende_com.Trame(8)) * 65526) / Résolution
                Position_rêgle_2 = (CDbl(Commende_com.Trame(9)) + CDbl(Commende_com.Trame(10)) * 65526) / Résolution
            Else
                Position_rêgle_1 = 0
                Position_rêgle_2 = 0
            End If
        End If
        Commende_com.Nombre_trame = Commende_com.Nombre_trame + 1
    Loop
    
    
    Lecture_position_rêgle = Commende_com.Nombre_mots
    Fonction_com_en_cours = False
End Function

Public Function Lecture_paramétre_produit(Donnée() As Long) As Long

Dim Adresse_table As Integer
Dim Longueur As Integer

Dim i As Integer
Dim j As Integer

    Début_temps = Timer
    Do While Socket_occupé = True 'Or Fonction_com_en_cours = True
        If Timer > Début_temps + MAX_TEMPS_OCCUPE Then
           Fonction_com_en_cours = False
           Lecture_paramétre_produit = COMMUNICATION_OCCUPE
           Exit Function
        End If
        DoEvents
    Loop
    
    Fonction_com_en_cours = True
    
    Commende_com.Nombre_total_trame = 1
    Commende_com.Nombre_trame = 1
    Commende_com.Commende = Trame_commende.Fonction_Lecture_paramétre_produit
    
    Do While Commende_com.Nombre_total_trame >= Commende_com.Nombre_trame
        Longueur = 126
        
        Adresse_table = 420 + (Longueur * (Commende_com.Nombre_trame - 1))
        Commende_com.Nombre_mots = Longeur
        Utwmait.Envoie_UTW (Chr$(&H36) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(Adresse_table Mod 256) & Chr$(Int(Adresse_table / 256)) & Chr$(Longueur Mod 256) & Chr$(Int(Longueur / 256)))
        Début_temps = Timer
        Do While Socket_occupé = True
            If Timer > Début_temps + MAX_TEMPS_REPONCE Then
               Socket_occupé = False
               Fonction_com_en_cours = False
               Lecture_paramétre_produit = COMMUNICATION_TIME_OUT
               Exit Function
            End If
            DoEvents
        Loop
        If Commende_com.Commende <> Trame_commende.Fonction_Lecture_paramétre_produit Then
            Lecture_paramétre_produit = FONCTION_ABONDONE
            Exit Function
         Else
             If Commende_com.Nombre_mots >= Longueur Then
                ReDim Donnée(Longueur)
                j = 1
                For i = (Commende_com.Nombre_mots * (Commende_com.Nombre_trame - 1)) To (Commende_com.Nombre_mots * (Commende_com.Nombre_trame - 1)) + Commende_com.Nombre_mots
                    Donnée(i) = Commende_com.Trame(j)
                j = j + 1
                Next i
            End If
        End If
        Commende_com.Nombre_trame = Commende_com.Nombre_trame + 1
    Loop
    Lecture_paramétre_produit = Commende_com.Nombre_mots
    Fonction_com_en_cours = False
End Function

Public Function Ecrir_mots_etat_banc(Données() As Long, Optional Longueur As Integer = -1, Optional Decalage As Integer = 0) As Long
   Ecrir_mots_etat_banc = Ecrir_mots(ADRESSE_BASE_ETAT_BANC, Données(), Longueur, Decalage)
End Function

Public Function Ecrir_mots(Adresse As Integer, Données() As Long, Optional Longueur As Integer = -1, Optional Decalage As Integer = 0) As Long

Dim Adresse_table As Integer

Dim i As Integer
Dim j As Integer

Dim Text_données As String

    Début_temps = Timer
    Do While Socket_occupé = True 'Or Fonction_com_en_cours = True
        If Timer > Début_temps + MAX_TEMPS_OCCUPE Then
           Fonction_com_en_cours = False
           Ecrir_mots = COMMUNICATION_OCCUPE
           Exit Function
        End If
        DoEvents
    Loop
    
    Fonction_com_en_cours = True
    
    Commende_com.Nombre_total_trame = 1
    Commende_com.Nombre_trame = 1
    Commende_com.Commende = Trame_commende.Fonction_Ecriture_mots
    
    Do While Commende_com.Nombre_total_trame >= Commende_com.Nombre_trame
        
        If Longueur < 0 Then
            Longueur = UBound(Données) - LBound(Données)
        End If
        If Decalage < 0 Then Decalage = 0
        'test si indice initial apartien au tableau
        If LBound(Données) + Decalage >= UBound(Données) Then
            Ecrir_mots = INDICE_HORS_PLAGE
            Exit Function
        End If
        
        Text_données = ""
        
    
        j = 1
        For i = LBound(Données) + Decalage To UBound(Données)
            If j <= Longueur Then
                Text_données = Text_données & Chr$(Données(i) Mod 256) & Chr$(Int(Données(i) / 256))
                j = j + 1
            Else
                Exit For
            End If
        Next i
         i = Len(Text_données)
        
        Adresse_table = Adresse + Decalage
        Commende_com.Nombre_mots = Longueur
        
        Utwmait.Envoie_UTW (Chr$(&H37) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(Adresse_table Mod 256) & Chr$(Int(Adresse_table / 256)) & Chr(Longueur) & Chr$(0) & Text_données)
        
        Début_temps = Timer
        Do While Socket_occupé = True
            If Timer > Début_temps + MAX_TEMPS_REPONCE Then
               Socket_occupé = False
               Fonction_com_en_cours = False
               Ecrir_mots = COMMUNICATION_TIME_OUT
               Exit Function
            End If
            DoEvents
        Loop
        
        If Commende_com.Commende <> Trame_commende.Fonction_Ecriture_mots Then
            Ecrir_mots = FONCTION_ABONDONE
            Exit Function
        End If
        
        Commende_com.Nombre_trame = Commende_com.Nombre_trame + 1
    Loop
    Ecrir_mots = Commende_com.Nombre_mots
    Fonction_com_en_cours = False
End Function
