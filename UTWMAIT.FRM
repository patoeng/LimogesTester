VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Utwmait 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unitelway MASTER"
   ClientHeight    =   4848
   ClientLeft      =   2772
   ClientTop       =   2568
   ClientWidth     =   6756
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "UTWMAIT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4848
   ScaleWidth      =   6756
   WindowState     =   1  'Minimized
   Begin VB.Timer Time_out 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   4080
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   3840
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      RemoteHost      =   "API_Banc_XS"
      RemotePort      =   502
   End
   Begin VB.TextBox TXT_nb_mots 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "1"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton CMD_lit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Read"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CMD_ecrit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Write"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox TXT_adresse 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Text            =   "1"
      Top             =   2880
      Width           =   375
   End
   Begin VB.Frame FRM_stat2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3360
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
      Begin VB.CheckBox CHK_TX_valid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data ready to send"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox CHK_RX_valid 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data ready"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
   End
   Begin VB.Frame FRM_RX 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Received frame"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5535
      Begin VB.TextBox TXT_RX 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame FRM_status 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Link Status"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      Begin VB.Label LBL_status 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Etat liaison"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Timer TIM_polling 
      Interval        =   1
      Left            =   4095
      Top             =   3585
   End
   Begin VB.Label LBL_adresse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "# of words"
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label LBL_adresse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label LBL_comment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"UTWMAIT.frx":030A
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5115
   End
End
Attribute VB_Name = "Utwmait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ----------------------------------------------------------------------
'
' Quelques routines utiles pour une communication UNITELWAY maître à
' partir de Visual Basic
' Le code de ce programme appartient à "Schneider Electric".
'
' Contact :
' Jean-Luc BARRAUD Telemecanique ZI N°3 16340 L'Isle D'Espagnac
' Tel 45 69 22 22
' Fax 45 68 94 00
' Compuserve CIS : 100343,3613
' Internet : f1jek@amsat.org
' AX25 : f1jek @ f8kka.fpoc.fra.eu
' Si vous n'arrivez pas à me joindre avec ça...
'
' ----------------------------------------------------------------------
'
' Original sur Asterix:/u_bd03/ass/projets/xs_singa
'
' ----------------------------------------------------------------------
'
' Version anglaise 1.00e du 24/06/94 - 08/11/95
'
' ----------------------------------------------------------------------
'
' Visual Basic 3.0 version PRO
' ---> Necessite MSCOMM.VBX <---
'
' ----------------------------------------------------------------------


' Variables globales :
Option Explicit
DefInt A-Z

Dim kill_me

' Constantes UTW
Const STX = &H2
Const ETX = &H3
Const EOT = &H4
Const ENQ = &H5
Const ACK = &H6
Const NAK = &H15
Const DLE = &H10

'---------------------------------------
'Comm Control
'---------------------------------------
'Handshaking
Const MSCOMM_HANDSHAKE_NONE = 0
Const MSCOMM_HANDSHAKE_XONXOFF = 1
Const MSCOMM_HANDSHAKE_RTS = 2
Const MSCOMM_HANDSHAKE_RTSXONXOFF = 3

'Event constants
Const MSCOMM_EV_SEND = 1
Const MSCOMM_EV_RECEIVE = 2
Const MSCOMM_EV_CTS = 3
Const MSCOMM_EV_DSR = 4
Const MSCOMM_EV_CD = 5
Const MSCOMM_EV_RING = 6
Const MSCOMM_EV_EOF = 7

'Error code constants
Const MSCOMM_ER_BREAK = 1001
Const MSCOMM_ER_CTSTO = 1002
Const MSCOMM_ER_DSRTO = 1003
Const MSCOMM_ER_FRAME = 1004
Const MSCOMM_ER_OVERRUN = 1006
Const MSCOMM_ER_CDTO = 1007
Const MSCOMM_ER_RXOVER = 1008
Const MSCOMM_ER_RXPARITY = 1009
Const MSCOMM_ER_TXFULL = 1010

Dim Unitel_Adresse As Structure_unite_adresse



Private Function car_RX(Index As Integer, tamp As String) As Integer
    
On Error GoTo 200

    ' Returns only ONE character from the received frame
    If Index <= Len(tamp$) Then
        car_RX = Asc(Mid$(tamp$, Index, 1))
    End If

    Exit Function
200 MsgBox "Erreur " & Str$(Err) & " module car_RX. Signalez immediatement cette erreur. Merci."
    Resume Next

End Function

Private Sub CMD_ecrit_Click()
Dim texte$
Static ref$

    ref$ = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890:;?!/.)àç_e-('""e&}]@^\`|[{#~¤*$"

    CMD_ecrit.Enabled = False
    
    ' On saisit le texte à entrer
    'Do
    '    texte$ = InputBox$("Saisissez un texte à ecrire (nombre PAIR de caracteres)", "Ecriture mots")
    'Loop While Len(texte$) Mod 2 <> 0
    
    TX_requete$ = Chr$(&H37) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(0) & Chr$(0) & Chr$(val(TXT_nb_mots.text)) & Chr$(0) & Left$(ref$, val(TXT_nb_mots.text) * 2)
    TX_requete_len% = Len(TX_requete$)
    TX_requete_ok% = True

    CMD_ecrit.Enabled = True

End Sub

Private Sub CMD_lit_Click()
    
    CMD_lit.Enabled = False
    
    TX_requete$ = Chr$(&H36) & Chr$(7) & Chr$(&H68) & Chr$(7) & Chr$(0) & Chr$(0) & Chr$(val(TXT_nb_mots.text)) & Chr$(0)
    TX_requete_len% = Len(TX_requete$)
    TX_requete_ok% = True

    CMD_lit.Enabled = True

End Sub

Private Sub CMD_requete_Click()

    TX_requete$ = Chr$(15) & Chr$(7)
    TX_requete_len% = 2
    TX_requete_ok% = True
    DoEvents
    TXT_RX.text = ""

End Sub



Private Sub envoie_requette()
Dim Index%, TX_tampon$
Dim car As String * 1
Dim chaine_utw$


    ' This routine sends the data frame
    If nb_timer_no_RX% = 0 Then
        

        ' Case for a request to be sent
        If TX_requete_len% <> 0 And TX_requete_ok% = True Then
            
            'Transfered to buffer-string
            RX_requete_ok% = 0

            'On compose le IDENT
            TX_tampon$ = Chr$(&HF1)   ' NPDU
            TX_tampon$ = TX_tampon$ & Chr$(Unitel_Adresse.Host_Station) & Chr$(&H0)   ' expéditeur
            TX_tampon$ = TX_tampon$ & Chr$(Unitel_Adresse.Remote_Station) & Chr$(&H0)  ' dest.
            ' N° de réseau émetteur
            TX_tampon$ = TX_tampon$ & Chr$(&H21) & Chr$(Unitel_Adresse.Host_Network)  ' expéditeur
            ' N° réseau destinataire
            TX_tampon$ = TX_tampon$ & Chr$(&H39) & Chr$(Unitel_Adresse.Remote_Network)  ' dest.
            
            TX_tampon$ = TX_tampon$ & Chr$(&HF9) & Chr$(0) & TX_requete$    ' requête
                
            ' 8) Sending...
            Winsock1.SendData Chr$(0) & Chr$(0) & Chr$(0) & Chr$(1) & Chr$(0) & Chr$(Len(chaine_utw$) + 1) & Chr$(0) & chaine_utw$

            'nb_timer_no_RX% = 0
            'etape_polling% = 1
            'TX_requete_ok% = False
        End If

    End If

End Sub

Public Sub Close_Com()
    
    ' Rend le n° de port
    'Ferme la connexion
    Winsock1.Close
    Do While Winsock1.State <> sckClosed
        DoEvents
    Loop

End Sub

Private Function Open_Com(adresse_ip_tsx As String) As Long
Dim IP_out As String

Début:
    With Winsock1
    'Connexion au serveur
        If .State <> sckConnected And .State <> sckOpen And .State <> sckClosing Then
            .LocalPort = 0
            
            If adresse_ip_tsx <> "" Then
                .RemotePort = 502
                .RemoteHost = adresse_ip_tsx
            End If
            .Connect
            DoEvents
            'Do
                ' C'est OK
                If .State = sckConnected Then
                    Open_Com = .LocalPort
                    Exit Function
                End If
                
                DoEvents
                
                If MsgBox("Imposible ouvir le port de communication à l'adresse : " & adresse_ip_tsx & " avec API", vbCritical + vbRetryCancel, "Communication API") = vbRetry Then
                    IP_out = InputBox("adresse : ", "Addresse de IP de l'API ?", adresse_ip_tsx)
                    If IP_out <> adresse_ip_tsx Then
                        Set_IP_API App.Path & "\Banc_XS.ini", IP_out
                    End If
                    adresse_ip_tsx = IP_out
                    GoTo Début
                Else
                    ' On laisse tomber
                    Open_Com = -2
                    Beep
                    Exit Function
                End If
            'Loopw
        Else
            ' Déjà ouvert
            Open_Com = .LocalPort
            MsgBox "Le port de communication avec l'API est déjas ouvert", vbInformation + vbOKOnly, "Communication API"
        End If
        
    End With
End Function

Private Sub Form_Load()
    
    kill_me = False
    ' Opening UNITELWAY
    'If Open_Com(Get_IP_API(App.Path & "\Banc_XS.ini")) < 0 Then
    If Open_Com("") < 0 Then
        MsgBox "Pas communication possible avec l'API" & vbCrLf & _
            "L'execution du programme vat être arrété!!", vbExclamation, App.Comments
         kill_me = True
        End
    End If
    
    'Unitel_Adresse = Get_Adresse_API(App.Path & "\Banc_XS.ini")
    Unitel_Adresse.Host_Network = 1
    Unitel_Adresse.Host_Station = 30
    Unitel_Adresse.Remote_Network = 1
    Unitel_Adresse.Remote_Station = 10
    
    
    TIM_polling.Enabled = True
    etape_polling% = 0
    adresse_scrutee% = val(TXT_adresse)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    kill_me = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' Closing UNITELWAY port
    TIM_polling.Enabled = False
    Close_Com

End Sub

Private Sub raz_parametres()
    
    nb_timer_no_RX% = 0
    RX_en_cours% = False
    tampon$ = ""

End Sub

Private Sub TIM_polling_Timer()
Dim boucle
Dim loose_time

 On Error GoTo 100

    ' To be adjusted
    'For loose_time = 1 To 10
    '    DoEvents
    'Next
    
    ' Buttons
     If TX_requete_ok% = True Then
         CHK_TX_valid.Value = 1
     Else
         CHK_TX_valid.Value = 0
     End If
    
     If RX_requete_ok% = True Then
         CHK_RX_valid.Value = 1
     Else
         CHK_RX_valid.Value = 0
     End If
    
     Select Case etape_polling%
    
        ' We must send a polling
         Case 0
             envoie_requette
        ' Waiting for RX
         Case 1
             'decode_trame_RX
    
     End Select
    
    'On fait retomber les flags de requete en cours
    If (TX_requete_ok% = False And RX_requete_ok% = True) Or code_erreur_utw <> 0 Then
        Lecture_en_cours = False
        Ecriture_en_cours = False
    End If

     Exit Sub
100 MsgBox "Erreur " & Str$(Err) & " module timer de TIM_polling. Signalez immediatement cette erreur. Merci."
     Resume Next

End Sub

Private Sub Time_out_Timer()
    com_err_tsx = 12
    etape_polling% = 0
    Time_out.Enabled = False
    RX_requete_ok% = False
    RX_requete_len% = 0
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim reponse$
    
    etape_polling% = 0
    Time_out.Enabled = False
    
    Winsock1.GetData RX_requete$, vbString
    RX_requete_ok% = False
    RX_requete_len% = 0

    
    
    If Len(reponse$) = 18 + bytesTotal Or Len(reponse$) = 20 + bytesTotal Then
        RX_requete$ = Right(RX_requete$, bytesTotal)
        RX_en_cours% = False
        RX_requete_ok% = True
        code_erreur_utw% = 0
    End If
    
    
    RX_requete_len% = bytesTotal

    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    com_err_tsx = 12
    etape_polling% = 0
    Time_out.Enabled = False
    RX_requete_ok% = False
    RX_requete_len% = 0
End Sub

Private Sub Winsock1_SendComplete()

    nb_timer_no_RX% = 0
    etape_polling% = 1
    TX_requete_ok% = False
    
    Time_out.Enabled = True
    
End Sub
