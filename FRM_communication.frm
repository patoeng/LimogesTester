VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Utwmait 
   Caption         =   "unitelway TCPIP communication"
   ClientHeight    =   5715
   ClientLeft      =   -90
   ClientTop       =   300
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.OptionButton Option_fonction_en_cours 
      Caption         =   "Function running"
      Height          =   252
      Left            =   5160
      TabIndex        =   12
      Top             =   3960
      Width           =   1572
   End
   Begin VB.OptionButton Option_socket_valide 
      Caption         =   "Socket OK"
      Height          =   252
      Left            =   5160
      TabIndex        =   11
      Top             =   3720
      Width           =   1332
   End
   Begin VB.OptionButton Option_socket_occupée 
      Caption         =   "Socket used"
      Height          =   192
      Left            =   5160
      TabIndex        =   10
      Top             =   3480
      Width           =   1572
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Read"
      Height          =   492
      Left            =   5520
      TabIndex        =   9
      Top             =   1920
      Width           =   1212
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Write"
      Height          =   492
      Left            =   5520
      TabIndex        =   8
      Top             =   960
      Width           =   1212
   End
   Begin VB.TextBox Text7 
      Height          =   372
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   3720
      Width           =   2652
   End
   Begin VB.TextBox Text6 
      Height          =   372
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   3240
      Width           =   2652
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1920
      Width           =   2892
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   960
      Width           =   2892
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   480
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1920
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   732
      Left            =   4920
      TabIndex        =   2
      Top             =   4440
      Width           =   1332
   End
   Begin VB.Timer Timer_mis_a_jour 
      Interval        =   50
      Left            =   4920
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start COM"
      Height          =   732
      Left            =   720
      TabIndex        =   0
      Top             =   4440
      Width           =   1212
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1920
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "API_Banc_XS"
      RemotePort      =   502
   End
   Begin VB.Label Label4 
      Caption         =   "Local host"
      Height          =   252
      Index           =   3
      Left            =   2400
      TabIndex        =   21
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Local IP"
      Height          =   252
      Index           =   2
      Left            =   2400
      TabIndex        =   20
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Local port"
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Status socket"
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   18
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "mm"
      Height          =   252
      Index           =   1
      Left            =   4200
      TabIndex        =   17
      Top             =   3840
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "mm"
      Height          =   252
      Index           =   0
      Left            =   4200
      TabIndex        =   16
      Top             =   3240
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Rule 2"
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   3720
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Rule 1"
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Régle"
      Height          =   132
      Left            =   840
      TabIndex        =   13
      Top             =   2040
      Width           =   12
   End
End
Attribute VB_Name = "Utwmait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Compt As Integer 'varial interne pour test écriture

Private Sub Command1_Click()
    Debut_communication
    Compt = 0
End Sub

Private Sub Command2_Click()
    Fin_communication
End Sub

Public Sub Events()
    DoEvents
    Timer_mis_a_jour_Timer
End Sub

Public Sub Envoie_UTW(requete As String)
Dim chaine_utw$
Dim I As Integer
    
    If Not Socket_valid Then
        Exit Sub 'pas d'envoie de commande de communication à API si le socket est déclarée comme non valide
    End If
        
    If Winsock1.State <> sckConnected Then
        Winsock1.Close
        While Winsock1.State = sckClosing
            DoEvents
        Wend
        Winsock1.Connect
        While Winsock1.State <> sckConnected
            DoEvents
        Wend
    End If
    
    Unload Com_erreure
       
    'On compose le IDENT
    chaine_utw$ = Chr$(&HF1)   ' NPDU
    chaine_utw$ = chaine_utw$ & Chr$(Unitel_Adresse.Host_Station) & Chr$(&H0)   ' expéditeur
    chaine_utw$ = chaine_utw$ & Chr$(Unitel_Adresse.Remote_Station) & Chr$(&H0)  ' dest.
    ' N° de réseau émetteur
    chaine_utw$ = chaine_utw$ & Chr$(&H21) & Chr$(Unitel_Adresse.Host_Network)  ' expéditeur
    ' N° réseau destinataire
    chaine_utw$ = chaine_utw$ & Chr$(&H39) & Chr$(Unitel_Adresse.Remote_Network)  ' dest.
        
    chaine_utw$ = chaine_utw$ & Chr$(&HF9) & Chr$(0) & requete$    ' requête
        
    I = Len(chaine_utw$)
        
    Winsock1.SendData Chr$(0) & Chr$(0) & Chr$(0) & Chr$(1) & Chr$(0) & Chr$((I + 1) Mod 256) & Chr$(Int((I + 1) / 256)) & chaine_utw$
    Socket_occupé = True
End Sub

Private Sub Command5_Click()
Dim aze(260) As Long
Dim I As Integer
Dim num_res As Long

    Command5.Enabled = False
            
    For I = 0 To UBound(aze)
        If Compt + I > 25000 Then Compt = 0
        aze(I) = Compt + I
    Next I
    Compt = Compt + 1
    num_res = Ecrir_mots(6500, aze, 115, 0)
    Command5.Enabled = True
End Sub

Private Sub Timer_mis_a_jour_Timer()

    Text1.Text = Winsock1.State
    Text2.Text = Winsock1.LocalPort
    Text3.Text = Winsock1.LocalIP
    Text4.Text = Winsock1.LocalHostName
    
    Option_socket_occupée.Value = Socket_occupé
    Option_socket_valide.Value = Socket_valid
    Option_fonction_en_cours.Value = Fonction_com_en_cours

Dim num_mes As Long

Dim regle_1 As Double
Dim regle_2 As Double

    If Not Socket_valid Then
        Exit Sub 'pas d'envoie de commande de communication à API si le socket est déclarée comme non valide
    End If
    
    If Socket_occupé Or Fonction_com_en_cours Then
        Exit Sub
    End If
    
    If Winsock1.State = sckConnected Then
        'cycle = 1
        Select Case cycle_mise_a_jour_données_API
            Case 0
               num_mes = Lecture_etat_banc()
               
            Case 1
                num_mes = Lecture_position_rêgle(regle_1, regle_2)
                If num_mes > 0 Then
                    Text6.Text = regle_1
                    Text7.Text = regle_2
                Else
                    Text6.Text = "#####"
                    Text7.Text = "#####"
                End If
        End Select
        DoEvents
        cycle_mise_a_jour_données_API = cycle_mise_a_jour_données_API + 1
        If cycle_mise_a_jour_données_API > 2 Then
            cycle_mise_a_jour_données_API = 0
        End If
    Else
        If Winsock1.State = sckClosed Then
                Winsock1.Connect
        End If
        If Winsock1.State = sckClosing Or Winsock1.State = sckError Then
                Winsock1.Close
        End If
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim strData As String

Dim I As Integer
Dim j As Integer

    Winsock1.GetData strData, vbString
    
    Select Case Commende_com.Commende
        Case Trame_commende.Fonction_Lecture_etat_banc
            j = 0
            I = 1
            Commende_com.Nombre_mots = (bytesTotal - 21) / 2
            For I = 21 To bytesTotal Step 2
                If I + 1 <= bytesTotal Then
                    Tab_Status_banc(j) = (Asc(Mid(strData, I, 1)) * 1) + (CLng(Asc(Mid(strData, I + 1, 1))) * 256)
                    j = j + 1
                End If
            Next I
        Case Trame_commende.Fonction_Lecture_position_régle
            j = 1
            I = 1
            Commende_com.Nombre_mots = (bytesTotal - 21) / 2
            For I = 21 To bytesTotal Step 2
                If I + 1 <= bytesTotal Then
                    Commende_com.Trame(j) = (Asc(Mid(strData, I, 1)) * 1) + (CLng(Asc(Mid(strData, I + 1, 1))) * 256)
                    j = j + 1
                End If
            Next I
            j = j + 0
        Case Trame_commende.Fonction_Lecture_paramétre_produit
            j = 1
            I = 1
            Commende_com.Nombre_mots = (bytesTotal - 21) / 2
            For I = 21 To bytesTotal Step 2
                If I + 1 <= bytesTotal Then
                    Commende_com.Trame(j) = (Asc(Mid(strData, I, 1)) * 1) + (CLng(Asc(Mid(strData, I + 1, 1))) * 256)
                    j = j + 1
                End If
            Next I
            symb_com = Mid$(strData, 21, 16)
            Ref_indus = Mid$(strData, 37, 16)
            Num_lot = Mid$(strData, 53, 8)
            
        Case Trame_commende.Fonction_Lecture_messages_produit_1
            j = 1
            I = 1
            Commende_com.Nombre_mots = (bytesTotal - 21) / 2
            For I = 21 To bytesTotal Step 2
                If I + 1 <= bytesTotal Then
                    Commende_com.Trame(j) = (Asc(Mid(strData, I, 1)) * 1) + (CLng(Asc(Mid(strData, I + 1, 1))) * 256)
                    j = j + 1
                End If
            Next I
        Case Trame_commende.Fonction_Lecture_messages_produit_2
            j = 1
            I = 1
            Commende_com.Nombre_mots = (bytesTotal - 21) / 2
            For I = 21 To bytesTotal Step 2
                If I + 1 <= bytesTotal Then
                    Commende_com.Trame(j) = (Asc(Mid(strData, I, 1)) * 1) + (CLng(Asc(Mid(strData, I + 1, 1))) * 256)
                    j = j + 1
                End If
            Next I
        Case Trame_commende.Fonction_Ecriture_mots
            If bytesTotal < 19 Then
                Commende_com.Nombre_mots = 0
            End If
            Socket_occupé = False
        Case Else
            Socket_occupé = False
    End Select

    Socket_occupé = False
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Text5.text = Error & Description
    Com_erreure.Message.Caption = Description
    Com_erreure.Show
End Sub
