VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_DEGRADE 
   Caption         =   "Mode degrade"
   ClientHeight    =   6000
   ClientLeft      =   660
   ClientTop       =   1335
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   9
      Top             =   5748
      Width           =   6024
      _ExtentX        =   10636
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   " ENTER for OK, ESC for Cancel"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Not"
      Height          =   252
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   1452
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Not"
      Height          =   252
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Width           =   1572
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hot Temperature"
      Height          =   252
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   1932
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Final"
      Height          =   252
      Index           =   3
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   1932
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "After Trim"
      Height          =   252
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   2520
      Width           =   1932
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Active"
      Height          =   252
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   1452
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Active"
      Height          =   252
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1572
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FF00&
      Caption         =   "Pas à pas"
      Height          =   252
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   3840
      Width           =   1212
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cold Temperature"
      Height          =   252
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   1932
   End
   Begin VB.Image Image1 
      Height          =   5628
      Left            =   0
      Picture         =   "Frm_degr.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5628
   End
End
Attribute VB_Name = "FRM_DEGRADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Judul_Pesan As String = "Manual Operation Mode"

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1(1).Visible = True
    Else
        Check1(1).Visible = False
    End If
End Sub

Private Sub Check3_Click()

    If Check3.Value = 1 Then
        Check1(2).Visible = True
    Else
        Check1(2).Visible = False
    End If
    Option1(0).Visible = Check1(2).Visible
    Option1(1).Visible = Check1(2).Visible
    Option1(2).Visible = Check1(2).Visible And si_ctrl_auto_adaptatif
    Option1(3).Visible = Check1(2).Visible
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Commande_annuler
    End If
    
    If KeyCode = vbKeyReturn Then
        Commande_valider
    End If
End Sub

Private Sub Form_Load()
    'affichage_dégradé = True
    'out_degrade = False
    
    Form_Resize
    
    Check1(1).Visible = False
    Check1(2).Visible = False
    
    Option1(0).Visible = Check1(2).Visible
    Option1(1).Visible = Check1(2).Visible
    Option1(2).Visible = Check1(2).Visible And si_ctrl_auto_adaptatif
    Option1(3).Visible = Check1(2).Visible
    
    'met la fenetre au premier plan, et on ne peut plus l'effacer !
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
End Sub

Private Sub Form_Paint()
    'met la fenetre au premier plan, et on ne peut plus l'effacer !
    'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
End Sub

Private Sub Form_Resize()
    Me.Width = Image1.Width + 50
    Me.Height = Image1.Height + 400
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Double
Dim q As Double

Dim a As OptionButton

Dim Module As Double
Dim Angle As Double

    i = CDbl(x) - CDbl(Image1.Width / 2)
    q = CDbl(y) - CDbl(Image1.Height / 2)
    
    Module = Sqr((i * i) + (q * q))
    If q <> 0 Then
        Angle = Atn(i / q)
    End If
    
    If Module > Image1.Width / 2 Then
        If q > 0 Then
            If Angle > 0 Then 'Command valider
                Commande_valider
            Else ' Commande Annuler
                Commande_annuler
            End If
        End If
    End If
End Sub

Private Sub Commande_valider()
Dim Mode_chargement As Long
Dim Mode_dielectric As Long
Dim Mode_final As Long
Dim Cycle As Integer

Dim Cycle_BD As Integer

Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i

Dim a As OptionButton

    Mode_chargement = Mode_chargement + BIT8 'toujours actif
    If Check1(0).Value = CHECKED Then
        Mode_chargement = Mode_chargement + BIT9 'pas à pas
    End If
    If Check1(1).Value = CHECKED Then
        Mode_dielectric = Mode_dielectric + BIT9 'pas à pas
    End If
    If Check1(2).Value = CHECKED Then
        Mode_final = Mode_final + BIT9 'pas à pas
    End If
    
    If Check3.Value = CHECKED Then
        Mode_final = Mode_final + BIT8 ' cycle final actif
        For Each a In Option1
            If a.Value = True Then
                Cycle = a.Index + 1
            End If
        Next
        If Cycle = 0 Then
            'MsgBox "Choisisez un type cycle", vbInformation, "Mode de marche manuel"
            MsgBox "Please choose one cycle type", vbInformation, Judul_Pesan
            Exit Sub
        End If
        Mode_final = Mode_final + Cycle
    End If
    
    If Check2.Value = CHECKED Then
        Mode_dielectric = Mode_dielectric + BIT8 'cycle dielectric actif
    End If
    
    'Vérifie qu'il y a au moins un poste activée
    If Not Check2.Value = CHECKED And Not Check3.Value = CHECKED Then
        'MsgBox "Activer au moins un post", vbInformation, "Mode de marche manuel"
        MsgBox "Activated at least one station", vbInformation, Judul_Pesan
        Exit Sub
    End If
    
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
    
    Select Case Cycle
        Case API_PREMIER_CONTROLE
            Cycle_BD = CYCLE_PREMIER_CONTROLE
            type_controle = PREMIER_CONTROLE
        Case API_TEMPERATURE_HAUTE
            Cycle_BD = CYCLE_TEMPERATURE_HAUTE
            type_controle = TEMPERATURE_HAUTE
        Case API_TEMPERATURE_BASSE
            Cycle_BD = CYCLE_TEMPERATURE_BASSE
            type_controle = TEMPERATURE_BASSE
        Case API_CONTROLE_FINAL
            Cycle_BD = CYCLE_CONTROLE_FINAL
            type_controle = CONTROLE_FINAL
        Case Else
            Cycle_BD = CYCLE_CONTROLE_FINAL
            type_controle = CONTROLE_FINAL
    End Select
    
    If Générer_sequence_API_contrôle_pos_final(Code_produit_badge, Cycle_BD) < 0 Then
        'Défaut de com ou de base de donnée
        Commande_annuler
        Exit Sub
    End If
    
    Commande_mode_de_marche_OK = True
    Commande_mode_de_marche_SET = True
    
    Unload Me
End Sub

Private Sub Commande_annuler()
    Commande_mode_de_marche_OK = False
    Commande_mode_de_marche_SET = True
    Unload Me
End Sub
