VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC814C0-FB17-4942-BF27-958161F0C351}#1.0#0"; "AfficheProj1.ocx"
Begin VB.Form FRM_controle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FINAL TEST BENCH XS"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10150.63
   ScaleMode       =   0  'User
   ScaleWidth      =   11806.21
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_traceability 
      Caption         =   "Traceability Check"
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin AfficheProj1.Affiche Affiche1 
      Height          =   3612
      Left            =   2760
      OleObjectBlob   =   "FRM_CONT.frx":0000
      TabIndex        =   9
      Top             =   1800
      Width           =   4092
   End
   Begin AfficheProj1.Affiche Affiche_prod_1 
      Height          =   3612
      Left            =   240
      OleObjectBlob   =   "FRM_CONT.frx":0107
      TabIndex        =   22
      Top             =   1800
      Width           =   2412
   End
   Begin AfficheProj1.Affiche Affiche_prod_2 
      Height          =   3612
      Left            =   6960
      OleObjectBlob   =   "FRM_CONT.frx":020D
      TabIndex        =   23
      Top             =   1800
      Width           =   2412
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   5400
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1440
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2880
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1440
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   6705
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5379
            MinWidth        =   5379
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5379
            MinWidth        =   5379
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6085
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   6960
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   847
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Cmd_prod2_mauvais 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "F11- Prod. 2 BAD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   1932
      End
      Begin VB.CommandButton Cmd_auto 
         Caption         =   "F11-AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5520
         TabIndex        =   14
         Top             =   120
         Width           =   1932
      End
      Begin VB.CommandButton CMD_retour 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "F7 - &END"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   1332
      End
      Begin VB.CommandButton CMD_Next_Step 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "F12 -Next Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1932
      End
      Begin VB.CommandButton Cmd_prod1_mauvais 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "F1- Prod. 1 BAD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   1932
      End
   End
   Begin Threed.SSFrame FRM_ref_produit 
      Height          =   660
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9372
      _Version        =   65536
      _ExtentX        =   16536
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "Product..."
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Lbl_mauvais 
         Caption         =   "QTY BAD product :"
         Height          =   252
         Left            =   6360
         TabIndex        =   8
         Top             =   360
         Width           =   2892
      End
      Begin VB.Label Lbl_bon 
         Caption         =   "QTY GOOD product :"
         Height          =   252
         Left            =   6360
         TabIndex        =   7
         Top             =   120
         Width           =   2880
      End
      Begin VB.Label LBL_no_de_lot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lot n° "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   252
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label LBL_symb_com 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4212
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   6120
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Label LBL_etat_cycle_claquage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   240
      TabIndex        =   20
      Top             =   5520
      Width           =   2412
   End
   Begin VB.Label LBL_pbon 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   7104
      TabIndex        =   19
      Top             =   792
      Width           =   2352
   End
   Begin VB.Label LBL_caract 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produit Number 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   7080
      TabIndex        =   18
      Top             =   1080
      Width           =   2352
   End
   Begin VB.Label LBL_pbon 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MAUVAIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   780
      Width           =   2196
   End
   Begin VB.Label LBL_caract 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produit Number 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label LBL_etat_cycle_ctrl_final 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   6960
      TabIndex        =   6
      Top             =   5520
      Width           =   2412
   End
   Begin VB.Label LBL_etat_cycle_chargement 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   3600
      TabIndex        =   5
      Top             =   5520
      Width           =   2292
   End
   Begin VB.Label LBL_limites 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Spec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4320
      TabIndex        =   3
      Top             =   1440
      Width           =   972
   End
   Begin VB.Label LBL_dialogue_operateur 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   2640
      TabIndex        =   0
      Top             =   705
      Width           =   4335
   End
   Begin VB.Shape SHP_voyant_vert 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   396
      Index           =   0
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2292
   End
   Begin VB.Shape SHP_voyant_vert 
      BackColor       =   &H0000FF00&
      BorderColor     =   &H00000000&
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   396
      Index           =   1
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2292
   End
End
Attribute VB_Name = "FRM_controle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Public WithEvents XsClient As Xs156Client
Attribute XsClient.VB_VarHelpID = -1


Private Sub Cmd_auto_Click()

Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
    
    If (Tab_Status_banc(0) And BIT0) <> 0 Then
        reset_bit Copy_Tab(0), 0
        Dim Com_erreure As Integer
            Com_erreure = -1
            Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
                Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 1)
                If Com_erreure <= 0 Then ' Communication en erreur
                    If MsgBox(ERREUR_ECRITURE_API, vbCritical + vbRetryCancel) = vbCancel Then
                        If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                            End
                        End If
                    End If
                Else ' Communication sans erreur
                    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                        Tab_Status_banc(i) = Copy_Tab(i)
                    Next i
                    'memo defaut precedent
                    previous_def_chargement = def_chargement
                    Exit Sub
                End If
            Loop
    End If
End Sub

Private Sub CMD_Next_Step_Click()
Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
    
    If CMD_Next_Step.Visible = True Then
       
        set_bit Copy_Tab(0), 1
        
        Dim Com_erreure As Integer
        Com_erreure = -1
        Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
            Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 1)
            If Com_erreure <= 0 Then ' Communication en erreur
                If MsgBox(ERREUR_ECRITURE_API, vbCritical + vbRetryCancel) = vbCancel Then
                    If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                        End
                    End If
                End If
            Else ' Communication sans erreur
                'memo defaut precedent
                For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                    Tab_Status_banc(i) = Copy_Tab(i)
                Next i
                previous_def_chargement = def_chargement
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub Cmd_prod1_mauvais_Click()
    If Cmd_prod1_mauvais.Visible = True Then
        del_P1 = True
    End If
End Sub

Private Sub Cmd_prod2_mauvais_Click()
    If Cmd_prod2_mauvais.Visible = True Then
       del_P2 = True
    End If
End Sub

Private Sub CMD_retour_click()
    '=========DEBUT MODIF BEEA 30/11/2007==========
    If ControleStationClaquageEnCours = False Then
        FRM_FIN_CONTROLE.Show
    Else
        FRM_FIN_TEST_CLAQUAGE.Show
    End If
    '=========FIN MODIF BEEA 30/11/2007============
End Sub

Private Sub cmd_traceability_Click()
On Error GoTo xsError
      MsgBox XsClient.ConnectionTest
    GoTo noError
xsError: MsgBox Error
noError:
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_F1 Then
    Cmd_prod1_mauvais_Click
    End If

    If KeyCode = KEY_F7 Or KeyCode = vbKeyDivide Then
        CMD_retour_click
    End If
    
    If KeyCode = KEY_F11 Then
    Cmd_prod2_mauvais_Click
    End If
    
    If KeyCode = KEY_F9 Then
    Cmd_auto_Click
    End If
    
    If KeyCode = KEY_F12 Or KeyCode = vbKeyDelete Then
        Call CMD_Next_Step_Click
    End If
End Sub

Private Function TrimAfterSpace(ByVal stringVal As String) As String
Dim l As Long
Dim stringResult As String
stringResult = Mid(stringVal, 1, 12)
l = InStr(stringVal, " ")
If l > 0 Then
    stringResult = Mid(stringVal, 1, l - 1)
End If
TrimAfterSpace = stringResult
End Function
Private Sub Form_Load()
Dim i
Dim xs As String
Dim xsWithoutGE As String
    LBL_symb_com.Caption = symb_com
    LBL_no_de_lot.Caption = "Lot n° " + Num_lot
    
    If Num_lot <> "" Then
        LBL_no_de_lot.Visible = True
        ''' Disini Insert Traceability
        On Error GoTo xsError
        If XsClient Is Nothing Then
           xs = "Set XS"
           Set XsClient = New Xs156Client
        End If
        If Not XsClient Is Nothing Then
            xs = "Load XS"
           If Not XsClient.LoadByOrderNumber(Num_lot) Then
           xs = "XS trim"
                xsWithoutGE = TrimAfterSpace(symb_com)
                xs = "StartNew"
                xs = TrimAfterSpace(Num_lot)
              XsClient.StartNewTrackingProcess xsWithoutGE, 48, Num_lot
           End If
        End If
        GoTo noError
xsError:  MsgBox "Loading Control : " & xs & " " & Error
noError:
    Else
        LBL_no_de_lot.Visible = False
    End If
    
    'Initialise l'ecran de contrôle
    SHP_voyant_vert(0).Visible = False
    LBL_pbon(0).Visible = False
    SHP_voyant_vert(1).Visible = False
    LBL_pbon(1).Visible = False
    Affiche_prod_1.Visible = True
    Affiche_prod_2.Visible = True
    CMD_Next_Step.Visible = False
    Cmd_prod1_mauvais.Visible = False
    Cmd_prod2_mauvais.Visible = False
    
    'Initialise les fins d'affichage de contrôle ou du résulatat de test station de claquage
    FinTestClaquageAffiche = False
    FinControleAffiche = False
    
    'init table echange PC<->TSX
    'Preparation mode dégradé
    If type_controle = DEGRADE Then
        'FRM_controle.Caption = "Banc XS : MODE DEGRADE"
        FRM_controle.Caption = "Bench XS : MODE DEGRADE"
    End If
        
    cycle_chargement_precedent = 0
    cycle_claquage_precedent = 0
    cycle_final_precedent = 0
    
        ' Préparation controle température
    If type_controle = TEMPERATURE_HAUTE Then
        'FRM_controle.Caption = "Banc XS : CONTROLE TEMPERATURE HAUTE"
        FRM_controle.Caption = "Bench XS : TEMPERATURE TEST HOT"
    End If
    If type_controle = TEMPERATURE_BASSE Then
        'FRM_controle.Caption = "Banc XS : CONTROLE TEMPERATURE BASSE"
        FRM_controle.Caption = "Bench XS : TEMPERATURE TEST COLD"
    End If
        ' Préparation controle complet
    If type_controle = CONTROLE_FINAL Then
        'FRM_controle.Caption = "Banc XS : CONTROLE FINAL"
        FRM_controle.Caption = "Bench XS : FINAL TEST"
        'Rafraichit les données en mémoire
        If DataEnvironment1.rsTable_LISTE_MODES.State <> adStateClosed Then
            DataEnvironment1.rsTable_LISTE_MODES.Close
        End If
        DataEnvironment1.rsTable_LISTE_MODES.Open
        
        DataEnvironment1.rsTable_LISTE_MODES.Filter = ("code_mode='" & ModeMarcheDielec & "'")
        If DataEnvironment1.rsTable_LISTE_MODES.RecordCount <> 0 Then
                FRM_controle.Caption = FRM_controle.Caption & " - MODE CLAQUAGE " & DataEnvironment1.rsTable_LISTE_MODES.Fields("Mode")
        End If
    End If
    
    ' Préparation premier contrôle
    If type_controle = PREMIER_CONTROLE Then
        'FRM_controle.Caption = "Banc XS : PREMIER CONTROLE"
        FRM_controle.Caption = "Bench XS : AFTER TRIM TEST"
    End If
    
    'affiche les valeurs limites
    Call maj_text_limites
    

    'Soyons sûrs que l'ecran a ete mis à jour
    'sauv_com_cycle = 999

    'setcursorpos 640, 480
    
    Me.WindowState = Mode_ouverture_fenetre
    
    Me.Show
    
    DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim result
 
    'Ferme les fichiers LOG
    Close
    FRM_startup!Lbl_code_lu.Caption = ""
'=======DEBUT MODIF BEEA 30/11/2007================
    Call LitModeTestStationClaquage
    Call FRM_startup.AfficheModeDielec
    Call FRM_startup.VerifSymboleTestClaquage
'=======FIN MODIF BEEA 30/11/2007==================
    FRM_startup.Show

End Sub

Private Sub XsClient_ExceptionEvent(ByVal guid As String)
Label1.Caption = guid
End Sub
