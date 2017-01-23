VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FRM_startup 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   8265
   ClientLeft      =   0
   ClientTop       =   -150
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FRM_STAR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTestStationClaquage 
      Caption         =   "Dummy Test Hipot"
      Height          =   3135
      Left            =   6960
      TabIndex        =   22
      Top             =   1440
      Width           =   2535
      Begin VB.CommandButton cmdTestClaquage 
         Caption         =   "F4 - DUMMY TEST HIPOT"
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   2200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblDemandeTestClaquage 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dummy test for Hipot station : use product reference 'XS1D..CLAQUAGE'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1400
         Left            =   140
         TabIndex        =   23
         Top             =   400
         Width           =   2300
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer timRefreshAffClaqu 
      Interval        =   2000
      Left            =   6390
      Top             =   2880
   End
   Begin VB.CommandButton cmdParamTestClaquage 
      Caption         =   "C - Parameter DUMMY TEST HIPOT"
      Height          =   615
      Left            =   5400
      TabIndex        =   25
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Timer timTestClaqu 
      Interval        =   60000
      Left            =   6390
      Top             =   2340
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3720
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListEmpreinte 
      Left            =   600
      Top             =   4560
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   461
      ImageHeight     =   192
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_STAR.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_STAR.frx":16240
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_STAR.frx":2C176
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FRA_etiquette 
      Caption         =   "Print error label"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   5175
      Begin VB.CommandButton Cmd_etiquette 
         Caption         =   "CTRL F12 - &No"
         Height          =   495
         Index           =   2
         Left            =   3600
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_etiquette 
         Caption         =   "F11 - All &Product"
         Height          =   495
         Index           =   1
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_etiquette 
         Caption         =   "F10 - &Normal"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Lbl_etiquette 
         Alignment       =   2  'Center
         Caption         =   "Print label for fail product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "F5 - &Autotest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   7080
      TabIndex        =   10
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton CMD_AFF_PARAM 
      Caption         =   "F6 - &Parameter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   8
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton CMD_version 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "F7 - A&bout"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   1815
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   3135
      Left            =   6960
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "Start test..."
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.CommandButton Cmd_ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F1 - After Trim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   4
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2325
      End
      Begin VB.CommandButton Cmd_ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F4 - &Mode degrade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   0
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2325
      End
      Begin VB.CommandButton Cmd_ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F3 - &Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2325
      End
      Begin VB.CommandButton Cmd_ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "F2 - &Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   20
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   2325
      End
   End
   Begin VB.CommandButton CMD_quitter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ESC - &Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1815
   End
   Begin Threed.SSPanel Panel3D1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9660
      _Version        =   65536
      _ExtentX        =   17039
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Final && Dielectric Test XS"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   24.01
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   6
      BevelOuter      =   0
      BevelInner      =   1
      Font3D          =   4
   End
   Begin Threed.SSPanel SSP_appel_maint 
      Height          =   525
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   926
      _StockProps     =   15
      Caption         =   "F9-Maintenance call"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Outline         =   -1  'True
      Enabled         =   0   'False
   End
   Begin Threed.SSCommand Cmd_appel_maint 
      Height          =   540
      Left            =   0
      TabIndex        =   11
      Top             =   7320
      Width           =   2205
      _Version        =   65536
      _ExtentX        =   3889
      _ExtentY        =   953
      _StockProps     =   78
      Caption         =   "F9-&Appel maintenance"
   End
   Begin Threed.SSCommand CMD_sauve_log 
      Height          =   732
      Left            =   7080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6132
      Visible         =   0   'False
      Width           =   2292
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "F8 - Copy log file"
      Picture         =   "FRM_STAR.frx":420AC
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label lblModeClaquage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HIPOT RUNNING MODE :"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2640
      TabIndex        =   27
      Top             =   1500
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label lblModeDielec 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblModeDielec"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2175
      TabIndex        =   26
      Top             =   1725
      Visible         =   0   'False
      Width           =   4530
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Put the support"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2520
      TabIndex        =   19
      Top             =   3000
      Width           =   3612
   End
   Begin VB.Image ImageEmpreinte 
      Height          =   1572
      Left            =   2520
      Picture         =   "FRM_STAR.frx":423C6
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   3612
   End
   Begin VB.Label Lbl_code_lu 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Lbl_code_prod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Product reference"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   2175
      Width           =   3000
   End
   Begin VB.Label Lbl_info_type_controle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choose testing type by Pressing [F1] or  [F2] or [F3] or [F4]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   336
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   5976
   End
End
Attribute VB_Name = "FRM_startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Private Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer

Private Sub CMD_AFF_PARAM_Click()
    If ((symbole_existe = False) Or (nouvelle_barquette = True)) Then
        Exit Sub
    End If
    aff_param = 1
    Contexte$ = "AFFICHAGE"
End Sub

Private Sub Cmd_appel_maint_Click()
Dim i

Dim a As Integer
Dim Copy_Tab(20) As Long

    For a = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(a) = Tab_Status_banc(i)
    Next a
    
    If appel_maintenance = True Then
        reset_bit Copy_Tab(4), 9
        appel_maintenance = False
        SSP_appel_maint.BackColor = &HC0C0C0
    Else
        set_bit Copy_Tab(4), 9
        appel_maintenance = True
        SSP_appel_maint.BackColor = &H80FF&
    End If
    
Dim Com_erreure As Integer
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 1, 4)
        If Com_erreure <= 0 Then ' Communication en erreur
            If MsgBox(ERREUR_ECRITURE_APPELE_MAINTENANCE, vbCritical + vbRetryCancel) = vbCancel Then
                If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                    End
                End If
            End If
        Else ' Communication sans erreur
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            Exit Sub
        End If
    Loop
End Sub

Private Sub Cmd_etiquette_Click(Index As Integer)
    If Index = 0 Then
        forçage_etiquette = 0
        sans_etiquette = 0
        'Lbl_etiquette.Caption = "Impression etiquette sur défaut produit"
        Lbl_etiquette.Caption = "Printing label for fail product"
    Else
        If Index = 1 Then
            forçage_etiquette = 1
            sans_etiquette = 0
            'Lbl_etiquette.Caption = "Impression etiquette pour tous les produits"
            Lbl_etiquette.Caption = "Printing label for all product"
        Else
            If Index = 2 Then
            forçage_etiquette = 0
            sans_etiquette = 1
            'Lbl_etiquette.Caption = "Aucune impression d'etiquette"
            Lbl_etiquette.Caption = "No label printing"
            End If
        End If
    End If
End Sub

Private Sub CMD_ok_Click(Index As Integer)
' Charge les donnees et les prgrammes de contrôle
Dim ComErreure As Long

Dim Mode_chargement As Long
Dim Mode_dielectric As Long
Dim Mode_Deg_Diel As Long

Dim Mode_final As Long
Dim Cycle As Integer

Dim Com_erreure As Integer

Dim index_produit

Dim i As Integer
Dim Copy_Tab(20) As Long

    If ((symbole_existe = False) Or (nouvelle_barquette = True)) Then Exit Sub

    ' si le produit n'est pas autoadaptable le premier contrôl n'est pas utilisabe
    If Not si_ctrl_auto_adaptatif And Index = PREMIER_CONTROLE Then
        'MsgBox "Fonction non utilisable sur ce symbole", vbOKOnly + vbInformation + vbApplicationModal
        MsgBox "Function not available for this reference", vbOKOnly + vbInformation + vbApplicationModal
        Exit Sub
    End If
    
    ' Jeu de donnees
    type_controle = Index   ' !!!
    
    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
    
    Select Case Index
        Case PREMIER_CONTROLE
            Com_erreure = COMMUNICATION_OCCUPE 'initialise le code erreure de com avec une erreure pour rentrer dans le boucle while
            While Com_erreure <= 0
                Mode_chargement = BIT8 'toujours actif
                Mode_dielectric = 0 'toujours inatif
                Mode_final = API_PREMIER_CONTROLE + BIT8
                
                Copy_Tab(10) = Mode_chargement
                Copy_Tab(11) = Mode_dielectric
                Copy_Tab(12) = Mode_final
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
            Wend
            
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            
            'on refait une lecture des infos du tsx avant de faire l'acquis defaut, car des choses se sont produites entre temps...
            Lecture_etat_banc

            If Générer_sequence_API_contrôle_pos_final(Code_produit_badge, CYCLE_PREMIER_CONTROLE) < 0 Then
                'Défaut de com ou de base de donnée
                Exit Sub
            End If
            
            demande_station = 0
            lance_controle = 1
            Contexte$ = "CONTROLE"
        
        Case CONTROLE_FINAL
            Com_erreure = COMMUNICATION_OCCUPE 'initialise le code erreure de com avec une erreure pour rentrer dans le boucle while
            While Com_erreure <= 0
                Mode_chargement = BIT8 'toujours actif
                Mode_Deg_Diel = 0
                Select Case ModeMarcheDielec
                    Case 0, 1
                        Mode_Deg_Diel = 0
                    Case 2
                        Mode_Deg_Diel = BIT11
                    Case 3
                        Mode_Deg_Diel = BIT12
                    Case Else
                        Mode_Deg_Diel = BIT11 + BIT12
                End Select
                Mode_dielectric = BIT8 + Mode_Deg_Diel 'toujours actif + dégardés
                Mode_final = API_CONTROLE_FINAL + BIT8
                
                Copy_Tab(10) = Mode_chargement
                Copy_Tab(11) = Mode_dielectric
                Copy_Tab(12) = Mode_final
                Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 3, 10)
                If Com_erreure <= 0 Then
                    If MsgBox(ERREUR_ECRITURE_MODE_DE_MARCHE, vbCritical + vbRetryCancel) = vbCancel Then
                        If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                            End
                        End If
                    End If
                    Exit Sub
                
                End If
            Wend
            
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            
            'on refait une lecture des infos du tsx avant de faire l'acquis defaut, car des choses se sont produites entre temps...
            Lecture_etat_banc
            
            If Générer_sequence_API_contrôle_pos_final(Code_produit_badge, CYCLE_CONTROLE_FINAL) < 0 Then
                'Défaut de com ou de base de donnée
                Exit Sub
            End If
            
            demande_station = 0
            lance_controle = 1
            Contexte$ = "CONTROLE"
          
        Case TEMPERATURE
            Commande_mode_de_marche_OK = False
            Unload FRM_choix_température
            Commande_mode_de_marche_SET = False
            FRM_choix_température.Show
            While Not Commande_mode_de_marche_SET
                DoEvents
            Wend
            'on refait une lecture des infos du tsx avant de faire l'acquis defaut, car des choses se sont produites entre temps...
            Lecture_etat_banc
            If Commande_mode_de_marche_OK = False Then
                Exit Sub
            End If
            If Commande_mode_de_marche_OK = False Then
                Exit Sub
            End If
            
            demande_station = 0
            lance_controle = 1
            Contexte$ = "CONTROLE"
            
        Case DEGRADE
            Commande_mode_de_marche_OK = False
            Unload FRM_DEGRADE
            Commande_mode_de_marche_SET = False
            FRM_DEGRADE.Show
            'on refait une lecture des infos du tsx avant de faire l'acquis defaut, car des choses se sont produites entre temps...
            Lecture_etat_banc
            While Not Commande_mode_de_marche_SET
                DoEvents
            Wend
            If Commande_mode_de_marche_OK = False Then
                Exit Sub
            End If
            demande_station = 0
            lance_controle = 1
            Contexte$ = "CONTROLE"
            
        Case AUTO_TEST
            lance_autotest = 1
            out_autotest = 0
            Contexte$ = "AUTOTEST"
            
        Case TEST_STA_CLAQU
            Dim Cycle_BD As Integer
        
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Copy_Tab(i) = Tab_Status_banc(i)
            Next i
            
            Mode_chargement = Mode_chargement + BIT8 'toujours actif
            Mode_dielectric = Mode_dielectric + BIT8 + BIT10 'cycle dielectric actif avec mode test de la station de claquage
          
            Copy_Tab(10) = Mode_chargement
            Copy_Tab(11) = Mode_dielectric
            Copy_Tab(12) = Mode_final
            
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
            
            If Générer_sequence_API_contrôle_pos_final(Code_produit_badge, API_CONTROLE_FINAL) < 0 Then
                'Défaut de com ou de base de donnée
                Exit Sub
            End If
    
            'on refait une lecture des infos du tsx avant de faire l'acquis defaut, car des choses se sont produites entre temps...
            Lecture_etat_banc
            demande_station = 0
            lance_controle = 1
            Contexte$ = "CONTROLE"
    End Select
   
End Sub

Private Sub CMD_quitter_Click()

    Unload Utwmait

'Enregistrement de l'arrêt de l'application dans l'historique
    Open file_hist_def$ For Append As #3
    Print #3, "------------------------------    "; Chr$(13); Chr$(10);
    'Print #3, "ARRET DU PROGRAMME DE CONTROLE le "; Str(Date); " à "; Str(Time); Chr$(13); Chr$(10);
    Print #3, "START TEST at "; Str(Date); " à "; Str(Time); Chr$(13); Chr$(10);
    Print #3, "------------------------------    "; Chr$(13); Chr$(10);
    Print #3, Chr$(13); Chr$(10);
    Print #3, Chr$(13); Chr$(10);
    Close #3
    
    End
   
End Sub

Private Sub CMD_sauve_log_Click()

On Error GoTo ErrHandler

'Const TITRE_DIALOGUE_TRANSFAIRE = "Transfaire des données de la base de donnée journal"
Const TITRE_DIALOGUE_TRANSFAIRE = "Transfer data from database log"

Dim fs As New Scripting.FileSystemObject 'object system
Dim f  As Scripting.File  'object fichier

Dim Fichier_source As String
Dim Fichier_final As String

Dim Text As String
Dim i As Long

    'Ouvre les tables si naisére
    If DataEnvironment1.rsTable_Log.State = adStateClosed Then
        DataEnvironment1.rsTable_Log.Open
    End If
    If DataEnvironment1.rsTable_Log_Banc.State = adStateClosed Then
        DataEnvironment1.rsTable_Log_Banc.Open
    End If
    
    'Ouvre une boite de dialogue pour choisir une le fichier final
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = file_save
    'CommonDialog1.DialogTitle = "Transfaire données jounal"
    CommonDialog1.DialogTitle = "Transfer data log"
    CommonDialog1.Filter = "All file (*.*)|*.*|Access file (*.mdb)|*.mdb"
    CommonDialog1.DefaultExt = ".mdb"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowSave
      
    'Cherche le fichier source
    Fichier_source = DataEnvironment1.Data_Log.Properties.Item("Data Source Name").Value
    
    'Vérie si le fichier source existe
    If Not fs.FileExists(Fichier_source) Then
        'MsgBox " Le fichier source" & vbCrLf & Fichier_source & vbCrLf & "n'exite pas", vbCritical, TITRE_DIALOGUE_TRANSFAIRE
        MsgBox "Source file " & vbCrLf & Fichier_source & vbCrLf & "not exist", vbCritical, TITRE_DIALOGUE_TRANSFAIRE
        Exit Sub
    End If
    
    'Copy le fichier source vers le fichier final
    Fichier_final = CommonDialog1.FileName
    Set f = fs.GetFile(Fichier_source)
    f.Copy Fichier_final
      
    'Vide la base de donnée journal si l'utilisateur valide le bouton "OUI"
    'If (MsgBox("Souhaitez vous vider la base de donnée journal", vbQuestion + vbYesNo + vbDefaultButton2, TITRE_DIALOGUE_TRANSFAIRE) = vbYes) Then
    If (MsgBox("Do you want to empty the database log?", vbQuestion + vbYesNo + vbDefaultButton2, TITRE_DIALOGUE_TRANSFAIRE) = vbYes) Then
       
       DataEnvironment1.Data_Log.Execute "DELETE Log.* FROM Log"
       DataEnvironment1.Data_Log.Execute "DELETE Banc.* FROM Banc"
       
        'For i = 1 To DataEnvironment1.rsTable_Log.RecordCount
        '    DataEnvironment1.rsTable_Log.MoveFirst
        '    DataEnvironment1.rsTable_Log.Delete
        'Next i
        
        
        'For i = 1 To DataEnvironment1.rsTable_Log_Banc.RecordCount
        '    DataEnvironment1.rsTable_Log_Banc.MoveFirst
        '    DataEnvironment1.rsTable_Log_Banc.Delete
        'Next i
    End If
     
    Exit Sub
    
ErrHandler:
        'L'utilisateur a cliqué sur Annuler
        Exit Sub
        
End Sub

Public Sub CMD_Vide_base_donnee()

On Error GoTo ErrHandler

Const MAX_TAILLE_FICHIER_LOG = 5000000 '5 Méga octé
Const FICHIER_AURIGINE = "BDD XS Vide.mdb"

Dim fs As New Scripting.FileSystemObject 'object system
Dim f  As Scripting.File  'object fichier

Dim Fichier_source As String
Dim Fichier_final As String
Dim Fichier_auri As String

Dim Active_connection_log As String
Dim Active_connection_banc_log As String

Dim Text As String
Dim Date_text As String
Dim i As Long

Dim aze As Object

Dim Table_Log_LockType As LockTypeEnum
Dim Table_Log_CursorType As CursorTypeEnum
Dim Table_Log_Source As String

Dim Table_Log_Banc_LockType As LockTypeEnum
Dim Table_Log_Banc_CursorType As CursorTypeEnum
Dim Table_Log_Banc_Source As String

Dim Requette_update_LockType As LockTypeEnum
Dim Requette_update_CursorType As CursorTypeEnum
Dim Requette_update_Source As String

    'Cherche le fichier source
    Fichier_source = DataEnvironment1.Data_Log.Properties.Item("Data Source Name").Value
    
    'Défini le non du nouveau fichier
    Text = Now
    Date_text = ""
    For i = 1 To Len(Text)
        Select Case Mid(Text, i, 1)
            Case ":"
                Date_text = Date_text & "_"
            Case "\"
                Date_text = Date_text & "_"
             Case "/"
                Date_text = Date_text & "_"
            Case "."
                Date_text = Date_text & "_"
            Case Else
                Date_text = Date_text & Mid(Text, i, 1)
        End Select
    Next i
1     Fichier_final = fs.GetParentFolderName(Fichier_source) & "\" & fs.GetBaseName(Fichier_source) & "_" & Date_text & "." & fs.GetExtensionName(Fichier_source)
    
    'Copy le fichier source vers le fichier final
2     Set f = fs.GetFile(Fichier_source)
    
3     If (f.Size > MAX_TAILLE_FICHIER_LOG) Then
        
        'MsgBox "La base de donnée Log" & vbCrLf & "a dépacé sa taille limite." & _
        '    vbCrLf & "Le programme va vider la base de donnée Log" & vbCrLf _
        '    & "et metre fin à son execution" & vbCrLf & "Relancer le programme de controle", vbInformation + vbOKOnly
        MsgBox "Database log " & vbCrLf & "length already at limit." & _
            vbCrLf & "Program will empty the database log" & vbCrLf _
            & "and quit the program" & vbCrLf & "Restart control program", vbInformation + vbOKOnly
        
           
        'Ouvre les tables si naisére
4        If DataEnvironment1.rsTable_Log.State <> adStateClosed Then
5            DataEnvironment1.rsTable_Log.Close
        End If
6        If DataEnvironment1.rsTable_Log_Banc.State <> adStateClosed Then
7            DataEnvironment1.rsTable_Log_Banc.Close
        End If
8        If DataEnvironment1.rsrequette_update.State <> adStateClosed Then
9            DataEnvironment1.rsrequette_update.Close
        End If
        
10        DataEnvironment1.Data_Log.Close
    
    
11        f.Copy Fichier_final
12        f.Delete
          
        'Copy le fichier aurigine vers le fichier source
13        Fichier_auri = fs.BuildPath(fs.GetParentFolderName(Fichier_source), FICHIER_AURIGINE)
14        Set f = fs.GetFile(Fichier_auri)
15        f.Copy Fichier_source
          
16        Set fs = CreateObject("Scripting.FileSystemObject")
        
17        fs.CreateTextFile App.Path & "/Relance.flag"
          
        End
        
    End If
Exit Sub

ErrHandler:
    'MsgBox err.Number & vbCrLf & err.Description & vbCrLf & "Ligne :" & Erl
    MsgBox err.Number & vbCrLf & err.Description & vbCrLf & "Line :" & Erl
    Resume Next
    'Exit Sub
End Sub

Private Sub CMD_version_Click()
    'Me.Hide
    Load FRM_propos
    FRM_propos.Show
End Sub

Private Sub cmdParamTestClaquage_Click()
    '=======DEBUT MODIF BEEA 30/11/2007================
    Load FRM_Data_CtrlRefClaquage
    FRM_Data_CtrlRefClaquage.Show
    '=======FIN MODIF BEEA 30/11/2007================
End Sub

Private Sub cmdTestClaquage_Click()
    '=======DEBUT MODIF BEEA 30/11/2007================
    ControleStationClaquageEnCours = True
    CMD_ok_Click (TEST_STA_CLAQU)
    '=======FIN MODIF BEEA 30/11/2007==================
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Avec ou sans le deverminage ?
    If (KeyCode = KEY_F4) And (symbole_existe = True) And (Frame3D1.Visible = True) Then
        CMD_ok_Click (DEGRADE)
    End If
    
    If (KeyCode = KEY_F4) And (symbole_existe = True) And (cmdTestClaquage.Visible = True And cmdTestClaquage.Enabled = True) Then
        '=======DEBUT MODIF BEEA 30/11/2007================
        cmdTestClaquage_Click
        '=======FIN MODIF BEEA 30/11/2007==================
    End If
    
    If KeyCode = KEY_F2 And (symbole_existe = True) And (Frame3D1.Visible = True) Then
        CMD_ok_Click (TEMPERATURE)
    End If
    
    If KeyCode = KEY_F3 And (symbole_existe = True) And (Frame3D1.Visible = True) Then
        CMD_ok_Click (CONTROLE_FINAL)
    End If
    
    If KeyCode = KEY_F1 And (symbole_existe = True) And (Frame3D1.Visible = True) Then
        CMD_ok_Click (PREMIER_CONTROLE)
    End If
    
    If KeyCode = KEY_F5 And symbole_existe = True Then
        CMD_ok_Click (AUTO_TEST)
    End If
    
    If KeyCode = KEY_F6 And symbole_existe = True Then
        CMD_AFF_PARAM_Click
    End If
    
    If KeyCode = KEY_F7 Or KeyCode = vbKeyDivide Then
        'CMD_quitter_Click
        CMD_version_Click
    End If
    
    If KeyCode = KEY_F8 Or KeyCode = vbKeyNumpad7 Then
        CMD_sauve_log_Click
    End If

    If KeyCode = KEY_F9 Then
        Cmd_appel_maint_Click
    End If
    
    If KeyCode = KEY_F10 Then
        Cmd_etiquette_Click (0)
    End If
    
    If KeyCode = KEY_F11 Then
        Cmd_etiquette_Click (1)
    End If
    
    If KeyCode = KEY_F12 Or (KeyCode = KEY_TAB And Shift = KEY_CTRL) Then
        Cmd_etiquette_Click (2)
    End If
    
    If KeyCode = KEY_ESCAPE Or KeyCode = vbKeyNumpad0 Then
        CMD_quitter_Click
    End If

    '=======DEBUT MODIF BEEA 30/11/2007================
    If KeyCode = vbKeyC Then
        cmdParamTestClaquage_Click
    End If
    '=======FIN MODIF BEEA 30/11/2007==================
End Sub

Private Sub Form_Load()
Dim result As Single

    'RAZ variables dialogue PC-TSX
     Tab_Status_banc(0) = 0
     Tab_Status_banc(3) = 0
     Tab_Status_banc(4) = 0
     Tab_Status_banc(5) = 0
     Tab_Status_banc(6) = 0
     
    Lbl_code_lu.Caption = symb_com
    appel_maintenance = False
    
    '=======DEBUT MODIF BEEA 30/11/2007================
    'charge le mode de marche du test de la station de claquage
    Call LitModeTestStationClaquage
    Call AfficheModeDielec
    Call VerifSymboleTestClaquage
    'RAZ test de la station de claquage
    ControleStationClaquageEnCours = False
    Enr_claqu_en_cours = False
    Me.lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    '=======FIN MODIF BEEA 30/11/2007=================
    
    'CHARGE LE TIMER DE DIALOGUE AVEC LE TSX
    Load FRM_dial_tsx
    FRM_dial_tsx.Hide
    Contexte$ = "STARTUP"
    out_controle = 0
    lance_controle = 0
    out_autotest = 0
    lance_autotest = 0
    aff_param = 0
    forçage_etiquette = 0
    sans_etiquette = 0
    'Lbl_etiquette.Caption = "Impression etiquette sur défaut produit"
    Lbl_etiquette.Caption = "Print label for BAD Product"
    
    Me.WindowState = Mode_ouverture_fenetre
    
    file_l = get_log_path
    file_save = get_log_path_copy
    
    If file_l <> "" Then
        file_log$ = file_l & "LOG.TXT"
        file_hist_def$ = file_l & "HIST_DEF.TXT"
        path_log$ = file_l
    Else
        'MsgBox "Il faut définir un chemin aced aux fichiers journaux" & vbCrLf & " dans le fichier de configuration : " & File_para_ini, vbCritical
        MsgBox "Must define a path to log files" & vbCrLf & " in the configuration file : " & File_para_ini, vbCritical
        End
    
        'file_log$ = cst_file_log$
        'file_hist_def$ = cst_file_hist_def$
        'path_log$ = cst_path_log$
    End If
    
    If file_save = "" Then
        'MsgBox "Il faut définir un chemin aced aux copis des fichiers journaux" & vbCrLf & " dans le fichier de configuration : " & File_para_ini, vbCritical
        'End
        file_save = "c:\"
    End If
    
'Enregistrement du démarrage de l'application dans l'historique
    Open file_hist_def$ For Append As #3
    Print #3, "----------------------------------    "; Chr$(13); Chr$(10);
    Print #3, "LANCEMENT DU PROGRAMME DE CONTROLE le "; Str(Date); " à "; Str(Time); Chr$(13); Chr$(10);
    Print #3, "----------------------------------    "; Chr$(13); Chr$(10);
    Print #3, Chr$(13); Chr$(10);
    Print #3, Chr$(13); Chr$(10);
    Close #3
               
    Unload frmSplash
         
           
    Exit Sub
    
'100 MsgBox "Erreur " & Str$(err) & " module load de FRM_star. Signalez immediatement cette erreur. Merci."
100 MsgBox "Erreur " & Str$(err) & " module load de FRM_star. Signalez immediatement cette erreur. Merci."
     Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim result
Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
     'RAZ variables dialogue PC-TSX
        Copy_Tab(0) = 0
        Copy_Tab(3) = 0
        Copy_Tab(4) = 0
        Copy_Tab(5) = 0
        Copy_Tab(6) = 0
       
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
                For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                    Tab_Status_banc(i) = Copy_Tab(i)
                Next i
                previous_def_claquage = def_claquage
                Exit Sub
            End If
        Loop

    'tue le timer
    Unload FRM_dial_tsx
' soda debut test prog



' soda fin test prog
'base_data.Close
'table.Close

End Sub


Private Sub simu_badge_Click()
Tab_Status_banc(0) = Tab_Status_banc(0) + BIT2
End Sub

Private Sub Lbl_code_lu_Change()
'=======DEBUT MODIF BEEA 30/11/2007================
    cmdTestClaquage.Enabled = False
    Call VerifSymboleTestClaquage
'=======FIN MODIF BEEA 30/11/2007==================
End Sub

Private Sub timRefreshAffClaqu_Timer()
If Enr_claqu_en_cours Then
    Exit Sub
End If
'=======DEBUT MODIF BEEA 30/11/2007================
If Contexte$ = "STARTUP" Then
    Call AfficheModeDielec
    Call VerifSymboleTestClaquage
End If
'=======FIN MODIF BEEA 30/11/2007==================
End Sub

Private Sub timTestClaqu_Timer()
'=======DEBUT MODIF BEEA 30/11/2007================
If Enr_claqu_en_cours Then
    Exit Sub
End If
If Contexte$ = "STARTUP" Then
    Call LitModeTestStationClaquage
End If
'=======FIN MODIF BEEA 30/11/2007==================
End Sub
'=======DEBUT MODIF BEEA 30/11/2007==================
Public Sub AfficheModeDielec()
On Error GoTo erreur
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_LISTE_MODES.State <> adStateClosed Then
        DataEnvironment1.rsTable_LISTE_MODES.Close
    End If
    DataEnvironment1.rsTable_LISTE_MODES.Open
    
    DataEnvironment1.rsTable_LISTE_MODES.Filter = ("code_mode='" & ModeMarcheDielec & "'")
    If DataEnvironment1.rsTable_LISTE_MODES.RecordCount <> 0 Then
            lblModeDielec.Caption = DataEnvironment1.rsTable_LISTE_MODES.Fields("Mode")
            If ModeMarcheDielec = 0 Then
                lblModeClaquage.Visible = False
                lblModeDielec.Visible = False
            Else
                lblModeClaquage.Visible = True
                lblModeDielec.Visible = True
            End If
    End If
    
    'Rafraichissement état station claquage (important avant test pour changement de mode)
    If ControleStationClaquageEnCours = False Then
        Select Case ModeMarcheDielec
            Case 0, 1
                TestStationClaquage1OK = True
                TestStationClaquage2OK = True
            Case 2
                TestStationClaquage1OK = False
                TestStationClaquage2OK = True
            Case 3
                TestStationClaquage1OK = True
                TestStationClaquage2OK = False
            Case 4
                TestStationClaquage1OK = False
                TestStationClaquage2OK = False
            Case Else
                TestStationClaquage1OK = TestStationClaquage1OK
                TestStationClaquage2OK = TestStationClaquage2OK
            End Select
    End If

Exit Sub
erreur:
'reponse = MsgBox("erreur procédure affichage mode test station de claquage", vbCritical, "ERREUR")
reponse = MsgBox("ERROR procedure display mode test station HIPOT", vbCritical, "ERROR")

'=======FIN MODIF BEEA 30/11/2007==================
End Sub

'=======DEBUT MODIF BEEA 30/11/2007================
Public Sub VerifSymboleTestClaquage()
On Error GoTo erreur
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_SYMBOLES_REF.State <> adStateClosed Then
        DataEnvironment1.rsTable_SYMBOLES_REF.Close
    End If
    DataEnvironment1.rsTable_SYMBOLES_REF.Open
    
    DataEnvironment1.rsTable_SYMBOLES_REF.Filter = ("")
    DataEnvironment1.rsTable_SYMBOLES_REF.Filter = ("symbole='" & Lbl_code_lu.Caption & "'")
        If (DataEnvironment1.rsTable_SYMBOLES_REF.RecordCount >= 1) Then
            Frame3D1.Enabled = False
            Frame3D1.Visible = False
            fraTestStationClaquage.Enabled = True
            fraTestStationClaquage.Visible = True
            cmdTestClaquage.Visible = True
            cmdTestClaquage.Enabled = True
        ElseIf (ModeMarcheDielec = 1) Or (ModeMarcheDielec = 5) Then
            Frame3D1.Enabled = False
            Frame3D1.Visible = False
            fraTestStationClaquage.Enabled = True
            fraTestStationClaquage.Visible = True
            cmdTestClaquage.Visible = False
            cmdTestClaquage.Enabled = False
        Else
            Frame3D1.Enabled = True
            Frame3D1.Visible = True
            fraTestStationClaquage.Enabled = False
            fraTestStationClaquage.Visible = False
            cmdTestClaquage.Visible = False
            cmdTestClaquage.Enabled = False
        End If

Exit Sub
erreur:
'reponse = MsgBox("erreur procédure vérification symbole de test claquage", vbCritical, "ERREUR")
reponse = MsgBox("ERROR procedure Verifikasi symbol HIPOT Test", vbCritical, "ERROR")

End Sub
'=======FIN MODIF BEEA 30/11/2007==================

