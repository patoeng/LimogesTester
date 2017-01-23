VERSION 5.00
Begin VB.Form FRM_controle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banc XS contrôle final"
   ClientHeight    =   7728
   ClientLeft      =   3996
   ClientTop       =   1860
   ClientWidth     =   9660
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
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
   ScaleHeight     =   10409.25
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox FRM_tab_limites 
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
      Height          =   3870
      Left            =   2640
      ScaleHeight     =   3828
      ScaleWidth      =   4284
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
      Begin VB.TextBox TXT_limites 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3645
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   4080
      End
   End
   Begin VB.CommandButton simu_affich_resul 
      Caption         =   "simu affiche resultat"
      Height          =   615
      Left            =   6000
      TabIndex        =   33
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Cmd_prod2_mauvais 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "F11- Prod. 2 Mauvais"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7680
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1875
   End
   Begin VB.CommandButton Cmd_auto 
      Caption         =   "F11-AUTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   7050
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_prod1_mauvais 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "F1- Prod. 1 Mauvais"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6315
      Width           =   1875
   End
   Begin VB.CommandButton CMD_Next_Step 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "F12 - Etape Suivante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5820
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1785
   End
   Begin VB.PictureBox FRM_COM_TSX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3585
      ScaleHeight     =   612
      ScaleWidth      =   2124
      TabIndex        =   7
      Top             =   6285
      Width           =   2175
      Begin VB.Label LBL_etape_ctrl_final 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
         Height          =   195
         Left            =   1455
         TabIndex        =   28
         Top             =   420
         Width           =   645
      End
      Begin VB.Label LBL_etape_chargement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
         Height          =   195
         Left            =   765
         TabIndex        =   27
         Top             =   420
         Width           =   645
      End
      Begin VB.Label LBL_etape_claquage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
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
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   420
         Width           =   645
      End
      Begin VB.Shape SHA_rx 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   195
         Left            =   1740
         Shape           =   3  'Circle
         Top             =   195
         Width           =   195
      End
      Begin VB.Shape SHA_tx 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   195
         Left            =   480
         Shape           =   3  'Circle
         Top             =   210
         Width           =   195
      End
      Begin VB.Label LBL_RX 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "RX"
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
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   210
         Width           =   315
      End
      Begin VB.Label LBL_TX 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TX"
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
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CommandButton CMD_retour 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "F7 - &FIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2055
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6315
      Width           =   1485
   End
   Begin VB.PictureBox FRM_ref_produit 
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
      Height          =   660
      Left            =   105
      ScaleHeight     =   612
      ScaleWidth      =   9324
      TabIndex        =   1
      Top             =   -45
      Width           =   9375
      Begin VB.Label Lbl_mauvais 
         Caption         =   "Nombre de produits mauvais :"
         Height          =   255
         Left            =   5355
         TabIndex        =   30
         Top             =   375
         Width           =   3870
      End
      Begin VB.Label Lbl_bon 
         Caption         =   "Nombre de produits bons :"
         Height          =   255
         Left            =   5355
         TabIndex        =   29
         Top             =   120
         Width           =   3840
      End
      Begin VB.Label LBL_no_de_lot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lot n° "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3570
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LBL_symb_com 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox SSP_prod1 
      Height          =   5580
      Index           =   0
      Left            =   120
      ScaleHeight     =   5532
      ScaleWidth      =   2460
      TabIndex        =   13
      Top             =   705
      Width           =   2505
      Begin VB.PictureBox FRM_tab_caract 
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
         Height          =   4005
         Index           =   0
         Left            =   60
         ScaleHeight     =   3960
         ScaleWidth      =   2352
         TabIndex        =   16
         Top             =   1530
         Width           =   2400
         Begin VB.PictureBox IMG_valeurs 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3825
            Index           =   0
            Left            =   90
            ScaleHeight     =   319
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   184
            TabIndex        =   17
            Top             =   75
            Width           =   2205
         End
      End
      Begin VB.Label LBL_caract 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Produit Numéro 1"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1260
         Width           =   2295
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
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2190
      End
      Begin VB.Shape SHP_voyant_vert 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         BorderWidth     =   3
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   1110
         Index           =   0
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   2385
      End
   End
   Begin VB.PictureBox SSPanel1 
      Height          =   5565
      Left            =   7005
      ScaleHeight     =   5520
      ScaleWidth      =   2472
      TabIndex        =   18
      Top             =   720
      Width           =   2520
      Begin VB.PictureBox FRM_tab_caract 
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
         Height          =   3990
         Index           =   1
         Left            =   45
         ScaleHeight     =   3948
         ScaleWidth      =   2400
         TabIndex        =   21
         Top             =   1545
         Width           =   2445
         Begin VB.PictureBox IMG_valeurs 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            FillColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3825
            Index           =   1
            Left            =   90
            ScaleHeight     =   319
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   188
            TabIndex        =   22
            Top             =   75
            Width           =   2250
         End
      End
      Begin VB.Label LBL_caract 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Produit Numéro 2"
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
         Left            =   90
         TabIndex        =   20
         Top             =   1245
         Width           =   2355
      End
      Begin VB.Label LBL_pbon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   480
         Width           =   2355
      End
      Begin VB.Shape SHP_voyant_vert 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         BorderWidth     =   3
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   1110
         Index           =   1
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   45
         Width           =   2415
      End
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   384
      Left            =   285
      ScaleHeight     =   336
      ScaleWidth      =   912
      TabIndex        =   34
      Top             =   7080
      Width           =   960
   End
   Begin VB.Label LBL_etat_cycle_ctrl_final 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   5565
      TabIndex        =   26
      Top             =   1440
      Width           =   1410
   End
   Begin VB.Label LBL_etat_cycle_chargement 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   4080
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label LBL_etat_cycle_claquage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2655
      TabIndex        =   24
      Top             =   1425
      Width           =   1395
   End
   Begin VB.Label LBL_limites 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limites"
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
      Height          =   255
      Left            =   4305
      TabIndex        =   5
      Top             =   2115
      Width           =   975
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
End
Attribute VB_Name = "FRM_controle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
DefInt A-Z
















Private Sub Cmd_auto_Click()

If (MW350(0) And BIT0) <> 0 Then
    reset_bit MW350(0), 0
    Write_Buf_W350(0) = MW350(0)
    Call ecrit_tsx(350, Write_Buf_W350(), 1)
End If

End Sub


Private Sub CMD_Next_Step_Click()

    If CMD_Next_Step.Visible = True Then
        Write_Buf_W350(0) = MW350(0)
        set_bit MW350(0), 1
        set_bit Write_Buf_W350(0), 1
        Call ecrit_tsx(350, Write_Buf_W350(), 1)
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
    FRM_FIN_CONTROLE.Show
End Sub


Private Sub Command1_Click()

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


Private Sub Form_Load()
Dim i

    'soda debut prog test
    If prog_test Then
        simu_affich_resul.Visible = True
    Else
        simu_affich_resul.Visible = False
    End If
    'fin soda prog test
    
    LBL_symb_com.Caption = symb_com
    LBL_no_de_lot.Caption = "Lot n° " + Num_lot
    
    If Num_lot <> "" Then
        LBL_no_de_lot.Visible = True
    Else
        LBL_no_de_lot.Visible = False
    End If
    
    'Initialise l'ecran de contrôle
    SHP_voyant_vert(0).Visible = False
    LBL_pbon(0).Visible = False
    SHP_voyant_vert(1).Visible = False
    LBL_pbon(1).Visible = False
    FRM_tab_caract(0).Visible = True
    IMG_valeurs(0).Visible = True
    FRM_tab_caract(1).Visible = True
    IMG_valeurs(1).Visible = True
    CMD_Next_Step.Visible = False
    Cmd_prod1_mauvais.Visible = False
    Cmd_prod2_mauvais.Visible = False
    
    
    'init table echange PC<->TSX
    'Preparation mode dégradé
    If type_controle = 0 Then
                If pas_a_pas_chargement = True Then
                    FRM_controle.Caption = "Banc XS : MODE PAS A PAS CHARGEMENT"
                    pas_a_pas_chargement = False
                    CMD_Next_Step.Visible = True

                End If
            
                If pas_a_pas_claquage = True Then
                    FRM_controle.Caption = "Banc XS : MODE PAS A PAS CLAQUAGE"
                    pas_a_pas_claquage = False
                    CMD_Next_Step.Visible = True

                End If
                
                If pas_a_pas_final = True Then
                    FRM_controle.Caption = "Banc XS : MODE PAS A PAS CONTROLE FINAL"
                    pas_a_pas_final = False
                    CMD_Next_Step.Visible = True

                End If
                
                If auto_chargement = True Then
                    FRM_controle.Caption = "Banc XS : MODE AUTO CHARGEMENT"
                    auto_chargement = False
                End If
            
                If auto_claquage = True Then
                    FRM_controle.Caption = "Banc XS : MODE AUTO CHARGEMENT ET CLAQUAGE"
                    auto_claquage = False
                End If
                
                If auto_final = True Then
                    FRM_controle.Caption = "Banc XS : MODE AUTO CHARGEMENT ET FINAL"
                    auto_final = False
                End If
    End If
    
    
        ' Préparation controle température
    If type_controle = 1 Then
        FRM_controle.Caption = "Banc XS : CONTROLE TEMPERATURE"
    End If
    
        ' Préparation controle complet
    If type_controle = 2 Then
        FRM_controle.Caption = "Banc XS : CONTROLE FINAL"
    End If
    
    
    'affiche les valeurs limites
    Call prepa_params
    Call maj_text_limites
    

    'Soyons sûrs que l'ecran a ete mis à jour
    'sauv_com_cycle = 999

    'setcursorpos 640, 480

    Me.Show
    
    DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim result
 
    'Ferme les fichiers LOG
    Close
    FRM_startup!Lbl_code_lu.Caption = ""
    FRM_startup.Show

End Sub


    
 







Private Sub simu_affich_resul_Click()
test_prog.Show
set_bit MW350(66), 3
Call RAZ_ecran
prepa_resultats
MW350(0) = MW350(0) + BIT3
End Sub


