VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FRM_AUTOTEST 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AUTOTEST"
   ClientHeight    =   6375
   ClientLeft      =   1845
   ClientTop       =   1665
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Clignoteur 
      Interval        =   750
      Left            =   2880
      Top             =   5760
   End
   Begin VB.CommandButton CMB_fin_autotest 
      Caption         =   "F1 - OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   2100
   End
   Begin VB.CommandButton CMB_Annulation_autotest 
      Caption         =   "F5 - Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   2100
   End
   Begin VB.CommandButton CMB_Validation_autotest 
      BackColor       =   &H00C0C0C0&
      Caption         =   "F1 - Valider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   15
      TabIndex        =   0
      Top             =   3105
      Width           =   2100
   End
   Begin Threed.SSFrame FRM_COM_TSX 
      Height          =   555
      Left            =   360
      TabIndex        =   6
      Top             =   5640
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   979
      _StockProps     =   14
      Caption         =   "Communication..."
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
      Begin VB.Label Lbl_pas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "001"
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
         Height          =   192
         Left            =   600
         TabIndex        =   7
         Top             =   312
         Width           =   648
      End
   End
   Begin Threed.SSPanel SSP_resultat 
      Height          =   2055
      Index           =   0
      Left            =   2280
      TabIndex        =   10
      Top             =   3120
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   3625
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox IMG_valeurs 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   120
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   219
         TabIndex        =   11
         Top             =   120
         Width           =   3285
      End
   End
   Begin Threed.SSPanel SSP_resultat 
      Height          =   2055
      Index           =   1
      Left            =   6000
      TabIndex        =   12
      Top             =   3120
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   3625
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox IMG_valeurs 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   120
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   219
         TabIndex        =   13
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Label Lbl_message_operateur 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   9612
   End
   Begin VB.Label LBL_F11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "F11 :  for return to NORMAL mode"
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
      Height          =   420
      Left            =   6930
      TabIndex        =   9
      Top             =   8115
      Width           =   1635
   End
   Begin VB.Label Label5 
      Caption         =   "F11 : for enter MAINTENANCE MODE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   9360
      TabIndex        =   8
      Top             =   8535
      Width           =   2640
   End
   Begin VB.Label LBL_Autotest_termine 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "AUTOTEST finish and OK !  Press OK to exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   7800
   End
   Begin VB.Label LBL_AUTOTEST_EN_COURS 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "             AUTOTEST IN PROGRESS        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   9630
   End
   Begin VB.Label LBL_debut_autotest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press OK to START AUTOTEST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1725
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   8895
   End
End
Attribute VB_Name = "FRM_AUTOTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Clignoteur_Timer()

    If LBL_debut_autotest.Visible = True Then
        If LBL_debut_autotest.ForeColor = 0 Then
            LBL_debut_autotest.ForeColor = &HFFFF&
        Else
            LBL_debut_autotest.ForeColor = 0
        End If
    End If
End Sub

Private Sub CMB_Annulation_autotest_Click()
Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i

    Clignoteur.Enabled = False
    
    resultat_autotest = 0

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
            If MsgBox(ERREUR_ECRITURE_AUTO_TEST, vbCritical + vbRetryCancel) = vbCancel Then
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
    Unload FRM_AUTOTEST
    FRM_startup.Show
End Sub

Private Sub CMB_fin_autotest_Click()
    Unload FRM_AUTOTEST
    out_autotest = 1
End Sub

Private Sub CMB_Validation_autotest_Click()
Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i


        Copy_Tab(0) = 0
        Copy_Tab(3) = 0
        Copy_Tab(4) = 0
        Copy_Tab(5) = 0
        Copy_Tab(6) = 0
        
    set_bit Tab_Status_banc(0), 8 'demande autotest

    Dim Com_erreure As Integer
    Com_erreure = -1
    Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
        Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 10)
        If Com_erreure <= 0 Then ' Communication en erreur
            If MsgBox(ERREUR_ECRITURE_AUTO_TEST, vbCritical + vbRetryCancel) = vbCancel Then
                If MsgBox(MSG_CONFIRMER_QUITE, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
                    End
                End If
            End If
        Else ' Communication sans erreur
            For i = LBound(Copy_Tab) To UBound(Copy_Tab)
                Tab_Status_banc(i) = Copy_Tab(i)
            Next i
            'memo defaut precedent
         '   previous_def_chargement = def_chargement
         '   Exit Sub
        End If
    Loop

    autoris = 1  'Pour ne pas sortir de l'autotest avant de l'avoir lancé
    LBL_debut_autotest.Visible = False
    CMB_Validation_autotest.Visible = False
    'CMB_Annulation_autotest.Visible = False
    LBL_AUTOTEST_EN_COURS.Visible = True
    LBL_Autotest_termine.Visible = False
    FRM_AUTOTEST!LBL_Autotest_termine.Visible = False
    CMB_fin_autotest.Visible = False
    
    Clignoteur.Enabled = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEY_F1 Then
        If CMB_fin_autotest.Visible = True Then
        CMB_fin_autotest_Click
        End If
        
        If CMB_Validation_autotest.Visible = True Then
        CMB_Validation_autotest_Click
        End If
    End If
    
    If KeyCode = KEY_F5 Then
        If CMB_Annulation_autotest.Visible = True Then
        CMB_Annulation_autotest_Click
        End If
    End If
End Sub

Private Sub Form_Load()
Dim init

    autoris = 0
    Clignoteur.Enabled = True
    LBL_debut_autotest.Visible = True
    CMB_Validation_autotest.Visible = True
    CMB_Annulation_autotest.Visible = True
    LBL_AUTOTEST_EN_COURS.Visible = False
    LBL_Autotest_termine.Visible = False
    FRM_AUTOTEST!LBL_Autotest_termine.Visible = False
    CMB_fin_autotest.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim result
    
    ' Tue le driver IEEE
    '  result = mesure(2)

End Sub



