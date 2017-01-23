VERSION 5.00
Begin VB.Form FRM_FIN_CONTROLE 
   Caption         =   "Verification qty of good products"
   ClientHeight    =   4035
   ClientLeft      =   870
   ClientTop       =   1575
   ClientWidth     =   8100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_valider 
      Caption         =   "  F10-&OK end test"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   3000
      Width           =   2205
   End
   Begin VB.CommandButton Cmd_annuler 
      BackColor       =   &H00808080&
      Caption         =   "   F7- &Continue Test"
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
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   2190
   End
   Begin VB.TextBox Txt_nbre_bon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Lbl_bons 
      Caption         =   "QTY good products :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Lbl_symbole 
      Caption         =   "Lbl_symbole"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Lbl_lot 
      Caption         =   "Lbl_lot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Lbl_titre 
      Alignment       =   2  'Center
      Caption         =   "TEST FINISHED"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "FRM_FIN_CONTROLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_annuler_Click()
    out_controle = 0
    Unload Me
End Sub


Private Sub Cmd_valider_Click()
Dim nbr As Integer

Dim i As Integer
Dim Copy_Tab(20) As Long

    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i

    'Interdiction de faire fin des série si les mauvais n'ont pas été enlevés
    If (Copy_Tab(4) And BIT10) <> 0 Then
        'MsgBox "Enlever les produits mauvais dans le bac"
        MsgBox "Remove bad product in the jig"
    End If
    
    'Interdiction de faire fin des série s'il reste des produits dans le poste
    ' soda If (Tab_Status_banc(4) And BIT11) = 0 Then
    ' soda test prog
    If (Copy_Tab(4) And BIT11) = 0 And prog_test = 0 Then
        'MsgBox "Impossible de valider la fin de série, il reste des produits dans le poste"
        MsgBox "Impossible to finish test, there are still product in the jig"
        Exit Sub
    End If
    
    'validation fin de série
    set_bit Copy_Tab(0), 6
    'Fin déroulement autorisé par le PC
    reset_bit Copy_Tab(0), 5
    
    '=========DEBUT MODIF BEEA 30/11/2007==========
    ControleStationClaquageEnCours = False
    '=========FIN MODIF BEEA 30/11/2007============
    
    nbr = Val(Txt_nbre_bon.Text)
    If nbr > 0 Then
        Copy_Tab(1) = nbr
    Else
        Copy_Tab(1) = 0
    End If
    
    ' finalise l'enregistrement des dernier parametre
    ' de la dernière entré dans la table "Log_banc" de la DB "Banc_XS_log"
    Fin_barquette_log Tab_Status_banc(1)
    
    out_controle = 1
    '=======DEBUT MODIF BEEA 30/11/2007================
    Call LitModeTestStationClaquage
    Call FRM_startup.AfficheModeDielec
    Call FRM_startup.VerifSymboleTestClaquage
    '=======FIN MODIF BEEA 30/11/2007==================
    Unload Me
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        Cmd_annuler_Click
    End If
    
   If KeyCode = vbKeyF10 Then
        Cmd_valider_Click
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Copy_Tab(20) As Long

    FinControleAffiche = True
    
    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
    
    If (Tab_Status_banc(1) + Tab_Status_banc(2)) < Nombre_produit_à_tester Then
        'MsgBox "Attention le nombre de produits testé est inférieur au nombre de produits bons déclarés à la phase précédente"
        MsgBox "WARNING qty product tested is less than the qty of good product in the previous phase"
    End If
    
    Lbl_symbole.Caption = symb_com
    Lbl_lot.Caption = "N° de lot" & Num_lot
    Txt_nbre_bon.Text = Tab_Status_banc(1)
    
    'Gestion des produits mauvais dans le bac
    If Tab_Status_banc(2) > 0 Then
        set_bit Copy_Tab(4), 10
        
        Dim Com_erreure As Integer
        Com_erreure = -1
        Do While Com_erreure <= 0 ' Boucle tanque la communication avec API n'a pas réusi
            Com_erreure = Ecrir_mots_etat_banc(Copy_Tab, 1, 4)
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
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
End Sub

Private Sub Form_Paint()
'met la fenetre au premier plan, et on ne peut plus l'effacer !
'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW

End Sub


Private Sub Form_Unload(Cancel As Integer)
    FinControleAffiche = False
    FRM_controle.SetFocus
End Sub
