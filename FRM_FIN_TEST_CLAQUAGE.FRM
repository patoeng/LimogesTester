VERSION 5.00
Begin VB.Form FRM_FIN_TEST_CLAQUAGE 
   Caption         =   "Result of DUMMY TEST Hipot"
   ClientHeight    =   5310
   ClientLeft      =   3615
   ClientTop       =   5415
   ClientWidth     =   8100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_valider 
      Caption         =   "  F10-&OK to finish"
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
      TabIndex        =   2
      Top             =   4440
      Width           =   2205
   End
   Begin VB.CommandButton Cmd_annuler 
      BackColor       =   &H00808080&
      Caption         =   "   F7- &Continue test"
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
      TabIndex        =   1
      Top             =   4440
      Width           =   2190
   End
   Begin VB.Label lblTexteClaquage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "lblTexteClaquage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   840
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Width           =   6765
   End
   Begin VB.Label lblEtatClaquage2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "lblEtatClaquage2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblEtatClaquage1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "lblEtatClaquage1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblEtatStationClaquage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "lblEtatStationClaquage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2505
      TabIndex        =   3
      Top             =   1560
      Width           =   3165
   End
   Begin VB.Label Lbl_titre 
      Alignment       =   2  'Center
      Caption         =   "FINISH DUMMY TEST HIPOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "FRM_FIN_TEST_CLAQUAGE"
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

    'Interdiction de faire fin des série s'il reste des produits dans le poste
    ' soda If (Tab_Status_banc(4) And BIT11) = 0 Then
    ' soda test prog
    If (Copy_Tab(4) And BIT11) = 0 And prog_test = 0 Then
        'MsgBox "Impossible de valider la fin de série, il reste des produits dans le poste"
        MsgBox "Impossible to finish test, there are still product in the jig"
        Exit Sub
    End If
    
    'pour éviter les rebonds de contact ...?...
    If Enr_claqu_en_cours Then
        Exit Sub
    End If
    Enr_claqu_en_cours = True
    
    'validation fin de série
    set_bit Copy_Tab(0), 6
    'Fin déroulement autorisé par le PC
    reset_bit Copy_Tab(0), 5
    
    nbr = Tab_Status_banc(1)
    If nbr > 0 Then
        Copy_Tab(1) = nbr
    Else
        Copy_Tab(1) = 0
    End If
    
    ' finalise l'enregistrement des dernier parametre
    ' de la dernière entré dans la table "Log_banc" de la DB "Banc_XS_log"
    Fin_barquette_log Tab_Status_banc(1)
    
    '=========DEBUT MODIF BEEA 30/11/2007==========
    ControleStationClaquageEnCours = False
    Call EcritModeTestStationClaquage(ModeMarcheDielec)
    '=========FIN MODIF BEEA 30/11/2007============
    
    out_controle = 1
    Enr_claqu_en_cours = False
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
    FinTestClaquageAffiche = True
    
    Dim i As Integer
    Dim Copy_Tab(20) As Long
    
    For i = LBound(Copy_Tab) To UBound(Copy_Tab)
        Copy_Tab(i) = Tab_Status_banc(i)
    Next i
    
    'Résultats du test
    'Les 2 stations sont OK = le poste est bon, retour à la normale
    If TestStationClaquage1OK And TestStationClaquage2OK Then
           ModeMarcheDielec = 0
    End If
    '1 des station est HS = le poste est mauvais, retour à l'état de demande
    If (Not TestStationClaquage1OK) Or (Not TestStationClaquage2OK) Then
           ModeMarcheDielec = 5
    End If
    
    'Affichage résultat du test
    Select Case ModeMarcheDielec
        Case 0
            lblEtatStationClaquage.ForeColor = &HFF00&
            'lblEtatStationClaquage.Caption = "STATION DE CLAQUAGE OK"
            'lblTexteClaquage.Caption = "Vous pouvez utiliser ce banc. Le passage des produits de référence 'claquage' a été concluant."
            lblEtatStationClaquage.Caption = "HIPOT TEST OK"
            lblTexteClaquage.Caption = "You can use the bench. Dummy product 'claquage' test is OK."
            lblEtatStationClaquage.ForeColor = &H8000&
        Case Else
            'lblEtatStationClaquage.Caption = "STATION DE CLAQUAGE HS"
            'lblTexteClaquage.Caption = "Vous ne pouvez plus utiliser ce banc jusqu'à un nouveau passage concluant des produits de référence 'claquage', ou une mise en mode dégradé par une personnne autorisée."
            lblEtatStationClaquage.Caption = "HIPT TEST FAIL"
            lblTexteClaquage.Caption = "You can NOT use the bench until Dummy product 'claquage' test is OK, or degrade mode is set by authorized person."
            lblEtatStationClaquage.ForeColor = &HFF&
            lblEtatStationClaquage.ForeColor = &HC0&
    End Select
    lblEtatClaquage1.Caption = IIf(TestStationClaquage1OK = True, "Station 1 OK", "Station 1 FAIL")
    lblEtatClaquage1.BackColor = IIf(TestStationClaquage1OK = True, &HFF00&, &HFF&)
    lblEtatClaquage2.Caption = IIf(TestStationClaquage2OK = True, "Station 2 OK", "Station 2 FAIL")
    lblEtatClaquage2.BackColor = IIf(TestStationClaquage2OK = True, &HFF00&, &HFF&)
End Sub

Private Sub Form_Paint()
'met la fenetre au premier plan, et on ne peut plus l'effacer !
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FinTestClaquageAffiche = False
    FRM_controle.SetFocus
End Sub
