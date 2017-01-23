VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Data_CtrlRefClaquage 
   Caption         =   "Parameter for DUMMY TEST HIPOT"
   ClientHeight    =   7440
   ClientLeft      =   855
   ClientTop       =   900
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   0  'User
   ScaleWidth      =   592
   Begin MSComDlg.CommonDialog ComDlgOpenBase 
      Left            =   8280
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Base de données à ouvrir"
      FileName        =   "ctrl_ref_claquage.mdb"
      Filter          =   "Fichiers access (*.mdb)|*.mdb|Tous (*.*)|*.*"
      Flags           =   4100
   End
   Begin VB.Frame fraModeClaqu 
      Caption         =   "Current operating mode for Dummy Test Hipot"
      Height          =   975
      Left            =   120
      TabIndex        =   36
      Top             =   3480
      Width           =   8655
      Begin VB.CommandButton cmdEnrModeClaqu 
         Caption         =   "Register"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   39
         Top             =   315
         Width           =   1125
      End
      Begin VB.ComboBox cmbModeClaqu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   38
         ToolTipText     =   "Running mode of bench during Hipot Mode."
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "HIPOT current operating mode :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   405
         Width           =   2895
      End
   End
   Begin VB.Frame fraSymboleTestClaquage 
      Caption         =   "Product reference symbole for Dummy Test Hipot"
      Height          =   855
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   8655
      Begin VB.CommandButton cmdEnrSymbRef 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   42
         Top             =   195
         Width           =   1125
      End
      Begin VB.ComboBox cmbSymboleTestClaquage 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "Selection to change, type new to create."
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Product reference symbole :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   285
         Width           =   2775
      End
   End
   Begin VB.Frame FraNonConnecte 
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8655
      Begin VB.Label LblNonConnecte 
         Alignment       =   2  'Center
         Caption         =   "No Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame FraDataClaqu 
      Caption         =   "Frequency data for Dummy Test Hipot"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   8655
      Begin VB.TextBox txtHeureDemande 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   4
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Request test time"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtProchaineExecution 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "date and time of the next request test. (format  dd/mm/yyyy hh:mm:ss)"
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton btnEnrData 
         Caption         =   "Register"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   19
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtQuiDernierModif 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "the last modification of database is set by :"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtDerniereModif 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "the last modification of database took place on :"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtIntervalle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "only enter interval in days"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox CmbDernierTestOK 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Dernier Test OK"
         ToolTipText     =   "the last successful test was on :"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label12 
         Caption         =   "hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   40
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "next request for the execution :"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   885
         Width           =   2550
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Carried out by :"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   2325
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Last updated :"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1845
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "days,          Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   405
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Test Interval :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   405
         Width           =   2055
      End
      Begin VB.Label LblDernierTestOK 
         Alignment       =   1  'Right Justify
         Caption         =   "Last OK execution :"
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   1365
         Width           =   1935
      End
   End
   Begin VB.Frame FraCreatLog 
      Caption         =   "User Info"
      Height          =   1215
      Left            =   240
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   8415
      Begin VB.ComboBox cmbProfils 
         Height          =   315
         ItemData        =   "FRM_Data_CtrlRefClaquage.frx":0000
         Left            =   840
         List            =   "FRM_Data_CtrlRefClaquage.frx":000D
         TabIndex        =   12
         ToolTipText     =   "choisir le type de profil associé au login"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtAffPassword2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtLogin 
         Height          =   315
         Left            =   840
         TabIndex        =   11
         ToolTipText     =   "saisir un login pour modifier ou créer."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnModifUtil 
         Caption         =   "Modification"
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtAffPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCreatUtil 
         Caption         =   "Create"
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblProfil 
         Caption         =   "Profil :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Password :  "
         Height          =   255
         Left            =   2400
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "New Password :  "
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Login :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame FraConnexion 
      Caption         =   "Connection parameter for Dummy Test Hipot"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton CmdQuitter 
         Caption         =   "QUIT"
         Height          =   375
         Left            =   6960
         TabIndex        =   10
         ToolTipText     =   "Goodbye !!"
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton BtnConnexionDataCtrlRefClaqu 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         ToolTipText     =   "Connect to current database"
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox TxtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "your password"
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox CmbLogin 
         Height          =   315
         ItemData        =   "FRM_Data_CtrlRefClaquage.frx":0038
         Left            =   840
         List            =   "FRM_Data_CtrlRefClaquage.frx":003A
         TabIndex        =   7
         ToolTipText     =   "your login"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Password :"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label LblLogin 
         Caption         =   "Login :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   410
         Width           =   495
      End
   End
End
Attribute VB_Name = "FRM_Data_CtrlRefClaquage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======DEBUT MODIF BEEA 30/11/2007================
Private UtilisateurConnecte As Boolean
Private requ As String
Private PASSWORD As String, LOGIN As String, reponse As String, DernierTestOK As String
Private ProfilUtil As Integer
Private ProchaineDateDemandeChangee As Boolean

Private Function EqualNewPassword() As Boolean
        If TxtAffPassword.Text = TxtAffPassword2.Text Then
            EqualNewPassword = True
        Else
            'reponse = MsgBox("les mots de passe sont différents" & Chr(&HD) & Chr(&HA) & "Corrigez et recommencez!", vbExclamation, "Erreur de saisie")
            reponse = MsgBox("Password are different" & Chr(&HD) & Chr(&HA) & "Correct it dan try again!", vbExclamation, "Error Input")
            EqualNewPassword = False
        End If
End Function

'' To applied profil setting for user
Private Sub AppliqueProfilDataUtil(profil As Integer)
        Select Case profil
            Case 0  'operateur -> Operator
                FraCreatLog.Visible = True
                FraCreatLog.Enabled = False
                TxtLogin.Enabled = False
                cmbProfils.Enabled = False
                'FraDataClaqu.Caption = "Données Fréquentielles pour le Test de la Station de claquage : mode lecture seule"
                FraDataClaqu.Caption = "Frequency data for Hipot Testing : read-only mode"
                btnEnrData.Enabled = False
                'FraCreatLog.Caption = "Données utilisateur : mode lecture seule"
                FraCreatLog.Caption = "User Data : read-only mode"
                btnModifUtil.Enabled = False
                btnCreatUtil.Enabled = False
                Label7.Visible = False
                Label8.Visible = False
                TxtAffPassword.Visible = False
                TxtAffPassword2.Visible = False
                txtHeureDemande.Locked = True
                btnEnrData.Enabled = False
                txtIntervalle.Locked = True
                txtProchaineExecution.Locked = True
                CmbDernierTestOK.Locked = True
                txtDerniereModif.Locked = True
                txtQuiDernierModif.Locked = True
                fraSymboleTestClaquage.Visible = True
                FraDataClaqu.Visible = True
                fraModeClaqu.Visible = True
                cmbModeClaqu.Locked = True
                cmdEnrModeClaqu.Enabled = False
                cmbSymboleTestClaquage.Locked = True
                cmdEnrSymbRef.Enabled = False
            Case 1  'technicien -> Teknisi
                FraCreatLog.Visible = True
                FraCreatLog.Enabled = True
                TxtLogin.Enabled = False
                cmbProfils.Enabled = False
                'FraDataClaqu.Caption = "Données Fréquentielles pour le Test de la Station de claquage : mode modification"
                FraDataClaqu.Caption = "Frequency data for Hipot Testing : edit mode"
                btnEnrData.Enabled = False
                'FraCreatLog.Caption = "Données utilisateur : mode modification"
                FraCreatLog.Caption = "User data : edit mode"
                btnModifUtil.Enabled = False
                btnCreatUtil.Enabled = False
                Label7.Visible = True
                Label8.Visible = True
                TxtAffPassword.Visible = True
                TxtAffPassword2.Visible = True
                btnEnrData.Enabled = False
                txtIntervalle.Locked = False
                txtProchaineExecution.Locked = False
                txtHeureDemande.Locked = False
                CmbDernierTestOK.Locked = True
                txtDerniereModif.Locked = True
                txtQuiDernierModif.Locked = True
                fraSymboleTestClaquage.Visible = True
                FraDataClaqu.Visible = True
                fraModeClaqu.Visible = True
                cmbModeClaqu.Locked = False
                cmdEnrModeClaqu.Enabled = False
                cmbSymboleTestClaquage.Locked = False
                cmdEnrSymbRef.Enabled = False
            Case 2  'administrateur -> Administrator
                FraCreatLog.Visible = True
                FraCreatLog.Enabled = True
                TxtLogin.Enabled = True
                cmbProfils.Enabled = True
                FraDataClaqu.Caption = "Données Fréquentielles pour le Test de la Station de claquage : mode modification"
                FraDataClaqu.Caption = "Frequency data for Hipot Testing  : edit mode"
                btnEnrData.Enabled = False
                FraCreatLog.Caption = "Données utilisateur : mode création/modification"
                FraCreatLog.Caption = "User data : create/edit mode"
                btnModifUtil.Enabled = False
                btnCreatUtil.Enabled = False
                Label7.Visible = True
                Label8.Visible = True
                TxtAffPassword.Visible = True
                TxtAffPassword2.Visible = True
                btnEnrData.Enabled = False
                txtIntervalle.Locked = True
                txtProchaineExecution.Locked = True
                txtHeureDemande.Locked = True
                CmbDernierTestOK.Locked = True
                txtDerniereModif.Locked = True
                txtQuiDernierModif.Locked = True
                fraSymboleTestClaquage.Visible = False
                FraDataClaqu.Visible = False
                fraModeClaqu.Visible = False
                cmbModeClaqu.Locked = True
                cmdEnrModeClaqu.Enabled = False
                cmbSymboleTestClaquage.Locked = True
                cmdEnrSymbRef.Enabled = False
        End Select

End Sub

Private Function creer_ou_modifier() As Integer 'O = rien, 1= modifier, 2= créer
    requ = "SELECT * FROM ID_QUALITE WHERE login='" & TxtLogin.Text & "'"
    Set ligne_sel = DataEnvironment1.DataCtrlRefClaquage.Execute(requ)
    creer_ou_modifier = IIf(ligne_sel.RecordCount = 0, 2, 1)
End Function

Private Sub remplit_champs_frequentiels()
    CmbDernierTestOK.Clear
    Dim AvantDernierTestOK As String
    'Rafraichit les données en mémoire
    '' Referesh data in memory
    If DataEnvironment1.rsTable_Frequentiel.State <> adStateClosed Then
        DataEnvironment1.rsTable_Frequentiel.Close
    End If
    DataEnvironment1.rsTable_Frequentiel.Open
    DataEnvironment1.rsTable_Frequentiel.Filter = ("")
    If DataEnvironment1.rsTable_Frequentiel.EOF Or DataEnvironment1.rsTable_Frequentiel.BOF Then
        Exit Sub
    End If
    DataEnvironment1.rsTable_Frequentiel.MoveLast
    txtIntervalle.Text = CStr(DataEnvironment1.rsTable_Frequentiel.Fields("frequence").Value)
    txtHeureDemande.Text = CStr(DataEnvironment1.rsTable_Frequentiel.Fields("heure_demande").Value)
    txtDerniereModif.Text = CStr(DataEnvironment1.rsTable_Frequentiel.Fields("derniere_modif").Value)
    txtQuiDernierModif = CStr(DataEnvironment1.rsTable_Frequentiel.Fields("modifie_par").Value)
    AvantDernierTestOK = ""
    DernierTestOK = ""
    While DataEnvironment1.rsTable_Frequentiel.BOF <> True
        AvantDernierTestOK = DernierTestOK
        DernierTestOK = CStr(DataEnvironment1.rsTable_Frequentiel.Fields("derniere_execution").Value)
        If AvantDernierTestOK <> DernierTestOK Then
            CmbDernierTestOK.AddItem (DernierTestOK)
        End If
        DataEnvironment1.rsTable_Frequentiel.MovePrevious
    Wend
    If Me.CmbDernierTestOK.ListCount > 0 Then
        Me.CmbDernierTestOK.ListIndex = 0
        DernierTestOK = CmbDernierTestOK.Text
    End If
    DataEnvironment1.rsTable_Frequentiel.MoveLast
    txtProchaineExecution.Text = CStr(DataEnvironment1.rsTable_Frequentiel.Fields("prochaine_execution").Value)
    
End Sub

Private Sub remplit_champs_symboles()
    cmbSymboleTestClaquage.Clear
    'Rafraichit les données en mémoire
    '' Refersh data in memory
    If DataEnvironment1.rsTable_SYMBOLES_REF.State <> adStateClosed Then
        DataEnvironment1.rsTable_SYMBOLES_REF.Close
    End If
    DataEnvironment1.rsTable_SYMBOLES_REF.Open
    DataEnvironment1.rsTable_SYMBOLES_REF.Filter = ("")
    If DataEnvironment1.rsTable_SYMBOLES_REF.EOF Or DataEnvironment1.rsTable_SYMBOLES_REF.BOF Then
        Exit Sub
    End If
    DataEnvironment1.rsTable_SYMBOLES_REF.MoveFirst
    While DataEnvironment1.rsTable_SYMBOLES_REF.EOF <> True
        cmbSymboleTestClaquage.AddItem (CStr(DataEnvironment1.rsTable_SYMBOLES_REF.Fields("symbole").Value))
        DataEnvironment1.rsTable_SYMBOLES_REF.MoveNext
    Wend
    cmbSymboleTestClaquage.ListIndex = 0
End Sub

Private Sub BtnConnexionDataCtrlRefClaqu_Click()
    If TxtPassword.Text = PASSWORD And CmbLogin.Text = LOGIN Then
        UtilisateurConnecte = True
        FraNonConnecte.Visible = False
        FraCreatLog.Visible = True
        TxtLogin.Text = LOGIN
        cmbProfils.ListIndex = ProfilUtil
        Call AppliqueProfilDataUtil(ProfilUtil)
    End If
End Sub

Private Sub btnCreatUtil_Click()
If EqualNewPassword Then
    DataEnvironment1.rsTable_ID_QUALITE.AddNew
    DataEnvironment1.rsTable_ID_QUALITE.Fields("login").Value = TxtLogin.Text
    DataEnvironment1.rsTable_ID_QUALITE.Fields("password").Value = TxtAffPassword.Text
    DataEnvironment1.rsTable_ID_QUALITE.Fields("profil").Value = cmbProfils.ListIndex
    DataEnvironment1.rsTable_ID_QUALITE.Fields("password").Value = TxtAffPassword.Text
    DataEnvironment1.rsTable_ID_QUALITE.Update
End If
End Sub

Private Sub btnEnrData_Click()
    Dim ProchaineDemande As Date
    If ProchaineDateDemandeChangee Then
        ProchaineDateDemandeChangee = False
        If IsDate(Me.txtProchaineExecution.Text) Then
            ProchaineDemande = CDate(Me.txtProchaineExecution.Text)
        Else
            'reponse = MsgBox("La prochaine demande n'est pas au format jj/mm/aaaa hh:mm. Saisie non enregistrée", vbCritical, "EREUR DE SAISIE")
            reponse = MsgBox("The next request field is not follow format dd/mm/yyyy hh:mm. Entry not saved", vbCritical, "ERROR Input")
            Call remplit_champs_frequentiels
            Me.btnEnrData.Enabled = False
            Me.txtProchaineExecution.SetFocus
            Exit Sub
        End If
    Else
        ProchaineDemande = Int(CDate(DernierTestOK) + CDate(txtIntervalle.Text)) + CDate(txtHeureDemande.Text)
    End If
    DataEnvironment1.rsTable_Frequentiel.AddNew
    DataEnvironment1.rsTable_Frequentiel.Fields("frequence").Value = CInt(txtIntervalle.Text)
    DataEnvironment1.rsTable_Frequentiel.Fields("heure_demande").Value = CDate(txtHeureDemande.Text)
    DataEnvironment1.rsTable_Frequentiel.Fields("derniere_modif").Value = Now 'CDate(txtDerniereModif.Text)
    DataEnvironment1.rsTable_Frequentiel.Fields("modifie_par").Value = LOGIN
    DataEnvironment1.rsTable_Frequentiel.Fields("derniere_execution").Value = CDate(DernierTestOK)
    DataEnvironment1.rsTable_Frequentiel.Fields("prochaine_execution").Value = ProchaineDemande
    DataEnvironment1.rsTable_Frequentiel.Update
    Call remplit_champs_frequentiels
    btnEnrData.Enabled = False
End Sub

Private Sub btnModifUtil_Click()
If EqualNewPassword Then
    DataEnvironment1.rsTable_ID_QUALITE.Filter = "login='" & TxtLogin.Text & "'"
    DataEnvironment1.rsTable_ID_QUALITE.Fields("password").Value = TxtAffPassword.Text
    DataEnvironment1.rsTable_ID_QUALITE.Fields("profil").Value = cmbProfils.ListIndex
    DataEnvironment1.rsTable_ID_QUALITE.Update
    If TxtLogin.Text = LOGIN Then
        PASSWORD = TxtAffPassword.Text
    End If
End If
End Sub

Private Sub CmbLogin_Click()
    AppliqueProfilDataUtil (0)
    UtilisateurConnecte = False
    FraNonConnecte.Visible = True
    FraCreatLog.Visible = False
    FraCreatLog.Enabled = False
    FraDataClaqu.Visible = False
    DataEnvironment1.rsTable_ID_QUALITE.MoveFirst
    requ = "SELECT * FROM ID_QUALITE WHERE login='" & CmbLogin.Text & "'"
    Set ligne_sel = DataEnvironment1.DataCtrlRefClaquage.Execute(requ)
    If Not IsNull(ligne_sel.Fields("login").Value) Then
        LOGIN = CStr(ligne_sel.Fields("login").Value)
    Else
        LOGIN = ""
    End If
    If Not IsNull(ligne_sel.Fields("password").Value) Then
        PASSWORD = CStr(ligne_sel.Fields("password").Value)
    Else
        PASSWORD = ""
    End If
    If Not IsNull(ligne_sel.Fields("profil").Value) Then
        ProfilUtil = CStr(ligne_sel.Fields("profil").Value)
    Else
        ProfilUtil = ""
    End If
    TxtPassword = ""
    TxtAffPassword.Text = ""
    TxtAffPassword2.Text = ""
End Sub

Private Sub cmbModeClaqu_Click()
    If ProfilUtil = 1 Then
        cmdEnrModeClaqu.Enabled = True
    End If
End Sub

Private Sub cmbSymboleTestClaquage_Change()
    If ProfilUtil = 1 Then
        cmdEnrSymbRef.Enabled = True
        cmdEnrSymbRef.Caption = "Créer"
    End If
End Sub

Private Sub cmbSymboleTestClaquage_Click()
    If ProfilUtil = 1 Then
        cmdEnrSymbRef.Enabled = True
        cmdEnrSymbRef.Caption = "Supprimer"
    End If
End Sub

Private Sub cmdEnrModeClaqu_Click()
    If cmbModeClaqu.ListIndex = 0 Or (cmbModeClaqu.ListIndex = 2 And Not TestStationClaquage2OK) Or (cmbModeClaqu.ListIndex = 3 And Not TestStationClaquage1OK) Then
        'reponse = MsgBox("CE MODE N'EST PAS SELECTIONNABLE!" & Chr(&HD) & Chr(&HA) & "RECOMMENCEZ", vbExclamation, "ERREUR DECHOIX")
        reponse = MsgBox("THIS MODE IS CAN'T BE SELECTED!" & Chr(&HD) & Chr(&HA) & "TRY AGAIN", vbExclamation, "ERROR CHOICE")
        cmbModeClaqu.Text = CStr(DataEnvironment1.rsTable_LISTE_MODES.Fields("Mode").Value)
        GoTo sortie
    End If
On Error GoTo err
    DataEnvironment1.rsTable_LISTE_MODES.Filter = ("Mode='" & cmbModeClaqu.Text & "'")
    ModeMarcheDielec = DataEnvironment1.rsTable_LISTE_MODES.Fields("code_mode").Value
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Fields("MODE_ACTUEL_CLAQUAGE").Value = ModeMarcheDielec
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Update
    GoTo sortie
err:
        'reponse = MsgBox("CE MODE NE FAIT PAS PARTIE DE LA LISTE" & Chr(&HD) & Chr(&HA) & "RECOMMENCEZ", vbExclamation, "ERREUR DECHOIX")
        reponse = MsgBox("This MODE IS NOT PART OF THE LIST" & Chr(&HD) & Chr(&HA) & "TRY AGAIN", vbExclamation, "ERROR CHOICE")
sortie:
    cmdEnrModeClaqu.Enabled = False
End Sub

Private Sub cmdEnrSymbRef_Click()
    If cmdEnrSymbRef.Enabled = True Then
        'Rafraichit les données en mémoire
        If DataEnvironment1.rsTable_SYMBOLES_REF.State <> adStateClosed Then
            DataEnvironment1.rsTable_SYMBOLES_REF.Close
        End If
        DataEnvironment1.rsTable_SYMBOLES_REF.Open
        DataEnvironment1.rsTable_SYMBOLES_REF.Filter = ("symbole='" & cmbSymboleTestClaquage.Text & "'")
        Select Case cmdEnrSymbRef.Caption
            Case "Supprimer"
                If (DataEnvironment1.rsTable_SYMBOLES_REF.RecordCount > 0) Then
                    DataEnvironment1.rsTable_SYMBOLES_REF.Delete
                End If
            
            Case "Créer"
                If (DataEnvironment1.rsTable_SYMBOLES_REF.RecordCount <= 0) Then
                    DataEnvironment1.rsTable_SYMBOLES_REF.AddNew
                    DataEnvironment1.rsTable_SYMBOLES_REF.Fields("symbole").Value = cmbSymboleTestClaquage.Text
                    DataEnvironment1.rsTable_SYMBOLES_REF.Update
                End If
        End Select
    End If
    cmdEnrSymbRef.Enabled = False
    Call remplit_champs_symboles
End Sub

Private Sub Form_Paint()
'met la fenetre au premier plan, et on ne peut plus l'effacer !
'' Put the window in background and you can't erase
    SetWindowPos hwnd, HWND_TOPMOST, Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
End Sub

Private Sub TxtAffPassword_change()
    btnCreatUtil.Enabled = (creer_ou_modifier = 2) And (ProfilUtil > 1)
    btnModifUtil.Enabled = (creer_ou_modifier = 1) And (ProfilUtil > 0)
End Sub

Private Sub TxtAffPassword2_change()
    btnCreatUtil.Enabled = (creer_ou_modifier = 2) And (ProfilUtil > 1)
    btnModifUtil.Enabled = (creer_ou_modifier = 1) And (ProfilUtil > 0)
End Sub

Private Sub txtHeureDemande_Change()
If UtilisateurConnecte And (ProfilUtil > 0) Then
    btnEnrData.Enabled = True
End If
End Sub

Private Sub TxtLogin_change()
    btnCreatUtil.Enabled = (creer_ou_modifier = 2) And (ProfilUtil > 1)
    btnModifUtil.Enabled = (creer_ou_modifier = 1) And (ProfilUtil > 0)
End Sub

Private Sub CmdQuitter_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    UtilisateurConnecte = False
    FraNonConnecte.Visible = True
    Call AppliqueProfilDataUtil(0)
    'remplissage combo des identifiants
    'Rafraichit les données en mémoire
    '' Fill combo Login name
    '' Refresh data in memory
    DataEnvironment1.DataCtrlRefClaquage.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" _
        & "Persist Security Info=False;" _
        & "Data Source=" & App.Path & "\Database\ctrl_ref_claquage.mdb"
OuvrirDataBase:
    On Error GoTo ErreurConnBase
    If DataEnvironment1.rsTable_ID_QUALITE.State <> adStateClosed Then
        DataEnvironment1.rsTable_ID_QUALITE.Close
    End If
    GoTo Suite
ErreurConnBase:
        If err = &H80040200 Then
            Resume OpenFichier
        Else
            'MsgBox ("Erreur 0x" & Hex(err.Number) & " : " & err.Description)
            MsgBox ("Error 0x" & Hex(err.Number) & " : " & err.Description)
            Resume Annul
        End If
    Exit Sub
OpenFichier:
        On Error GoTo Annul
        Me.ComDlgOpenBase.ShowOpen
        On Error Resume Next
        'connexion base de données
        '' Connection to database
        DataEnvironment1.DataCtrlRefClaquage.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" _
                & Me.ComDlgOpenBase.FileName '_
        GoTo OuvrirDataBase
Annul:
        End
Suite:
    DataEnvironment1.rsTable_ID_QUALITE.Open
    DataEnvironment1.rsTable_ID_QUALITE.Filter = ("")
    If DataEnvironment1.rsTable_ID_QUALITE.RecordCount < 1 Then
        GoTo suite1
    End If
    DataEnvironment1.rsTable_ID_QUALITE.MoveFirst
        While DataEnvironment1.rsTable_ID_QUALITE.EOF <> True
            CmbLogin.AddItem (CStr(DataEnvironment1.rsTable_ID_QUALITE.Fields("login").Value))
            DataEnvironment1.rsTable_ID_QUALITE.MoveNext
        Wend
    CmbLogin.ListIndex = 0
suite1:
    'remplissage combo du mode de marche du claquage
    'Rafraichit les données en mémoire
    '' Fill combo for Run mode Hipot
    '' Refresh data in memory
    If DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.State <> adStateClosed Then
        DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Close
    End If
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Open
    DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Filter = ("")
    'Rafraichit les données en mémoire
    If DataEnvironment1.rsTable_LISTE_MODES.State <> adStateClosed Then
        DataEnvironment1.rsTable_LISTE_MODES.Close
    End If
    DataEnvironment1.rsTable_LISTE_MODES.Open
    DataEnvironment1.rsTable_LISTE_MODES.Filter = ("")
    If (DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.RecordCount + DataEnvironment1.rsTable_LISTE_MODES.RecordCount) < 1 Then
        GoTo suite2
    End If
    DataEnvironment1.rsTable_LISTE_MODES.MoveFirst
        While DataEnvironment1.rsTable_LISTE_MODES.EOF <> True
            cmbModeClaqu.AddItem (CStr(DataEnvironment1.rsTable_LISTE_MODES.Fields("Mode").Value))
            DataEnvironment1.rsTable_LISTE_MODES.MoveNext
        Wend
    DataEnvironment1.rsTable_LISTE_MODES.Filter = ("code_mode='" & DataEnvironment1.rsTable_MODE_DE_MARCHE_CLAQUAGE.Fields("MODE_ACTUEL_CLAQUAGE").Value) & "'"
    cmbModeClaqu.Text = CStr(DataEnvironment1.rsTable_LISTE_MODES.Fields("Mode").Value)
suite2:
    'remplissage champs fréquentiels
    '' Filling Frequency Fields
    Call remplit_champs_frequentiels
    FraDataClaqu.Visible = False
    'remplissage combo symbole produits de référence pour le test station de claquage
    '' Fill combo for product symbole reference for Hipot Test station.
    Call remplit_champs_symboles
End Sub

Private Sub Form_Unload(Cancel As Integer)
If DataEnvironment1.rsTable_ID_QUALITE.State = adStateOpen Then
    DataEnvironment1.rsTable_ID_QUALITE.Close
End If
End Sub

Private Sub txtIntervalle_Change()
If UtilisateurConnecte And (ProfilUtil > 0) Then
    btnEnrData.Enabled = True
End If
End Sub
Private Sub txtProchaineExecution_Change()
ProchaineDateDemandeChangee = True
If UtilisateurConnecte And (ProfilUtil > 0) Then
    btnEnrData.Enabled = True
End If
End Sub
'=======FIN MODIF BEEA 12/02/2008==================

