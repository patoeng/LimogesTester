VERSION 5.00
Begin VB.Form FRM_propos 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Final Test XS Bench"
   ClientHeight    =   6225
   ClientLeft      =   135
   ClientTop       =   435
   ClientWidth     =   7185
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1931.814
   ScaleMode       =   0  'User
   ScaleWidth      =   7684.686
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   2055
      Left            =   3120
      TabIndex        =   20
      Top             =   3600
      Width           =   3975
      Begin VB.Label Label16 
         Caption         =   "167398601A77 IED : 04"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label21 
         Caption         =   "Ajout contrôle périodique de références de claquage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label18 
         Caption         =   "N° affaire BEEA : 07006"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Modifié par P. Ardouin le 16/07/2008"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   21
         Top             =   1200
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   2892
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   240
         Picture         =   "FRM_VERS.frx":0000
         ScaleHeight     =   975
         ScaleWidth      =   2415
         TabIndex        =   24
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "5 av. Jeanne Pichenaud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "87700 AIXE SUR VIENNE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Tél/Fax : 05 55 70 07 63"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label12 
         Caption         =   "Mail : sarl.beea@beea.fr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin VB.Image Image1 
         Height          =   2295
         Left            =   240
         Picture         =   "FRM_VERS.frx":1007
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Fax : 05 45 24 87 22"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Tél : 05 45 24 87 21"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "16 410 GARAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Route de périgeux"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Version N° 10"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Modifié : CLERC Alban"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Label Label11 
         Caption         =   "N° affaire sodatec : 794"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Remerciements : P.Baudinaud"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label9 
         Caption         =   "Auteurs : W.Simonneau"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "N° dossier =S= : 3A 409 1336"
         Height          =   255
         Left            =   225
         TabIndex        =   8
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Banc de contrôle diélectrique et final des produits XS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1332
         Left            =   216
         TabIndex        =   7
         Top             =   120
         Width           =   3492
      End
   End
   Begin VB.CommandButton CMD_ok 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "F1 - &OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   2520
      TabIndex        =   0
      Top             =   5760
      Width           =   1095
   End
End
Attribute VB_Name = "FRM_propos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub CMD_ok_Click()
     
     Unload Me
     
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = vbKeyF1) Then
        CMD_ok_Click
    End If
    
End Sub

Private Sub Form_Load()
     
     Label3.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
     Label4.Caption = App.FileDescription
     Me.Show
    
End Sub

Private Sub Form_Paint()

'met la fenetre au premier plan, et on ne peut plus l'effacer !
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW

End Sub

