VERSION 4.00
Begin VB.Form FRM_attention 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2895
   ClientLeft      =   2580
   ClientTop       =   2580
   ClientWidth     =   5985
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   3300
   KeyPreview      =   -1  'True
   Left            =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   Top             =   2235
   Width           =   6105
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   465
      Left            =   255
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   2805
   End
   Begin VB.CommandButton CMD_annuler 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "F5-&Retour au contrôle"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1755
      Width           =   2235
   End
   Begin VB.CommandButton CMD_continuer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "F1-&Continuer"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2295
      Width           =   2220
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"FRM_ATTE.frx":0000
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1035
      Index           =   0
      Left            =   120
      Picture         =   "FRM_ATTE.frx":00AF
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "FRM_attention"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD_annuler_Click()

    go_config = False
    Unload Me

End Sub

Private Sub CMD_continuer_Click()

        go_config = True
        Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEY_F1 Then
        CMD_continuer_Click
    End If

    If KeyCode = KEY_F5 Or KeyCode = vbKeyDivide Then
        CMD_annuler_Click
    End If

End Sub

Private Sub Form_Paint()
    
    text1.SetFocus

End Sub

