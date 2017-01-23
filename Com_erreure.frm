VERSION 5.00
Begin VB.Form Com_erreure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communication Error"
   ClientHeight    =   2520
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   2892
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4560
      Top             =   960
   End
   Begin VB.Label Titre 
      Alignment       =   2  'Center
      Caption         =   "Communication Error to PLC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5292
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5412
   End
End
Attribute VB_Name = "Com_erreure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEY_ESCAPE Or KeyCode = vbKeyNumpad0 Then
       End
    End If

End Sub

Private Sub Form_Load()
'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
SetWindowPos hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Timer1_Timer()
    'Anime le clignotement de couleur du titre
    If Titre.ForeColor = vbRed Then
        Titre.ForeColor = vbWhite
    Else
        Titre.ForeColor = vbRed
    End If
End Sub
