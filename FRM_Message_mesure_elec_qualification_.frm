VERSION 5.00
Begin VB.Form FRM_Message_mes_electrique_qualification 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Measuring Electrical for qualification"
   ClientHeight    =   2670
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Press start button step by step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   732
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   6372
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "zeeezezeze  ez ze e ez e eze zezez e zez   zez  zee zez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Electrical Measurement is paused"
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
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6372
   End
End
Attribute VB_Name = "FRM_Message_mes_electrique_qualification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_SHOWWINDOW
    SetWindowPos hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Label2.Caption = ""
End Sub

Public Sub Set_message(Code_message As Integer)
Dim Message As String

    Select Case Code_message
        Case 600
        'Message = "Mesure Ic"
        Message = "Measure Ic"
        
        Case 605
        'Message = "Mesure Ud ou If charge 1"
        Message = "Measure Vd or ILoad 1"
        
        Case 610
        'Message = "Mesure Ud ou If charge 2"
        Message = "Measure Vd or ILoad 2"
        
        Case 616
        'Message = "Mesure Ic à 0 Sn"
        Message = "Measure Ic at 0 Sn"
        
        Case 617
        'Message = "Mesure Ic à 1/2 Sn "
        Message = "Measure Ic at 1/2 Sn "
        
        Case 618
        'Message = "Mesure Ic à Sn"
        Message = "Measure Ic at Sn"
        
        Case 621
        'Message = "Mesure Is à 0 Sn"
        Message = "Measure Is at 0 Sn"
        
        Case 622
        'Message = "Mesure Is à 1/2 Sn "
        Message = "Measure Is at 1/2 Sn "
        
        Case 623
        'Message = "Mesure Is à Sn"
        Message = "Measure Is at Sn"
        
        Case 626
        'Message = "Mesure Is + Ic à 0 Sn"
        Message = "Measure Is + Ic at 0 Sn"
        
        Case 627
        'Message = "Mesure Is + Ic à 1/2 Sn "
        Message = "Measure Is + Ic at 1/2 Sn "
        
        Case 628
        'Message = "Mesure Is + Ic à Sn"
        Message = "Mesure Is + Ic at Sn"
        
        Case Else
            Message = "?????"
    End Select
    Label2.Caption = Message
End Sub

