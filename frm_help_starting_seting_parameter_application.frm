VERSION 5.00
Begin VB.Form frm_help_seting_parameter_application 
   Caption         =   "HELP Argument Parameter to start the application"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5895
   Icon            =   "frm_help_starting_seting_parameter_application.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5895
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frm_help_starting_seting_parameter_application.frx":0442
      ScaleHeight     =   2655
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   120
      Width           =   5652
      Begin VB.Label Text_aide 
         BackColor       =   &H00FFFFFF&
         Height          =   2412
         Left            =   540
         TabIndex        =   2
         Top             =   180
         Width           =   4932
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "frm_help_seting_parameter_application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index dans la collection de l'astuce actuellement affichés.
Dim CurrentTip As Long


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Caption = Me.Caption & " " & App.Title
        
    Text_aide.Caption = App.EXEName & " Param1 Param2 Param n-1 Param n "
    
    Text_aide.Caption = Text_aide.Caption & vbCrLf
    Text_aide.Caption = Text_aide.Caption & vbCrLf & " Param1 Param2 Param n-1"
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "HELP MAX_WINDOWS "
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "SCREEN_PRINT PRINTER_PRINT "
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "LOG_DATA_BASE_REGISTER_NUMBER_VISIBLE "
    
    
    Text_aide.Caption = Text_aide.Caption & vbCrLf
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "Param n"
    'Text_aide.Caption = Text_aide.Caption & vbCrLf & "Fichier de configuration"
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "Configuration FILE"
    
    Text_aide.Caption = Text_aide.Caption & vbCrLf
    'Text_aide.Caption = Text_aide.Caption & vbCrLf & "Exemple"
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "Example"
    Text_aide.Caption = Text_aide.Caption & vbCrLf & "control WINDOWS_MAX c:\banc_xs\config.ini"

End Sub


