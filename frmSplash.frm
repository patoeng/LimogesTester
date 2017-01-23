VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   ClientHeight    =   4245
   ClientLeft      =   270
   ClientTop       =   1335
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6240
         Top             =   1440
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Width           =   6732
         _ExtentX        =   11880
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   1485
         Left            =   1320
         Picture         =   "frmSplash.frx":000C
         Top             =   240
         Width           =   5520
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   144
         TabIndex        =   1
         Top             =   3120
         Width           =   2988
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Produit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   2952
         TabIndex        =   2
         Top             =   1800
         Width           =   1188
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Unload Me
End Sub

Private Sub Form_Load()
    SetWindowPos hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    
End Sub

Private Sub Frame1_Click()
    'Unload Me
End Sub

Private Sub Timer1_Timer()
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Min
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
    End If
    
End Sub
