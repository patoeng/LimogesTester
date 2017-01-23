VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form test_prog 
   BorderStyle     =   0  'None
   Caption         =   "test prog simu"
   ClientHeight    =   8112
   ClientLeft      =   1356
   ClientTop       =   1668
   ClientWidth     =   5772
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8112
   ScaleWidth      =   5772
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check6 
      Caption         =   "Ic max"
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   70
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Ic toujours 0"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   69
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Is max"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   68
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Is  toujour 0"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   67
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Ic max"
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   66
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Ic toujours 0"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   65
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Is max"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   64
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Is  toujour 0"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   63
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Ic max"
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   62
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Ic toujours 0"
      Height          =   255
      Index           =   13
      Left            =   2280
      TabIndex        =   61
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Is max"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   60
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Is  toujour 0"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   59
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Ic max"
      Height          =   255
      Index           =   14
      Left            =   2280
      TabIndex        =   58
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Ic toujours 0"
      Height          =   255
      Index           =   13
      Left            =   2280
      TabIndex        =   57
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Is max"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   56
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Is  toujour 0"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   55
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Ic max"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   54
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Ic toujours 0"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   53
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Is max"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   52
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Is  toujour 0"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   51
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox Check 
      Caption         =   "Ic max"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   50
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox Check 
      Caption         =   "Ic toujours 0"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   49
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox Check 
      Caption         =   "Is max"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   48
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check 
      Caption         =   "Is  toujour 0"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   47
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox bon_2 
      Caption         =   "produit 2 bon"
      Height          =   255
      Left            =   3960
      TabIndex        =   46
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CheckBox bon_1 
      Caption         =   "produit 1 bon"
      Height          =   255
      Left            =   2040
      TabIndex        =   45
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CheckBox prod_2 
      Caption         =   "présence produit 2"
      Height          =   255
      Left            =   3960
      TabIndex        =   44
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CheckBox prod_1 
      Caption         =   "présence produit 1"
      Height          =   255
      Left            =   2040
      TabIndex        =   43
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton VALID 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3135
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "test à U1"
      Begin VB.CheckBox Check 
         Caption         =   "Is à Sn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Check 
         Caption         =   "Is à Sn/2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "Is à Sn = 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check 
         Caption         =   "Ic à Sn/2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check 
         Caption         =   "Ic à Sn"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check 
         Caption         =   "Ic à  Sn = 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "test à U2"
      Begin VB.CheckBox Check2 
         Caption         =   "Ic à  Sn = 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Ic à Sn"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Ic à Sn/2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Is à Sn = 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Is à Sn/2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Is à Sn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   3135
      Index           =   2
      Left            =   2160
      TabIndex        =   15
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "prod 1 bon U1"
      Begin VB.CheckBox Check3 
         Caption         =   "Ic à  Sn = 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Ic à Sn"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Ic à Sn/2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Is à Sn = 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Is à Sn/2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Is à Sn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   3135
      Index           =   3
      Left            =   2160
      TabIndex        =   22
      Top             =   3480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "prod 1 bon U2"
      Begin VB.CheckBox Check4 
         Caption         =   "Is à Sn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Is à Sn/2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Is à Sn = 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Ic à Sn/2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Ic à Sn"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Ic à  Sn = 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   3135
      Index           =   4
      Left            =   3840
      TabIndex        =   29
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "prod 2 bon à U1"
      Begin VB.CheckBox Check5 
         Caption         =   "Ic à  Sn = 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Ic à Sn"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Ic à Sn/2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Is à Sn = 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Is à Sn/2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Is à Sn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   3135
      Index           =   5
      Left            =   3840
      TabIndex        =   36
      Top             =   3480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   5530
      _StockProps     =   14
      Caption         =   "prod 2 bon à U2"
      Begin VB.CheckBox Check6 
         Caption         =   "Is à Sn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Is à Sn/2"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Is à Sn = 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Ic à Sn/2"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Ic à Sn"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Ic à  Sn = 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "test_prog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub SSCheck1_Click(Value As Integer)

End Sub

Private Sub Check1_Click(Index As Integer)

End Sub


Private Sub VALID_Click()

Dim i As Long
Dim a As Long


'test pour U1
If Check(4).Value = 1 Then
set_bit Tab_Status_banc(30), 4
Else
reset_bit Tab_Status_banc(30), 4
End If
If Check(5).Value = 1 Then
set_bit Tab_Status_banc(30), 5
Else
reset_bit Tab_Status_banc(30), 5
End If
If Check(6).Value = 1 Then
set_bit Tab_Status_banc(30), 6
Else
reset_bit Tab_Status_banc(30), 6
End If
If Check(7).Value = 1 Then
set_bit Tab_Status_banc(30), 7
Else
reset_bit Tab_Status_banc(30), 7
End If
If Check(8).Value = 1 Then
set_bit Tab_Status_banc(30), 8
Else
reset_bit Tab_Status_banc(30), 8
End If
If Check(15).Value = 1 Then
set_bit Tab_Status_banc(30), 15
Else
reset_bit Tab_Status_banc(30), 15
End If
If Check(2).Value = 1 Then
set_bit Tab_Status_banc(30), 2
Else
reset_bit Tab_Status_banc(30), 2
End If
If Check(3).Value = 1 Then
set_bit Tab_Status_banc(30), 3
Else
reset_bit Tab_Status_banc(30), 3
End If
If Check(13).Value = 1 Then
set_bit Tab_Status_banc(30), 13
Else
reset_bit Tab_Status_banc(30), 13
End If
If Check(14).Value = 1 Then
set_bit Tab_Status_banc(30), 14
Else
reset_bit Tab_Status_banc(30), 14
End If

'test pour U2
If Check2(4).Value = 1 Then
set_bit Tab_Status_banc(31), 4
Else
reset_bit Tab_Status_banc(31), 4
End If
If Check2(5).Value = 1 Then
set_bit Tab_Status_banc(31), 5
Else
reset_bit Tab_Status_banc(31), 5
End If
If Check2(6).Value = 1 Then
set_bit Tab_Status_banc(31), 6
Else
reset_bit Tab_Status_banc(31), 6
End If
If Check2(7).Value = 1 Then
set_bit Tab_Status_banc(31), 7
Else
reset_bit Tab_Status_banc(31), 7
End If
If Check2(8).Value = 1 Then
set_bit Tab_Status_banc(31), 8
Else
reset_bit Tab_Status_banc(31), 8
End If
If Check2(15).Value = 1 Then
set_bit Tab_Status_banc(31), 15
Else
reset_bit Tab_Status_banc(31), 15
End If
If Check2(2).Value = 1 Then
set_bit Tab_Status_banc(31), 2
Else
reset_bit Tab_Status_banc(31), 2
End If
If Check2(3).Value = 1 Then
set_bit Tab_Status_banc(31), 3
Else
reset_bit Tab_Status_banc(31), 3
End If
If Check2(13).Value = 1 Then
set_bit Tab_Status_banc(31), 13
Else
reset_bit Tab_Status_banc(31), 13
End If
If Check2(14).Value = 1 Then
set_bit Tab_Status_banc(31), 14
Else
reset_bit Tab_Status_banc(31), 14
End If

' test bon produit 1
'test pour U1
i = 60
If Check3(4).Value = 1 Then
set_bit Tab_Status_banc(i), 4
Else
reset_bit Tab_Status_banc(i), 4
End If
If Check3(5).Value = 1 Then
set_bit Tab_Status_banc(i), 5
Else
reset_bit Tab_Status_banc(i), 5
End If
If Check3(6).Value = 1 Then
set_bit Tab_Status_banc(i), 6
Else
reset_bit Tab_Status_banc(i), 6
End If
If Check3(7).Value = 1 Then
set_bit Tab_Status_banc(i), 7
Else
reset_bit Tab_Status_banc(i), 7
End If
If Check3(8).Value = 1 Then
set_bit Tab_Status_banc(i), 8
Else
reset_bit Tab_Status_banc(i), 8
End If
If Check3(15).Value = 1 Then
set_bit Tab_Status_banc(i), 15
Else
reset_bit Tab_Status_banc(i), 15
End If
If Check3(2).Value = 1 Then
set_bit Tab_Status_banc(i), 2
Else
reset_bit Tab_Status_banc(i), 2
End If
If Check3(3).Value = 1 Then
set_bit Tab_Status_banc(i), 3
Else
reset_bit Tab_Status_banc(i), 3
End If
If Check3(13).Value = 1 Then
set_bit Tab_Status_banc(i), 13
Else
reset_bit Tab_Status_banc(i), 13
End If
If Check3(14).Value = 1 Then
set_bit Tab_Status_banc(i), 14
Else
reset_bit Tab_Status_banc(i), 14
End If

'test pour U2
If Check4(4).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 4
Else
reset_bit Tab_Status_banc(i + 1), 4
End If
If Check4(5).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 5
Else
reset_bit Tab_Status_banc(i + 1), 5
End If
If Check4(6).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 6
Else
reset_bit Tab_Status_banc(i + 1), 6
End If
If Check4(7).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 7
Else
reset_bit Tab_Status_banc(i + 1), 7
End If
If Check4(8).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 8
Else
reset_bit Tab_Status_banc(i + 1), 8
End If
If Check4(15).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 15
Else
reset_bit Tab_Status_banc(i + 1), 15
End If

If Check4(2).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 2
Else
reset_bit Tab_Status_banc(i + 1), 2
End If
If Check4(3).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 3
Else
reset_bit Tab_Status_banc(i + 1), 3
End If
If Check4(13).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 13
Else
reset_bit Tab_Status_banc(i + 1), 13
End If
If Check4(14).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 14
Else
reset_bit Tab_Status_banc(i + 1), 14
End If

' test bon produit 2
'test pour U1
i = 62
If Check5(4).Value = 1 Then
set_bit Tab_Status_banc(i), 4
Else
reset_bit Tab_Status_banc(i), 4
End If
If Check5(5).Value = 1 Then
set_bit Tab_Status_banc(i), 5
Else
reset_bit Tab_Status_banc(i), 5
End If
If Check5(6).Value = 1 Then
set_bit Tab_Status_banc(i), 6
Else
reset_bit Tab_Status_banc(i), 6
End If
If Check5(7).Value = 1 Then
set_bit Tab_Status_banc(i), 7
Else
reset_bit Tab_Status_banc(i), 7
End If
If Check5(8).Value = 1 Then
set_bit Tab_Status_banc(i), 8
Else
reset_bit Tab_Status_banc(i), 8
End If
If Check5(15).Value = 1 Then
set_bit Tab_Status_banc(i), 15
Else
reset_bit Tab_Status_banc(i), 15
End If
If Check5(2).Value = 1 Then
set_bit Tab_Status_banc(i), 2
Else
reset_bit Tab_Status_banc(i), 2
End If
If Check5(3).Value = 1 Then
set_bit Tab_Status_banc(i), 3
Else
reset_bit Tab_Status_banc(i), 3
End If
If Check5(13).Value = 1 Then
set_bit Tab_Status_banc(i), 13
Else
reset_bit Tab_Status_banc(i), 13
End If
If Check5(14).Value = 1 Then
set_bit Tab_Status_banc(i), 14
Else
reset_bit Tab_Status_banc(i), 14
End If

'test pour U2
If Check6(4).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 4
Else
reset_bit Tab_Status_banc(i + 1), 4
End If
If Check6(5).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 5
Else
reset_bit Tab_Status_banc(i + 1), 5
End If
If Check6(6).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 6
Else
reset_bit Tab_Status_banc(i + 1), 6
End If
If Check6(7).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 7
Else
reset_bit Tab_Status_banc(i + 1), 7
End If
If Check6(8).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 8
Else
reset_bit Tab_Status_banc(i + 1), 8
End If
If Check6(15).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 15
Else
reset_bit Tab_Status_banc(i + 1), 15
End If
If Check6(2).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 2
Else
reset_bit Tab_Status_banc(i + 1), 2
End If
If Check6(3).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 3
Else
reset_bit Tab_Status_banc(i + 1), 3
End If
If Check6(13).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 13
Else
reset_bit Tab_Status_banc(i + 1), 13
End If
If Check6(14).Value = 1 Then
set_bit Tab_Status_banc(i + 1), 14
Else
reset_bit Tab_Status_banc(i + 1), 14
End If

If prod_1.Value = 1 Then
set_bit Tab_Status_banc(6), 7
Else
reset_bit Tab_Status_banc(6), 7
End If
If prod_2.Value = 1 Then
set_bit Tab_Status_banc(6), 8
Else
reset_bit Tab_Status_banc(6), 8
End If
If bon_1.Value = 1 Then
set_bit Tab_Status_banc(6), 11
Else
reset_bit Tab_Status_banc(6), 11
End If
If bon_2.Value = 1 Then
set_bit Tab_Status_banc(6), 12
Else
reset_bit Tab_Status_banc(6), 12
End If


'Me.Unload
End Sub


