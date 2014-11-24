VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnaryad_srezki_sel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Срезки"
   ClientHeight    =   3585
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "frmnaryad_srezki_sel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Тип срезок"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   4335
      Begin VB.OptionButton opt3 
         Caption         =   "Срезки + План"
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Срезки"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Срезки на сегодня"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
      Begin VB.CommandButton cmd_date 
         Caption         =   "Печать"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Срезки на дату"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmd_to_date 
         Caption         =   "Печать"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16056321
         CurrentDate     =   39290
      End
      Begin VB.Label Label1 
         Caption         =   "Выбирите дату для срезок"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmnaryad_srezki_sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim type_d As Long
Private Sub cmd_date_Click()
Call type_s
Call Cnv.To_naryad_srezki(DTPicker1, type_d)
End Sub

Private Sub cmd_to_date_Click()
Call type_s
Call Cnv.To_naryad_srezki(DTPicker1, type_d)
End Sub
Private Sub type_s()
If opt1 = True Then type_d = 1
If opt2 = True Then type_d = 2
If opt3 = True Then type_d = 3
End Sub
Private Sub Form_Load()
DTPicker1 = Date
End Sub

