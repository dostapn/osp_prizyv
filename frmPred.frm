VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPred 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "С предназначением"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49152001
      CurrentDate     =   41228
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49152001
      CurrentDate     =   41228
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Сгенирировать"
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ПО"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "С"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmPred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommAND1_Click()
Dim g_f As String
Dim g_t As String
g_f = DTPicker1
g_t = DTPicker2
Call Cnv.To_List_gen_pred(g_f, g_t)
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.VAlue = Format(Now - 1, "DD.MM.YYYY")
DTPicker2.VAlue = Format(Now - 1, "DD.MM.YYYY")
End Sub
