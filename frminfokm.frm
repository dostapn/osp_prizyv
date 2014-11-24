VERSION 5.00
Begin VB.Form Frmkminfo 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Инфо"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdact 
      Caption         =   "Заблокировать"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
   End
   Begin VB.ComboBox ypk_vibor 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Дата привоза"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label datapr 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4920
      TabIndex        =   9
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label vk_i 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label datar 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Дата рождения"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Военкомат"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Выбор УПК"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label fio_gen 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Фамилия Имя Отчество"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "Frmkminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public table As String
Public type_lock As Boolean
Private Sub lock_prnik()

End Sub

Private Sub cmdact_Click()
Call lock_prnik
End Sub

Private Sub Command2_Click()
On Error Resume Next
If ypk_vibor.Text > "0" Or ypk_vibor.Text < "99999" Then
    
    End If
End Sub

Private Sub ypk_vibor_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub command2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub command1_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub ypk_vibor_Click()
Command2.Enabled = True



End Sub




