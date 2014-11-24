VERSION 5.00
Begin VB.Form frmEditMK 
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Сохранить"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   7080
      Width           =   3615
   End
   Begin VB.TextBox txtKTO 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   11
      Top             =   6360
      Width           =   3615
   End
   Begin VB.TextBox txtKYDA 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   9
      Top             =   5040
      Width           =   3615
   End
   Begin VB.TextBox txtGR 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox lstVK 
      BackColor       =   &H00C0FFC0&
      CausesValidation=   0   'False
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
      ItemData        =   "frmEditMK.frx":0000
      Left            =   240
      List            =   "frmEditMK.frx":0002
      TabIndex        =   5
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox txtFIO 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Центровка
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Кто"
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
      Left            =   240
      TabIndex        =   10
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Куда"
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
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Год Рождения"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Военный комиссариат"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Ф.И.О."
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
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Прозрачно
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmEditMK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommAND1_Click()
Call mysql.query("UPDATE control_" & nowBase & " set fio='" & txtfio.Text & "', gr='" & txtGR.Text & "', vk='" & lstvk.Text & "', kyda='" & txtKYDA.Text & "', kto='" & txtKTO.Text & "' WHERE id='" & txtID.Text & "'")
menkontrol.gen_list (2)
Unload Me
End Sub

Private Sub FORm_Load()
On Error Resume Next
mysql.query ("SELECT id,fio,gr,vk,kyda,kto FROM control_" & nowBase & " WHERE id='" & id_mk & "'")
txtID.Text = DAT(1, 1)
txtfio.Text = DAT(2, 1)
txtGR.Text = DAT(3, 1)
Call Reg_VK_List
For x = 0 To 45
    lstvk.AddItem (nVK(x))
Next x
lstvk.Text = DAT(4, 1)
txtKYDA.Text = DAT(5, 1)
txtKTO.Text = DAT(6, 1)

End Sub
Private Sub txtFIO_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
