VERSION 5.00
Begin VB.Form frmuser_add 
   Caption         =   "Добавление Пользователя"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancel 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton add 
      Caption         =   "Добавить"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtcom 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtfio 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Коментарии"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   100
      TabIndex        =   5
      Top             =   3720
      Width           =   1740
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Фамилия, имя и отчество"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   100
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Уровень доступа"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   100
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Пароль"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   100
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Имя"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   100
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   100
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmuser_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_id As Long
Dim acs As String
Private Sub add_Click()
If Len(txtname.Text) < "1" Or Len(txtpass.Text) < "1" Then
    MsgBox "Поля ИМЯ и ПАРОЛЬ должны быть введены обязательно", vbCritical, "Добавление пользователя"
Else
        If Combo1.ListIndex = "0" Then acs = "s"
        If Combo1.ListIndex = "1" Then acs = "O"
        If Combo1.ListIndex = "2" Then acs = "G"
        Call mysql.query("insert into users VAlues ('" & new_id & "','" & txtname.Text & "','" & txtpass.Text & "','" & acs & "','" & txtfio.Text & "', '" & txtcom.Text & "')")
        frmuser_add.Hide
        MsgBox "Пользователь " & txtname.Text & " был успешно добавлен в базу", vbInformation, "Добавление пользователя"
        frmusers.Form_Load
        Unload frmuser_add
End If
End Sub

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call mysql.query("SELECT max(id) FROM users")
new_id = DAT(1, 1) + 1
Label7.Caption = new_id
    Combo1.AddItem ("Только чтение")
    Combo1.AddItem ("Обычный")
    Combo1.AddItem ("Администратор")
End Sub

