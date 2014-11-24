VERSION 5.00
Begin VB.Form frmdir_add 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Добавление директивщика"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4305
   Icon            =   "frmdir_add.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtkto 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   3855
   End
   Begin VB.ComboBox txtto 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4320
      Width           =   3975
   End
   Begin VB.ComboBox Years 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Text            =   "Год рождения"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Add_pr 
      Caption         =   "Добавить"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   1935
   End
   Begin VB.ComboBox lstvk 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "Военкомат"
      Top             =   3360
      Width           =   4000
   End
   Begin VB.TextBox txtfio 
      Height          =   350
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4000
   End
   Begin VB.Label lid 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label txtid 
      Caption         =   "ID"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Директива"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   4000
   End
   Begin VB.Label Label4 
      Caption         =   "В какую часть"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   4000
   End
   Begin VB.Label Label3 
      Caption         =   "Военкомат"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4000
   End
   Begin VB.Label Label2 
      Caption         =   "Год рождения"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4000
   End
   Begin VB.Label Label1 
      Caption         =   "Фамилия Имя отчество"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4000
   End
End
Attribute VB_Name = "frmdir_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_pr_Click()
If txtid.Caption = "ID" Then
mysql.query ("insert into directivi_" & nowBase & " set fio = '" & txtfio.Text & "', gr= '" & Years.Text & "', kto = '" & txtkto.Text & "', kyda = '" & txtto.Text & "', directivi_" & nowBase & ".vk = '" & lstvk.Text & "'")
Call mysql.query("update naryad_" & nowBase & " set major = (select count(*) from directivi_" & nowBase & " where directivi_" & nowBase & ".kyda = '" & txtto.Text & "') where naryad_" & nowBase & ".vch = '" & txtto.Text & "'")
MsgBox ("Призывник " & txtfio.Text & " добавлен.")
Else
mysql.query ("Update directivi_" & nowBase & " set fio = '" & txtfio.Text & "', gr= '" & Years.Text & "', kto = '" & txtkto.Text & "', kyda = '" & txtto.Text & "', directivi_" & nowBase & ".vk = '" & lstvk.Text & "' where id = '" & txtid.Caption & "'")
MsgBox ("Информация изменена.")
End If

Call frmdirectivi.gen_list("", "")
Unload Me
End Sub

Private Sub Cancel_Click()
If MsgBox("Не добавлять запись?", vbExclamation + vbOKCancel) = vbOK Then
Unload Me
Else
Call Add_pr_Click
End If
End Sub

Private Sub Form_Load()

'''''''''''''''''Список военкоматов
Call Reg_VK_List
For x = 0 To UBound(nVK())
    lstvk.AddItem (nVK(x))
Next x
''''''''''''''''

'''''''''''''''' Список дат
For x = 0 To 11
    Years.AddItem (Format(Now - ((16 + x) * 365), "YYYY"))
Next x
Years.Text = (Format(Now - (16 * 365), "YYYY"))
'''''''''''''''



Call mysql.query("SELECT vch FROM naryad_" & nowBase & " group by vch")
For x = 1 To st
txtto.AddItem (DAT(1, x))
Next x

Call mysql.query("SELECT kto FROM directivi_" & nowBase & " group by kto")
For x = 1 To st
txtkto.AddItem (DAT(1, x))
Next x
End Sub
