VERSION 5.00
Begin VB.Form mk_add 
   Caption         =   "Добавления в список Шиндлера"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3960
   Icon            =   "mk_add.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox mk_kto 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   3135
   End
   Begin VB.ComboBox mk_gr 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.ComboBox mk_vk 
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox mk_kyda 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox mk_add_fio 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Кто"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Куда"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Военный Камиссариат"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Год рождения"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Фамилия, Имя, отчество"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "mk_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommAND1_Click()

Dim max_id As Integer
max_id = "0"
mysql.query ("SELECT max(id) FROM control_" & nowBase & "")
If st = "0" Then
max_id = "1"
Else
max_id = DAT(1, 1) + 1
End If
If Len(mk_add_fio.Text) > 0 Then
    If Len(mk_gr.Text) > 0 Then
        If Len(mk_vk.Text) > 0 Then
            If Len(mk_kyda.Text) > 0 Then
                If Len(mk_kto.Text) > 0 Then
                    mysql.query ("INSERT INTO control_" & nowBase & " values ('" & max_id & "','" & mk_add_fio & "','" & mk_gr.Text & "','" & mk_vk.Text & "','" & mk_kyda & "','" & mk_kto & "','0','0')")
                    Unload mk_add
                End If
            End If
        End If
    End If
End If
menkontrol.gen_list (2)
End Sub

Private Sub FORm_Load()
Dim x As Long
Call Reg_VK_List
For x = 0 To UBound(nVK())
mk_vk.AddItem (nVK(x))
Next x
For x = 79 To 89
mk_gr.AddItem ("19" & x)
Next x
End Sub

Private Sub mk_add_fio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
