VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmusers 
   Caption         =   "Управление пользователями"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmusers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton del 
      Caption         =   "Удалить"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Отмена"
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton save 
      Caption         =   "Сохранить"
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton add 
      Caption         =   "Добавить"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Информация о пользователе"
      Height          =   4575
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cmbaccess 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtcomment 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtfio 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   10
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtpass 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtname 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label txtid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Коментарии"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Уровень доступа"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Пароль"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Фамилия, имя и отчество"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Имя пользователя"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView listusers 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Логин"
         Object.Width           =   3351
      EndProperty
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Long
Dim Y As Long
Dim chg As Boolean

Private Sub add_Click()

frmuser_add.Show vbModal, Me
End Sub

Private Sub cmbaccess_Change()
chg = True
End Sub

Private Sub CommAND2_Click()

End Sub

Private Sub CommAND3_Click()
Unload Me
End Sub

Private Sub del_Click()
If MsgBox("Вы уверены что хотите удалить пользователя " & txtname.Text & "?", vbOKCancel + vbQuestion, "Удаления пользователя") = vbOK Then
    Call mysql.query("DELETE FROM users WHERE name='" & listusers.ListItems.Item(listusers.SelectedItem.Index).SubItems(1) & "'")
    MsgBox "Пользователь " & txtname.Text & " был удачно удален из базы", vbOKOnly + vbInformation, "Удаление пользователя"
    Call Form_Load
End If
End Sub

Public Sub Form_Load()
listusers.ListItems.Clear
Call mysql.query("SELECT * FROM users ORder by name ASC")
    For x = 1 To st
        Set LF = listusers.ListItems.add(1, , DAT(1, x))
        LF.SubItems(1) = DAT(2, x)
    Next x
                        cmbaccess.AddItem ("Только чтение")
                        cmbaccess.AddItem ("Обычный")
                        cmbaccess.AddItem ("Администратор")
listusers.ListItems.Item(1).Selected = True
Call listusers_Click
End Sub
Private Sub listusers_Click()
    On Error Resume Next
    
    Call mysql.query("SELECT * FROM users WHERE name='" & listusers.ListItems.Item(listusers.SelectedItem.Index).SubItems(1) & "'")
    txtid.Caption = DAT(1, 1)
    txtname.Text = DAT(2, 1)
    txtpass.Text = DAT(3, 1)
    If DAT(4, 1) = "s" Then cmbaccess.ListIndex = "0"
    If DAT(4, 1) = "O" Then cmbaccess.ListIndex = "1"
    If DAT(4, 1) = "G" Then cmbaccess.ListIndex = "2"
    txtfio.Text = DAT(5, 1)
    txtcomment.Text = DAT(6, 1)
    
End Sub

Private Sub save_user()
Dim acs As String
        Call mysql.query("DELETE FROM users WHERE name='" & listusers.ListItems.Item(listusers.SelectedItem.Index).SubItems(1) & "'")
        If cmbaccess.ListIndex = "0" Then acs = "s"
        If cmbaccess.ListIndex = "1" Then acs = "O"
        If cmbaccess.ListIndex = "2" Then acs = "G"
        Call mysql.query("insert into users VAlues ('" & txtid.Caption & "','" & txtname.Text & "','" & txtpass.Text & "','" & acs & "','" & txtfio.Text & "', '" & txtcomment.Text & "')")
End Sub
Private Sub save_Click()
Call save_user
End Sub

Private Sub txtcomment_Change()
chg = True
End Sub

Private Sub txtfio_Change()
chg = True
End Sub

Private Sub txtname_Change()
chg = True
End Sub

Private Sub txtpass_Change()
chg = True
End Sub

