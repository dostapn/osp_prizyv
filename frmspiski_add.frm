VERSION 5.00
Begin VB.Form frmspiski_add 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4305
   Icon            =   "frmspiski_add.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Years 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Text            =   "��� ��������"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "������"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Add_pr 
      Caption         =   "��������"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtkto 
      Height          =   350
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   4000
   End
   Begin VB.ComboBox lstvk 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "���������"
      Top             =   3360
      Width           =   4000
   End
   Begin VB.TextBox txtto 
      Height          =   350
      Left            =   120
      TabIndex        =   6
      Top             =   4320
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
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label txtid 
      Caption         =   "ID"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "������"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   4000
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   4000
   End
   Begin VB.Label Label3 
      Caption         =   "���������"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4000
   End
   Begin VB.Label Label2 
      Caption         =   "��� ��������"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4000
   End
   Begin VB.Label Label1 
      Caption         =   "������� ��� ��������"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4000
   End
End
Attribute VB_Name = "frmspiski_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_pr_Click()
If txtid.Caption = "ID" Then
mysql.query ("insert into spiski_" & nowBase & " set fio = '" & txtfio.Text & "', gr= '" & Years.Text & "', kto = '" & txtkto.Text & "', kyda = '" & txtto.Text & "', spiski_" & nowBase & ".vk = '" & lstvk.Text & "'")
MsgBox ("��������� " & txtfio.Text & " ��������.")
Else
mysql.query ("Update spiski_" & nowBase & " set fio = '" & txtfio.Text & "', gr= '" & Years.Text & "', kto = '" & txtkto.Text & "', kyda = '" & txtto.Text & "', spiski_" & nowBase & ".vk = '" & lstvk.Text & "' where id = '" & txtid.Caption & "'")
MsgBox ("���������� ��������.")
End If
Call frmspiskikomp.gen_list("", "")
Unload Me
End Sub

Private Sub Cancel_Click()
If MsgBox("�� ��������� ������?", vbExclamation + vbOKCancel) = vbOK Then
Unload Me
Else
Call Add_pr_Click
End If
End Sub

Private Sub Form_Load()

'''''''''''''''''������ �����������
Call Reg_VK_List
For x = 0 To UBound(nVK())
    lstvk.AddItem (nVK(x))
Next x
''''''''''''''''

'''''''''''''''' ������ ���
For x = 0 To 11
    Years.AddItem (Format(Now - ((16 + x) * 365), "YYYY"))
Next x
Years.Text = (Format(Now - (16 * 365), "YYYY"))
'''''''''''''''
End Sub
