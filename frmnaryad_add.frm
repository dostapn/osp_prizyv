VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmnaryad_add 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Добавление рода войск"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   4560
      TabIndex        =   9
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4683
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Род войск"
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
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Подрод войск"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtrodv 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   4335
      Begin VB.Label Label3 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Основной Род войск"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Закрыть"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Название"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "frmnaryad_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommAND1_Click()
On Error Resume Next
Dim newid, major, newmajOR, newminor, posled As Integer
If opt1 = True Then
    'check tables
   Call mysql.query("SELECT id FROM naryad_rodv_" & nowBase & " WHERE name='" & txtrodv & "'")
    If st > 0 Then
        MsgBox "Данный вид(род) войск уже имеется в базе", vbCritical, "Внимание"
        txtrodv = ""
        Call rodv_refresh
    Else
        Call mysql.query("SELECT max(id) FROM naryad_rodv_" & nowBase)
        If DAT(1, 1) = "" Then
            newid = 1
        Else
            newid = DAT(1, 1) + 1
        End If
        
        Call mysql.query("SELECT max(major) FROM naryad_rodv_" & nowBase)
        If DAT(1, 1) = "" Then
            newmajOR = 1
        Else
            newmajOR = DAT(1, 1) + 1
        End If
        
        Call mysql.query("SELECT max(posled) FROM naryad_rodv_" & nowBase & " WHERE minor='1'")
        If DAT(1, 1) = "" Then
            posled = 1
        Else
            posled = DAT(1, 1) + 1
        End If
                
                
        Call mysql.query("insert into naryad_rodv_" & nowBase & " ( `id` , `name` , `major` , `minor`,`posled`) values('" & newid & "','" & txtrodv & "','" & newmajOR & "','1','" & posled & "')")
        Call rodv_refresh
    End If
       
    Else
    Dim root_rodv As String
    root_rodv = TreeView1.SelectedItem.Text
    If Len(root_rodv) > 0 Then
        Call mysql.query("SELECT id FROM naryad_rodv_" & nowBase & " WHERE name='" & txtrodv & "'")
            If st > 0 Then
                MsgBox "Данный подвид(подрод) войск уже имеется в базе", vbCritical, "Внимание"
                txtrodv = ""
                Call rodv_refresh
            Else
                Call mysql.query("SELECT max(id) FROM naryad_rodv_" & nowBase)
                newid = DAT(1, 1)
                newid = newid + 1
                Call mysql.query("SELECT major FROM naryad_rodv_" & nowBase & " WHERE name='" & root_rodv & "'")
                major = DAT(1, 1)
                Call mysql.query("SELECT max(minor) FROM naryad_rodv_" & nowBase & " WHERE majOR='" & major & "'")
                newminor = DAT(1, 1) + 1
                
                Call mysql.query("insert into naryad_rodv_" & nowBase & " ( `id` , `name` , `major` , `minor`, `posled`) values('" & newid & "','" & txtrodv & "','" & major & "','" & newminor & "','0')")
                txtrodv = ""
                Call rodv_refresh
                
            End If
    End If
 End If
 Call frmNaryad.rodv_refresh
End Sub
Private Sub rodv_refresh()
On Error Resume Next
Dim nodX As Node
Dim datt() As String
TreeView1.Nodes.Clear
 Call mysql.query("SELECT `name`,`major`,`posled` FROM naryad_rodv_" & nowBase & " WHERE minor='1' ORder by posled ASC")
 datt() = DAT()
    For x = 1 To st
        Set nodX = TreeView1.Nodes.add(, , "r" & datt(3, x), datt(1, x))
         Call mysql.query("SELECT `name`,`id` FROM naryad_rodv_" & nowBase & " WHERE majOR='" & datt(2, x) & "' AND minor > '1'")
            If st > 0 Then
                For y = 1 To st
                    Set nodX = TreeView1.Nodes.add("r" & datt(3, x), tvwChild, "c" & DAT(2, y), DAT(1, y))
                Next y
            End If
            nodX.EnsureVisible
    Next x
  nodX.EnsureVisible
        

End Sub

Private Sub CommAND2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call rodv_refresh
TreeView1.Enabled = False
End Sub

Private Sub listres_Click()
Label3.Caption = listres.ListItems(listres.SelectedItem.Index).SubItems(2)
End Sub

Private Sub opt1_Click()
TreeView1.Enabled = False

End Sub

Private Sub opt2_Click()
TreeView1.Enabled = True
End Sub

Private Sub TreeView1_Click()
On Error Resume Next
If TreeView1.SelectedItem.Parent = "Nothing" Then
    Label3.Caption = TreeView1.SelectedItem.Text
Else
    Label3.Caption = TreeView1.SelectedItem.Parent
End If
End Sub
