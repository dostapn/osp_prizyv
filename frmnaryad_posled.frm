VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnaryad_posled 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   1695
      Left            =   6120
      TabIndex        =   4
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   2990
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4215
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "123"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "321"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   1695
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   2990
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "123"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "321"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Подроды"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Основные роды войск"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmnaryad_posled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rodv, podrod As String
Dim ID, posled As Integer
Dim x, i, j As Long
Private minor, major As String
Private Sub Form_Load()
On Error Resume Next
ListView1.ListItems.Clear
Call mysql.query("SELECT `name`,`id` FROM naryad_rodv_" & nowBase & " WHERE `minor`='1' ORder by posled ASC")
For x = 1 To st
        Set LF = ListView1.ListItems.add(, , DAT(1, x))
        LF.SubItems(1) = DAT(2, x)
Next x
Call ListView1_Click
End Sub

Private Sub ListView1_Click()
On Error Resume Next
ListView2.ListItems.Clear

rodv = ListView1.SelectedItem.Text
i = ListView1.SelectedItem.SubItems(1)
'If Len(rodv) = 0 Then End Sub
Call mysql.query("SELECT `id`,`name` from naryad_rodv_" & nowBase & " WHERE major=(SELECT `major` from naryad_rodv_" & nowBase & " WHERE `id`='" & i & "') and `minor` > '1' ORDER BY `minor` ASC")
If st > 0 Then
    For x = 1 To st
        Set DF = ListView2.ListItems.add(, , DAT(2, x))
        DF.SubItems(1) = DAT(1, x)
    Next x
ListView2.Refresh
End If

End Sub

Private Sub UpDown1_DownClick()
On Error Resume Next
Static rodv As String
rodv = ListView1.SelectedItem.Text
'i = ListView1.SelectedItem.
Call mysql.query("SELECT id,posled FROM naryad_rodv_" & nowBase & " WHERE name='" & rodv & "'")
ID = DAT(1, 1)
posled = DAT(2, 1)

Call mysql.query("SELECT id FROM naryad_rodv_" & nowBase & " WHERE posled='" & posled + 1 & "'")
If st = 0 Then
    Exit Sub
Else
Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set posled='" & posled & "' WHERE id='" & DAT(1, 1) & "'")
Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set posled='" & posled + 1 & "' WHERE id='" & ID & "'")

End If
Call Form_Load
ListView1.FindItem(rodv).Selected = True
ListView1.FindItem(rodv).EnsureVisible
Call ListView1_Click
End Sub
Private Sub UpDown1_upClick()
On Error Resume Next
Static rodv As String
rodv = ListView1.SelectedItem.Text
i = ListView1.SelectedItem.SubItems(1)
Call mysql.query("SELECT id,posled FROM naryad_rodv_" & nowBase & " WHERE name='" & rodv & "'")
ID = DAT(1, 1)
posled = DAT(2, 1)
If posled = "1" Then Exit Sub
Call mysql.query("SELECT id FROM naryad_rodv_" & nowBase & " WHERE posled='" & posled - 1 & "' and `minor`='1'")
If st = 0 Then
    Exit Sub
Else
Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set posled='" & posled & "' WHERE id='" & DAT(1, 1) & "'")
Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set posled='" & posled - 1 & "' WHERE id='" & ID & "'")

End If
Call Form_Load
ListView1.FindItem(rodv).Selected = True
ListView1.FindItem(rodv).EnsureVisible
Call ListView1_Click
End Sub

Private Sub UpDown2_UpClick()
On Error Resume Next
Static podrod As String
Dim oldid() As String
Dim newid() As String
Dim idnew, stt, stt2 As Long
podrod = ListView2.SelectedItem.Text

Call mysql.query("SELECT `id`,`minor`,`major` FROM naryad_rodv_" & nowBase & " WHERE `name`='" & podrod & "'")
ID = DAT(1, 1)
minor = DAT(2, 1)
major = DAT(3, 1)
If minor = "2" Then Exit Sub
Call mysql.query("SELECT id FROM naryad_rodv_" & nowBase & " WHERE minor='" & minor - 1 & "' AND `major`='" & major & "'")
idnew = DAT(1, 1)
If st = 0 Then
    Exit Sub
Else
Call mysql.query("SELECT `narid` from `naryad_" & nowBase & "` WHERE `major`='" & major & "' and `minor`='" & minor & "'")
oldid() = DAT()
stt = st
Call mysql.query("SELECT `narid` from `naryad_" & nowBase & "` WHERE `major`='" & major & "' and `minor`='" & minor - 1 & "'")
newid() = DAT()
stt2 = st

Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set minor='" & minor & "' WHERE id='" & idnew & "'")
For x = 1 To stt2
Call mysql.query("UPDATE naryad_" & nowBase & " set minor='" & minor & "' WHERE `narid`='" & newid(1, x) & "'")
Next x

Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set minor='" & minor - 1 & "' WHERE id='" & ID & "'")
For x = 1 To stt
Call mysql.query("UPDATE naryad_" & nowBase & " set minor='" & minor - 1 & "' WHERE `narid`='" & oldid(1, x) & "'")
Next x
End If
Call ListView1_Click
ListView2.FindItem(podrod).Selected = True
ListView2.FindItem(podrod).EnsureVisible
Call ListView1_Click
End Sub
Private Sub UpDown2_DownClick()
On Error Resume Next
Static podrod As String
Dim oldid() As String
Dim newid() As String
Dim idnew, stt, stt2 As Long
podrod = ListView2.SelectedItem.Text
Call mysql.query("SELECT `id`,`minor`,`major` FROM naryad_rodv_" & nowBase & " WHERE `name`='" & podrod & "'")
ID = DAT(1, 1)
minor = DAT(2, 1)
major = DAT(3, 1)
Call mysql.query("SELECT id FROM naryad_rodv_" & nowBase & " WHERE minor='" & minor + 1 & "' AND `major`='" & major & "'")
idnew = DAT(1, 1)
If st = 0 Then
    Exit Sub
Else
Call mysql.query("SELECT `narid` from `naryad_" & nowBase & "` WHERE `major`='" & major & "' and `minor`='" & minor & "'")
oldid() = DAT()
stt = st
Call mysql.query("SELECT `narid` from `naryad_" & nowBase & "` WHERE `major`='" & major & "' and `minor`='" & minor + 1 & "'")
newid() = DAT()
stt2 = st

Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set minor='" & minor & "' WHERE id='" & idnew & "'")
For x = 1 To stt2
Call mysql.query("UPDATE naryad_" & nowBase & " set minor='" & minor & "' WHERE `narid`='" & newid(1, x) & "'")
Next x

Call mysql.query("UPDATE naryad_rodv_" & nowBase & " set minor='" & minor + 1 & "' WHERE id='" & ID & "'")
For x = 1 To stt
Call mysql.query("UPDATE naryad_" & nowBase & " set minor='" & minor + 1 & "' WHERE `narid`='" & oldid(1, x) & "'")
Next x
End If
Call ListView1_Click
ListView2.FindItem(podrod).Selected = True
ListView2.FindItem(podrod).EnsureVisible
End Sub

