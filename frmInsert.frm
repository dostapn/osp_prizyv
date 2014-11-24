VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInsert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Добавить команду"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListComm 
      Height          =   3375
      Left            =   45
      TabIndex        =   8
      Top             =   1395
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Команда"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Предназначение"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   345
      Left            =   5715
      TabIndex        =   7
      Top             =   855
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   5715
      TabIndex        =   6
      Top             =   405
      Width           =   1035
   End
   Begin VB.TextBox txtForCh 
      Height          =   285
      Left            =   2175
      TabIndex        =   5
      Top             =   960
      Width           =   2505
   End
   Begin VB.TextBox txtForP 
      Height          =   285
      Left            =   2175
      TabIndex        =   4
      Top             =   630
      Width           =   2505
   End
   Begin VB.TextBox txtOblKom 
      Height          =   285
      Left            =   2175
      TabIndex        =   0
      Top             =   300
      Width           =   2505
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Для части:"
      Height          =   195
      Left            =   1230
      TabIndex        =   3
      Top             =   990
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Для пункта:"
      Height          =   195
      Left            =   1125
      TabIndex        =   2
      Top             =   660
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Областная команда:"
      Height          =   195
      Left            =   495
      TabIndex        =   1
      Top             =   330
      Width           =   1575
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FOR_P As String
Dim FOR_CH As String
Dim CLICK_COM As String

Private Sub asxButtonStrip1_ButtonClick(ByVal PageKey As String, ByVal PageIndex As Integer, ByVal ButtonKey As String, ByVal ButtonIndex As Integer)

End Sub

Private Sub cmdCancel_Click()
      Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    ListComm_DblClick
    Call frmCommand.commANDs_load
    frmCommand.listcommands.FindItem(CLICK_COM).Selected = True
    frmCommand.listcommands.FindItem(CLICK_COM).EnsureVisible
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call frmCommAND.commANDs_load
End Sub

'Base_fOR_OSP.

Private Sub ListComm_DblClick()
        On Error Resume Next

        If ListComm.ListItems.Count > 0 Then
            Call AddComm
        End If

End Sub

Private Sub ListComm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If ListComm.ListItems.Count > 0 Then
        Call AddComm
    Else
        Call ShowLike
    End If
End If
End Sub


Private Sub txtOblKom_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then
            ShowLike
        End If
End Sub

Sub AddComm()
On Error Resume Next


            
            Dim indSel As Long
            If ListComm.ListItems.Count = 0 Then Exit Sub
            CLICK_COM = ListComm.ListItems(ListComm.SelectedItem.Index).Text
            Dim new_id As Long
            Dim NEW_NAR_ID As Long
            If Not mysql.query("SELECT max(otpravkaid) FROM otpravka_" & nowBase & "") Then
                MsgBox "Connect ErrOR", vbCritical, strMAIN_TITLE
                Exit Sub
            End If
            
            If Not st = 0 Then new_id = DAT(1, st) + 1 Else new_id = 1
            
            Call mysql.query("SELECT narid, oblkom FROM naryad_" & nowBase & " WHERE oblkom = '" & CLICK_COM & "'")
            For x = 1 To st
            If InStr(1, DAT(2, x), CLICK_COM, vbBinaryCompare) > 0 Then NEW_NAR_ID = CLng(DAT(1, x)): GoTo Cont
            Next x
            MsgBox "Ooops!", vbCritical, "STOP": Exit Sub
            
Cont:
            If new_id = 0 Then new_id = 1
            Call mysql.query("INSERT INTO `otpravka_" & nowBase & "` ( `otpravkaid` , `data` , `narid` , `fORpunkt` , `fORchast` , `kolvo` )VALUES ('" & new_id & "','" & dateotp & "','" & NEW_NAR_ID & "','" & txtForP & "','" & txtForCh & "','0');")
    
           
         
       

            Call log_sql("1", "7", CLICK_COM, "")
ret:
            DoEvents
        Call frmCommand.commANDs_load
        frmCommand.listcommands.FindItem(CLICK_COM).Selected = True
        frmCommand.listcommands.FindItem(CLICK_COM).EnsureVisible
        Call frmCommand.commANDs_load
        frmCommand.listprnik.Refresh
        Unload Me



End Sub

Sub ShowLike()
        On Error Resume Next
        FOR_P = txtForP
        FOR_CH = txtForCh
        ListComm.ListItems.Clear
        Call mysql.query("SELECT oblkom, vrp FROM naryad_" & nowBase & " WHERE oblkom like '" & txtoblkom & "%'")
        For x = 1 To st
             Set LF = ListComm.ListItems.add(, , DAT(1, x))
             LF.SubItems(1) = DAT(2, x)
        Next x
        ListComm.SetFocus
End Sub
