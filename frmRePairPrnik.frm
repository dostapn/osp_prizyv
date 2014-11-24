VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRePairPrnik 
   Caption         =   "Восстановить призывника"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7995
   Icon            =   "frmRePairPrnik.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Развернуто
   Begin MSComctlLib.ImageList ImageKilled 
      Left            =   540
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRePairPrnik.frx":0CCA
            Key             =   "KILL"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRepair 
      Caption         =   "Восстановить"
      Default         =   -1  'True
      Height          =   375
      Left            =   5220
      TabIndex        =   6
      Top             =   6420
      Width           =   1290
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   6615
      TabIndex        =   5
      Top             =   6420
      Width           =   1290
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Привязать вниз
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   7035
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7488
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAddPrnik 
      Align           =   1  'Привязать вверх
      Appearance      =   0  'Плоска
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Примечание: при восстановлении номер учетно-послужной карточки остается прежним."
         Height          =   195
         Left            =   525
         TabIndex        =   7
         Top             =   795
         Width           =   6780
      End
      Begin VB.Label lblVod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Для восстановления призывника выделите его и нажмите кнопку ""Восстановить"""
         Height          =   195
         Left            =   525
         TabIndex        =   2
         Top             =   555
         Width           =   6345
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "В списке перечислены удаленные из базы призывники"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   375
         TabIndex        =   1
         Top             =   300
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7425
         Picture         =   "frmRePairPrnik.frx":0FE4
         Top             =   210
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView ListFiles 
      Height          =   4965
      Left            =   15
      TabIndex        =   3
      Top             =   1170
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8758
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "УПК"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ф.И.О."
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Военкомат"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Удален"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Удален пользователем"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Примечание"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Список"
      Begin VB.Menu mnuToExcel 
         Caption         =   "В Excel"
         Shortcut        =   +{F12}
      End
   End
End
Attribute VB_Name = "frmRePairPrnik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRepair_Click()
    Call ListFiles_DblClick
End Sub

Private Sub Form_Load()
    Call Load_Column_Width
    Call RefList
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    listfiles.Move 0, picAddPrnik.Height, ScaleWidth, ScaleHeight - listfiles.Top - StatusBar1.Height - 500
    cmdClose.Move ScaleWidth - 2700, ScaleHeight - 700
    cmdRepair.Move ScaleWidth - 2600 + cmdRepair.Width, ScaleHeight - 700
    Image1.Left = ScaleWidth - Image1.Width - 100
End Sub

Private Sub Form_Unload(cancel As Integer)
    Call Save_Column_Width
End Sub
Function Save_Column_Width()
On Error Resume Next
Dim w As Long
    For w = 1 To listfiles.ColumnHeaders.Count
        Call SaveSetting("RNB_PRIZ_DB", "SETTING_RepaikPrnik", "Wid" & w, listfiles.ColumnHeaders.Item(w).Width)
    Next w
End Function

Function Load_Column_Width()
On Error Resume Next
Dim w As Long
    For w = 1 To listfiles.ColumnHeaders.Count
         listfiles.ColumnHeaders.Item(w).Width = GetSetting("RNB_PRIZ_DB", "SETTING_RepaikPrnik", "Wid" & w)
    Next w
End Function

Private Sub ListFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    
    listfiles.Sorted = True
    
    If listfiles.SortKey = ColumnHeader.Index - 1 Then
        If listfiles.SortOrder = lvwDescending Then listfiles.SortOrder = lvwAscending Else listfiles.SortOrder = lvwDescending
    Else
        listfiles.SortOrder = lvwAscending
        listfiles.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Sub RefList()
On Error Resume Next
Dim X As Long
Dim allp As Long
listfiles.ListItems.Clear
    Dim dat_o() As String
    Call mysql.query("select * from `delprnik_" & nowBase & "`")
    dat_o() = DAT()
    For X = 1 To st
                Set LF = listfiles.ListItems.add(, , dat_o(1, X))
                LF.SubItems(1) = dat_o(3, X) & " " & dat_o(4, X) & " " & dat_o(5, X)
                LF.SubItems(2) = dat_o(2, X)
                LF.SubItems(3) = dat_o(26, X)
                'get FIO
                Call mysql.query("select `fio` from users where `name`='" & dat_o(27, X) & "'")
                LF.SubItems(4) = DAT(1, 1)
                LF.SubItems(5) = dat_o(28, X)
                
     Next X
StatusBar1.Panels(1) = "Всего: " & listfiles.ListItems.Count

End Sub


Private Sub ListFiles_DblClick()
On Error Resume Next
Dim oldUPK As Long
Call MessageBeep(vbInformation)
If MsgBox("Восстановить призывника?", vbYesNo + vbQuestion, strMAIN_TITLE) = vbYes Then
    If listfiles.ListItems.Count = 0 Then MessageBeep (vbCritical): Exit Sub
          If Len(CRITICAL_OPER) > 0 Then MsgBox "Идет " & CRITICAL_OPER & ". Повторите попытку позже.", vbExclamation, "Процесс": Exit Sub
        oldUPK = Int(listfiles.ListItems(listfiles.SelectedItem.Index).Text)
        listfiles.ListItems.Remove (listfiles.SelectedItem.Index)
        Call MovePrnikToBase(oldUPK)
End If
    Call RefList
End Sub

Private Sub mnuToExcel_Click()
Call Cnv.ResoultSearch(listfiles, True, "ResoultDeleted", "Удаленные призывники")
End Sub
