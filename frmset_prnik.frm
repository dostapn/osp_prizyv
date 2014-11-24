VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmset_print 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Личные номера"
      TabPicture(0)   =   "frmset_prnik.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "listln"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Автопечать"
      TabPicture(1)   =   "frmset_prnik.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check2"
      Tab(1).Control(1)=   "Check1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmset_prnik.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CheckBox Check2 
         Caption         =   "Авто-Печать"
         Height          =   375
         Left            =   -74400
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Скрытное формирование документов"
         Height          =   375
         Left            =   -74400
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin MSComctlLib.ListView listln 
         Height          =   6855
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   12091
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "asd"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmset_print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim lichnom() As String
Dim lichnomb() As String
Dim contt As Boolean
Dim lichnomdb() As String

Private Sub Check1_Click()
Call mysql.query("UPDATE bases set view_print='" & Check1.VAlue & "' WHERE VAl='" & nowBase & "'")
End Sub

Private Sub Check2_Click()
Call mysql.query("UPDATE bases set auto_print='" & Check2.VAlue & "' WHERE VAl='" & nowBase & "'")
End Sub

Private Sub Form_Load()
On Error Resume Next
'''''''''''''Личные номера
ReDim lichnomb(0)

Call sorting_blok("rodv", "naryad_" & nowBase)
For x = 1 To UBound(outm())
    Set LF = listln.ListItems.add(, , outm(x))
    
Next x



If Not get_config("print_lichnom") = "" Then
    lichnomdb = Split(get_config("print_lichnom"), ";")
    ReDim lichnom(UBound(lichnomdb()))
        For x = 0 To UBound(lichnomdb())
           lichnom(x) = lichnomdb(x)
           If Not lichnom(x) = "" Then listln.FindItem(lichnom(x)).Checked = True
        Next x
End If
''''''''''''''
''''''''Авто-печать
Check1 = Int(get_config("view_print"))
Check2 = Int(get_config("auto_print"))
''''''''



End Sub


Private Sub Form_Unload(Cancel As Integer)
ReDim lichnom(0)
End Sub

Private Sub listln_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    
    Dim strsql As String
    Dim stt As Long
    
    If Item.Checked = True Then
        For x = 0 To UBound(lichnom)
            If Not lichnom(x) = Item Then contt = True
        Next x
                If contt = True Then
                    If lichnom(0) = "" Then
                        ReDim lichnomb(UBound(lichnom))
                    Else
                        ReDim lichnomb(UBound(lichnom) + 1)
                    End If
                    
                    For Y = 0 To UBound(lichnom())
                        lichnomb(Y) = lichnom(Y)
                    Next Y
            
                    lichnomb(UBound(lichnomb)) = Item
                    ReDim lichnom(UBound(lichnomb))
                    lichnom() = lichnomb()
                End If
        
    Else
        For x = 0 To UBound(lichnom())
            If lichnom(x) = Item Then lichnom(x) = vbNullString
        Next x
    End If
   

For x = 1 To UBound(lichnom())
    If Not lichnom(x) = "" Then stt = stt + 1
Next x
ReDim lichnomb(stt)
stt = 0
For x = 0 To UBound(lichnom())
        
        If Not lichnom(x) = "" Then
            lichnomb(stt) = lichnom(x)
            stt = stt + 1
        End If
Next x
ReDim lichnom(stt)
lichnom() = lichnomb()

For x = 0 To UBound(lichnom())
    If x = 0 Then
        strsql = strsql & lichnom(x)
    Else
         strsql = strsql & ";" & lichnom(x)
    End If
Next x
Call mysql.query("UPDATE bases set print_lichnom='" & strsql & "' WHERE VAl='" & nowBase & "'")
End Sub
